"""item 物理行实时定位回归测试：SAP 按 POSNR 重排后仍要写对行。

现象（用户 2026-09-02 实机）：SAP 已有 item 1000/1001/3000/5000，编辑时新增 2000，
回车后 SAP **强制把 2000 排到第 3 行**（物理 row 2），3000/5000 顺延。旧实现用重排前的
物理行号做身份排他（`used_rows`），把正确行误判为已占用 → 退回写入行 → 2000 的金额
写进了 5000。

本模块的 fake session 是关键：它**真的模拟 POSNR 重排与 item 详情路由**，
不模拟就复现不出该缺陷（既有测试的 fake 是纯 KV 缓存、send_vkey 空实现，故一直漏过）。
"""

from __future__ import annotations

import os
import re
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

import sap.test.helpers  # noqa: E402,F401  仅为注入 win32com mock
from sap.models import (  # noqa: E402
    OrderData,
    OrderItemData,
    PlanCostEntry,
    RevenueData,
    SapConfig,
    SapResult,
)
from sap.session import SapSession  # noqa: E402
from sap.transactions.order import OrderTransaction  # noqa: E402
from sap.transactions.order_edit import OrderEditTransaction  # noqa: E402

# item 概览表格单元格 ID 的列/行解析（列序号见 OrderTransaction._item_id 等）。
_CELL_RE = re.compile(
    r"tblSAPMV45ATCTRL_U_ERF_GUTLAST/"
    r"(txtVBAP-POSNR|ctxtRV45A-MABNR|txtVBAP-ZMENG|ctxtVBAP-ZIEME|txtVBAP-NETWR)\[\d+,(\d+)\]$"
)
_COLUMN_FIELD = {
    "txtVBAP-POSNR": "no",
    "ctxtRV45A-MABNR": "material",
    "txtVBAP-ZMENG": "quantity",
    "ctxtVBAP-ZIEME": "unit",
    "txtVBAP-NETWR": "net",
}
# 价格条件格：创建/新增走第 5 行，编辑已存在 item 走第 1 行；两者都只作用于"当前详情 item"。
_CONDITION_IDS = frozenset({
    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/"
    "tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]",
    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/"
    "tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,1]",
})
_PLAN_COST_MENU = "wnd[0]/mbar/menu[3]/menu[7]"
_BACK_BUTTON = "wnd[0]/tbar[0]/btn[3]"


class _Item:
    """fake SAP 侧的一条 item：概览列 + 详情条件价。"""

    def __init__(self, no: str = "", material: str = "", condition: str = "0.00"):
        self.no = no
        self.material = material
        self.quantity = ""
        self.unit = ""
        self.condition = condition
        self.net = condition


class _Cell:
    """概览表格单元格：读写直接落到 _Item 上；越界行为空且写入即建行。"""

    def __init__(self, session: "_ReorderingSession", row: int, field: str):
        self._session = session
        self._row = row
        self._field = field
        self.caretPosition = 0

    @property
    def text(self) -> str:
        item = self._session.item_at(self._row)
        return "" if item is None else getattr(item, self._field)

    @text.setter
    def text(self, value):
        setattr(self._session.ensure_item_at(self._row), self._field, str(value))

    def setFocus(self):
        self._session.focused_row = self._row

    def sendVKey(self, key):
        pass


class _Condition:
    """价格条件格：读写落到"当前进入详情的 item"，进错行就会写错 item——正是要验证的点。"""

    def __init__(self, session: "_ReorderingSession"):
        self._session = session
        self.caretPosition = 0

    @property
    def text(self) -> str:
        item = self._session.current_item
        return "" if item is None else item.condition

    @text.setter
    def text(self, value):
        if self._session.current_item is not None:
            self._session.current_item.condition = str(value)

    def setFocus(self):
        pass

    def sendVKey(self, key):
        pass


class _Generic:
    """其余控件：普通可读写元素，附带 select/press 记录供断言。"""

    def __init__(self, session: "_ReorderingSession", element_id: str):
        self._session = session
        self._id = element_id
        self.text = ""
        self.key = ""
        self.caretPosition = 0

    def setFocus(self):
        pass

    def sendVKey(self, key):
        self._session.window_vkey(key)

    def press(self):
        if self._id == _BACK_BUTTON:
            self._session.leave_detail()

    def select(self):
        if self._id == _PLAN_COST_MENU:
            # 记录打开计划成本编辑器时的焦点行，用于断言"按 item 号定位"是否正确。
            self._session.plan_cost_focus_row = self._session.focused_row

    def setSelectionIndexes(self, start, end):
        pass


class _ReorderingSession:
    """会按 POSNR 重排的 fake SAP session。

    行为约定（对齐实机）：
        - 概览页 send_vkey(0) → 给未编号的行分配新号，再**按 POSNR 升序整体重排**；
        - 概览页 send_vkey(2) → 以当前焦点行进入该 item 详情（后续条件写入只作用于它）；
        - btn[3] → 从详情返回概览，并把条件价同步到该 item 的净值列。
    """

    def __init__(self, items: list[_Item] | None = None, number_step: int = 10):
        self.items = items or []
        self.current_item: _Item | None = None
        self.focused_row = -1
        self.plan_cost_focus_row: int | None = None
        self._number_step = number_step
        self._generic: dict[str, _Generic] = {}

    # -------------------------- 表格模型 -------------------------- #
    def item_at(self, row: int) -> _Item | None:
        return self.items[row] if 0 <= row < len(self.items) else None

    def ensure_item_at(self, row: int) -> _Item:
        while len(self.items) <= row:
            self.items.append(_Item())
        return self.items[row]

    def item_nos(self) -> list[str]:
        return [item.no for item in self.items]

    # -------------------------- 事件 -------------------------- #
    def window_vkey(self, key: int) -> None:
        if self.current_item is not None:
            # 详情屏内的回车只是提交字段，不触发概览重排。
            return
        if key == 0:
            self._commit_and_reorder()
        elif key == 2:
            self.current_item = self.item_at(self.focused_row)

    def leave_detail(self) -> None:
        if self.current_item is not None:
            # 数量恒为 1，净值即条件价。
            self.current_item.net = self.current_item.condition
            self.current_item = None

    def _commit_and_reorder(self) -> None:
        numbered = [int(item.no) for item in self.items if item.no.isdigit()]
        next_no = (max(numbered) if numbered else 0) + self._number_step
        for item in self.items:
            if not item.no and item.material:
                item.no = str(next_no)
                next_no += self._number_step
        # SAP 强制按 POSNR 升序重排——物理行号自此与写入顺序无关。
        self.items = [item for item in self.items if item.no or item.material]
        self.items.sort(key=lambda item: int(item.no) if item.no.isdigit() else 10**9)

    # -------------------------- COM 入口 -------------------------- #
    def findById(self, element_id: str):
        match = _CELL_RE.search(element_id)
        if match:
            return _Cell(self, int(match.group(2)), _COLUMN_FIELD[match.group(1)])
        if element_id in _CONDITION_IDS:
            return _Condition(self)
        if element_id not in self._generic:
            self._generic[element_id] = _Generic(self, element_id)
        return self._generic[element_id]


def _config() -> SapConfig:
    return SapConfig(
        order_type="ZOR", sales_organization="3002", distribution_channels="10",
        sales_office="1000", sales_group="200", sub_cost_center_cs="1101",
        sub_cost_center_chm="1102", sub_cost_center_phy="1103",
        cs_code="CS001", sales_code="SA001",
    )


def _make(items: list[_Item]) -> tuple[OrderEditTransaction, OrderTransaction, _ReorderingSession]:
    raw = _ReorderingSession(items)
    session = SapSession(raw, raw, raw, raw)
    base = OrderTransaction(session, _config())
    return OrderEditTransaction(session, _config()), base, raw


def _order(*items: OrderItemData) -> OrderData:
    return OrderData(
        sap_no="123456", project_no="PRJ-001", currency_type="CNY",
        exchange_rate=1.0, short_text="T", items=list(items),
    )


def _condition_of(raw: _ReorderingSession, item_no: str) -> str:
    return next(item.condition for item in raw.items if item.no == item_no)


class EditItemsReorderTest(unittest.TestCase):
    """编辑：新增 item 触发重排后，金额必须落在新 item 上。"""

    def test_insert_middle_item_writes_amount_to_itself(self):
        # 用户报的场景：已有 1000/1001/3000/5000，新增 2000 → SAP 把它排到物理行 2。
        raw_items = [
            _Item("1000", "M1", "100.00"), _Item("1001", "M2", "200.00"),
            _Item("3000", "M3", "300.00"), _Item("5000", "M5", "500.00"),
        ]
        tx, _base, raw = _make(raw_items)
        order = _order(
            OrderItemData(item="1000", material_code="M1", revenue=100.0),
            OrderItemData(item="1001", material_code="M2", revenue=200.0),
            OrderItemData(item="2000", material_code="M2000", revenue=222.0),
            OrderItemData(item="3000", material_code="M3", revenue=300.0),
            OrderItemData(item="5000", material_code="M5", revenue=500.0),
        )
        diffs: list[str] = []

        result = tx.edit_items(order, diffs)

        self.assertTrue(result.success, result.message)
        # 重排后的物理顺序：2000 确实被 SAP 插到第 3 行。
        self.assertEqual(raw.item_nos(), ["1000", "1001", "2000", "3000", "5000"])
        # 核心断言：金额写在 2000 自己身上，5000 分文未动（旧实现此处被写成 222.00）。
        self.assertEqual(_condition_of(raw, "2000"), "222.00")
        self.assertEqual(_condition_of(raw, "5000"), "500.00")
        self.assertEqual(_condition_of(raw, "3000"), "300.00")

    def test_two_inserts_each_land_on_own_row(self):
        # 连续新增 2000 与 4000，两次重排都要各归各位。
        raw_items = [_Item("1000", "M1", "100.00"), _Item("5000", "M5", "500.00")]
        tx, _base, raw = _make(raw_items)
        order = _order(
            OrderItemData(item="1000", material_code="M1", revenue=100.0),
            OrderItemData(item="2000", material_code="M2000", revenue=222.0),
            OrderItemData(item="4000", material_code="M4000", revenue=444.0),
            OrderItemData(item="5000", material_code="M5", revenue=500.0),
        )
        diffs: list[str] = []

        result = tx.edit_items(order, diffs)

        self.assertTrue(result.success, result.message)
        self.assertEqual(raw.item_nos(), ["1000", "2000", "4000", "5000"])
        self.assertEqual(_condition_of(raw, "2000"), "222.00")
        self.assertEqual(_condition_of(raw, "4000"), "444.00")
        self.assertEqual(_condition_of(raw, "5000"), "500.00")

    def test_auto_numbered_insert_located_by_new_number(self):
        # ODM item 号已被占用 → 不写 POSNR、由 SAP 自动分配；按"写入前号集合"的差集定位。
        raw_items = [_Item("1000", "M1", "100.00"), _Item("2000", "M2", "200.00")]
        tx, _base, raw = _make(raw_items)
        order = _order(
            OrderItemData(item="1000", material_code="M1", revenue=100.0),
            # item 号 1000 已存在但物料不同 → 走新增，SAP 自动改号。
            OrderItemData(item="1000", material_code="MX", revenue=777.0),
        )
        diffs: list[str] = []

        result = tx.edit_items(order, diffs)

        self.assertTrue(result.success, result.message)
        new_item = next(item for item in raw.items if item.material == "MX")
        self.assertEqual(new_item.condition, "777.00")        # 金额落在新行
        self.assertNotIn(new_item.no, {"1000", "2000"})       # 号由 SAP 重新分配
        self.assertEqual(_condition_of(raw, "2000"), "200.00")  # 邻行未被波及

    def test_unlocatable_new_item_skips_write_and_warns(self):
        # 新增行被 SAP 丢弃（模拟：写入后行消失）→ 定位不到，宁可不写也不写到别的 item 上。
        raw_items = [_Item("1000", "M1", "100.00"), _Item("5000", "M5", "500.00")]
        tx, _base, raw = _make(raw_items)

        original_commit = raw._commit_and_reorder

        def _drop_new_row():
            original_commit()
            raw.items = [item for item in raw.items if item.material != "M2000"]

        raw._commit_and_reorder = _drop_new_row
        order = _order(
            OrderItemData(item="1000", material_code="M1", revenue=100.0),
            OrderItemData(item="2000", material_code="M2000", revenue=222.0),
            OrderItemData(item="5000", material_code="M5", revenue=500.0),
        )
        diffs: list[str] = []

        result = tx.edit_items(order, diffs)

        self.assertTrue(result.success, result.message)
        self.assertTrue(result.warning)
        self.assertTrue(any("无法定位物理行" in d for d in diffs), diffs)
        # 其余 item 的金额一个都不能被这条写坏。
        self.assertEqual(_condition_of(raw, "1000"), "100.00")
        self.assertEqual(_condition_of(raw, "5000"), "500.00")


class CreateItemsReorderTest(unittest.TestCase):
    """创建：Excel 顺序与 SAP 重排后的物理行不一致时，金额仍要各归各位。"""

    def test_write_item_rows_locates_after_reorder(self):
        # Excel 给的顺序是 2000 在前、1000 在后；SAP 回车后按 POSNR 重排 → 顺序反转。
        _tx, base, raw = _make([])
        order = _order(
            OrderItemData(item="2000", material_code="M2000", revenue=222.0, quantity="1", unit="pu"),
            OrderItemData(item="1000", material_code="M1000", revenue=111.0, quantity="1", unit="pu"),
        )

        base._write_item_rows(order, SapResult(step="test"))

        self.assertEqual(raw.item_nos(), ["1000", "2000"])
        # 旧实现按列表索引写：item[0](2000) 会被写到物理行 0（即 1000）上。
        self.assertEqual(_condition_of(raw, "1000"), "111.00")
        self.assertEqual(_condition_of(raw, "2000"), "222.00")

    def test_add_items_reports_amount_sum(self):
        # 端到端：add_items 走完后未税加和取两条之和，且各自金额没写串。
        _tx, base, raw = _make([])
        order = _order(
            OrderItemData(item="3000", material_code="M3", revenue=300.0, quantity="1", unit="pu"),
            OrderItemData(item="1000", material_code="M1", revenue=100.0, quantity="1", unit="pu"),
        )

        result = base.add_items(order, RevenueData(revenue=400.0, revenue_cny=400.0))

        self.assertTrue(result.success, result.message)
        self.assertEqual(result.sap_amount_vat, "400.00")
        self.assertEqual(_condition_of(raw, "1000"), "100.00")
        self.assertEqual(_condition_of(raw, "3000"), "300.00")


class PlanCostLocateTest(unittest.TestCase):
    """计划成本：按 item 号实时定位物理行，而非调用方的列表索引。"""

    def test_create_path_opens_editor_on_located_row(self):
        # SAP 物理顺序 1000/2000/3000；target_item=3000 应聚焦物理行 2，
        # 而调用方传的兜底 focus_row=0 必须被忽略。
        _tx, base, raw = _make([
            _Item("1000", "M1", "100.00"), _Item("2000", "M2", "200.00"), _Item("3000", "M3", "300.00"),
        ])

        result = base.apply_plan_cost_entries(
            [PlanCostEntry(cost_center="1100", category="FREMDL", amount=50.0)],
            focus_row=0, target_item="3000",
        )

        self.assertTrue(result.success, result.message)
        self.assertEqual(raw.plan_cost_focus_row, 2)

    def test_create_path_skips_when_item_absent(self):
        # SAP 无该 item → 不开编辑器、不失败，但标 warning 让 UI 提示。
        _tx, base, raw = _make([_Item("1000", "M1", "100.00")])

        result = base.apply_plan_cost_entries(
            [PlanCostEntry(cost_center="1100", category="FREMDL", amount=50.0)],
            focus_row=0, target_item="9000",
        )

        self.assertTrue(result.success, result.message)
        self.assertTrue(result.warning)
        self.assertIsNone(raw.plan_cost_focus_row)
        self.assertIn("9000", result.message)


if __name__ == "__main__":
    unittest.main()
