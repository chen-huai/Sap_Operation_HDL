"""VA02 编辑：Data B 一致性短路 + 删空/重建（三段式）的回归测试。

覆盖：①内容一致（含强制行）→ 判一致；②SAP 行被重排 → 仍判一致（多重集语义）；
③金额/行数/成本中心/item 任一不同 → 判不一致；④sales_group=240 不比 item；
⑤clear_data_b 判一致时零删除动作；⑥判不一致时先删配对行、后删强制行，各自 bottom-up。

背景约束（见 .claude/plan/data_b_clear_save_rewrite.md）：
删除与写入之间必须隔一次 save+open_order，否则 SAP 报 ZR520；强制成本中心行只在执行部门表
有行、成本表无行，故删除必须两阶段。费率成本中心列不可读，比对时由"费率==执行部门"不变量覆盖。
"""

from __future__ import annotations

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

import sap.test.helpers  # noqa: E402,F401  仅为注入 win32com mock
from sap.models import DataBEntry, OrderData, SapConfig  # noqa: E402
from sap.session import SapSession  # noqa: E402
from sap.transactions.order_edit import OrderEditTransaction  # noqa: E402

BASE = "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312"
ZUL = f"{BASE}/tblSAPMV45AZULEISTENDE"
KOS = f"{BASE}/tblSAPMV45AKOSTENSAETZE"
DELETE_BTN = f"{BASE}/btnTABLOESCH"


def _kostl(row): return f"{ZUL}/ctxtTABL-KOSTL[0,{row}]"
def _zposition(row): return f"{ZUL}/txtTABL-ZPOSITION[1,{row}]"
def _posnr(row): return f"{KOS}/txtTABD-POSNR[1,{row}]"
def _festpreis(row): return f"{KOS}/txtTABD-FESTPREIS[5,{row}]"


class _Row:
    """table control 的行对象，仅需支持 selected 赋值。"""

    def __init__(self, table, index):
        self._table = table
        self._index = index

    @property
    def selected(self): return self._index in self._table.selected_rows

    @selected.setter
    def selected(self, value):
        if value:
            self._table.selected_rows.append(self._index)


class _Element:
    def __init__(self, text: str = ""):
        self.text = text
        self.key = ""
        self.caretPosition = 0
        self.presses = 0
        self.selected_rows: list[int] = []

    def setFocus(self): pass
    def sendVKey(self, k): pass
    def press(self): self.presses += 1
    def select(self): pass
    def getAbsoluteRow(self, row): return _Row(self, row)


class _RawSession:
    """按 id 缓存控件；未预设的 id 自动生成空控件（等价于 SAP 空行）。"""

    def __init__(self, preset=None):
        self._cache = dict(preset or {})
        self.press_log: list[str] = []

    def findById(self, element_id: str):
        if element_id not in self._cache:
            self._cache[element_id] = _Element()
        element = self._cache[element_id]
        if element_id == DELETE_BTN:
            # 删除按钮按下时记录当前选中行，供断言删除顺序与选表范围。
            self.press_log.append(
                f"del zul={self._selected(ZUL)} kos={self._selected(KOS)}"
            )
            self._apply_delete()
            self._selected_reset()
        return element

    def _selected(self, table_id):
        table = self._cache.get(table_id)
        return list(table.selected_rows) if table else []

    def _selected_reset(self):
        for table_id in (ZUL, KOS):
            table = self._cache.get(table_id)
            if table:
                table.selected_rows = []

    def _apply_delete(self):
        """按选中行真删：被删行之后的内容整体上移一行，与 SAP 表格删除行为一致。

        不模拟真删，阶段二重数剩余行会永远读到原始行数，测试就无法覆盖两阶段删除的真实行数变化。
        """
        zul_selected = self._selected(ZUL)
        kos_selected = self._selected(KOS)
        if zul_selected:
            self._shift_up((_kostl, _zposition), zul_selected[0])
        if kos_selected:
            self._shift_up((_posnr, _festpreis), kos_selected[0])

    def _shift_up(self, id_funcs, start, max_rows=50):
        for row in range(start, max_rows - 1):
            for id_func in id_funcs:
                nxt = self._cache.get(id_func(row + 1))
                self._cache.setdefault(id_func(row), _Element()).text = nxt.text if nxt else ""


def _sap_rows(rows, forced=()):
    """构造 SAP 现值：rows=[(成本中心, item, 金额)] 配对行，forced=[成本中心] 只在执行部门表。"""
    preset = {ZUL: _Element(), KOS: _Element()}
    for row, (cost_center, item, amount) in enumerate(rows):
        preset[_kostl(row)] = _Element(cost_center)
        preset[_zposition(row)] = _Element(item)
        preset[_posnr(row)] = _Element(item)
        preset[_festpreis(row)] = _Element(amount)
    for offset, cost_center in enumerate(forced):
        preset[_kostl(len(rows) + offset)] = _Element(cost_center)
    return preset


def _make_tx(preset=None):
    raw = _RawSession(preset)
    session = SapSession(raw, raw, raw, raw)
    config = SapConfig(
        order_type="ZOR", sales_organization="3002", distribution_channels="10",
        sales_office="1000", sales_group="200", sub_cost_center_cs="1101",
        sub_cost_center_chm="1102", sub_cost_center_phy="1103", cs_code="CS001", sales_code="SA001",
    )
    return OrderEditTransaction(session, config), raw


def _order(sales_group="200"):
    return OrderData(
        sap_no="10000", project_no="P1", currency_type="CNY", exchange_rate=1.0,
        short_text="T", sales_group=sales_group,
    )


def _entry(cost_center, amount, item="", forced=False):
    return DataBEntry(
        performer_cost_center=cost_center,
        rate_cost_center="" if forced else cost_center,
        amount=amount,
        item=item,
        kostl_only=forced,
    )


# 标准场景：Excel 正常行 48601293 / 48601300，config 强制行 48601258。
ENTRIES = [
    _entry("48601293", 1000.0, "1000"),
    _entry("48601300", 250.5, "2000"),
    _entry("48601258", 0.0, forced=True),
]
SAP_ROWS = [("48601293", "1000", "1000.00"), ("48601300", "2000", "250.50")]
FORCED = ["48601258"]


class DataBMatchTest(unittest.TestCase):
    """_data_b_diff 判定口径（返回空串=一致）。"""

    def _diff(self, rows, forced, entries=None, sales_group="200"):
        tx, _ = _make_tx(_sap_rows(rows, forced))
        zul, kos, truncated = tx._read_data_b_snapshot(BASE)
        self.assertFalse(truncated)
        # entries 可为空列表（Excel Data B 全空场景），不能用 `or` 兜底。
        return tx._data_b_diff(
            zul, kos, ENTRIES if entries is None else entries, _order(sales_group)
        )

    def _matches(self, rows, forced, entries=None, sales_group="200"):
        return not self._diff(rows, forced, entries, sales_group)

    def test_identical_including_forced_row(self):
        self.assertTrue(self._matches(SAP_ROWS, FORCED))

    def test_reordered_rows_still_match(self):
        # SAP 按成本中心号重排：强制行 48601258 排到最前、正常行顺序颠倒，
        # 且执行部门表 3 行 / 成本表 2 行——强制行不在末尾也须判一致（多重集语义）。
        rows = [("48601300", "2000", "250.50"), ("48601293", "1000", "1000.00")]
        preset = _sap_rows(rows, [])
        preset[_kostl(0)] = _Element("48601258")
        preset[_zposition(0)] = _Element("")
        preset[_kostl(1)] = _Element("48601300")
        preset[_zposition(1)] = _Element("2000")
        preset[_kostl(2)] = _Element("48601293")
        preset[_zposition(2)] = _Element("1000")
        tx, _ = _make_tx(preset)
        zul, kos, _t = tx._read_data_b_snapshot(BASE)
        self.assertEqual(tx._data_b_diff(zul, kos, ENTRIES, _order()), "")

    def test_amount_diff_not_match(self):
        rows = [("48601293", "1000", "1000.00"), ("48601300", "2000", "250.51")]
        self.assertFalse(self._matches(rows, FORCED))

    def test_thousand_separator_still_match(self):
        # SAP 读回带千分位，归一后与 Excel 一致，不应误判为差异。
        entries = [_entry("48601293", 12345.6, "1000"), _entry("48601258", 0.0, forced=True)]
        rows = [("48601293", "1000", "12,345.60")]
        self.assertTrue(self._matches(rows, FORCED, entries))

    def test_missing_forced_row_not_match(self):
        self.assertFalse(self._matches(SAP_ROWS, []))

    def test_extra_sap_row_not_match(self):
        rows = SAP_ROWS + [("48601999", "3000", "10.00")]
        self.assertFalse(self._matches(rows, FORCED))

    def test_cost_center_diff_not_match(self):
        rows = [("48601293", "1000", "1000.00"), ("48601301", "2000", "250.50")]
        self.assertFalse(self._matches(rows, FORCED))

    def test_item_diff_not_match(self):
        rows = [("48601293", "1000", "1000.00"), ("48601300", "9000", "250.50")]
        self.assertFalse(self._matches(rows, FORCED))

    def test_sales_group_240_ignores_item(self):
        # 240 订单本就不写 item 号，SAP 侧 item 全空不应判为差异。
        rows = [("48601293", "", "1000.00"), ("48601300", "", "250.50")]
        self.assertTrue(self._matches(rows, FORCED, sales_group="240"))

    def test_excel_empty_vs_sap_rows_not_match(self):
        self.assertFalse(self._matches(SAP_ROWS, FORCED, entries=[]))

    def test_both_empty_match(self):
        self.assertTrue(self._matches([], [], entries=[]))

    def test_diff_text_carries_both_sides(self):
        # 差异描述须带两侧实际值，实测时能一眼区分"回读格式不同"与"真实数据不同"。
        rows = [("48601293", "1000", "1000.00"), ("48601300", "2000", "250.51")]
        desc = self._diff(rows, FORCED)
        self.assertIn("金额不同", desc)
        self.assertIn("250.51", desc)
        self.assertIn("250.50", desc)

    def test_diff_text_reports_item_mismatch_sides(self):
        # SAP item 带前导零这类格式差异，日志里能直接看出两侧形态。
        rows = [("48601293", "001000", "1000.00"), ("48601300", "2000", "250.50")]
        desc = self._diff(rows, FORCED)
        self.assertIn("item 不同", desc)
        self.assertIn("001000", desc)


class ClearDataBTest(unittest.TestCase):
    """clear_data_b 的短路与两阶段删除行为。"""

    def test_no_delete_when_identical(self):
        tx, raw = _make_tx(_sap_rows(SAP_ROWS, FORCED))
        diffs: list[str] = []
        result = tx.clear_data_b(ENTRIES, _order(), diffs)
        self.assertTrue(result.success)
        self.assertFalse(result.changed)
        self.assertTrue(result.warning)
        self.assertEqual(raw.press_log, [])
        self.assertEqual(diffs, [])

    def test_diff_desc_written_to_diffs(self):
        rows = [("48601293", "1000", "1000.00"), ("48601300", "2000", "999.00")]
        tx, _raw = _make_tx(_sap_rows(rows, FORCED))
        diffs: list[str] = []
        tx.clear_data_b(ENTRIES, _order(), diffs)
        self.assertIn("金额不同", diffs[0])
        self.assertIn("已删除旧 3 行", diffs[1])

    def test_two_phase_delete_when_different(self):
        rows = [("48601293", "1000", "1000.00"), ("48601300", "2000", "999.00")]
        tx, raw = _make_tx(_sap_rows(rows, FORCED))
        diffs: list[str] = []
        result = tx.clear_data_b(ENTRIES, _order(), diffs)
        self.assertTrue(result.success)
        self.assertTrue(result.changed)
        # 阶段一 bottom-up 删 2 个配对行（双表同选）；配对行删完后强制行上移到 row 0，
        # 阶段二重数剩余行（=1）再只选执行部门表删除——故第三次是 row 0 而非原始 row 2。
        self.assertEqual(raw.press_log, [
            "del zul=[1] kos=[1]",
            "del zul=[0] kos=[0]",
            "del zul=[0] kos=[]",
        ])
        # diffs[0] 是差异描述，diffs[1] 是删除汇总。
        self.assertIn("已删除旧 3 行", diffs[1])

    def test_delete_all_when_excel_empty(self):
        tx, raw = _make_tx(_sap_rows(SAP_ROWS, FORCED))
        result = tx.clear_data_b([], _order(), [])
        self.assertTrue(result.changed)
        self.assertEqual(len(raw.press_log), 3)


class RemapItemsTest(unittest.TestCase):
    """_remap_data_b_items：映射命中才替换，未命中/空映射一律保留原号。"""

    def test_maps_only_hit_items(self):
        entries = [_entry("48601293", 1.0, "1000"), _entry("48601300", 2.0, "2000")]
        mapped = OrderEditTransaction._remap_data_b_items(entries, {"1000": "1010"})
        self.assertEqual([e.item for e in mapped], ["1010", "2000"])

    def test_empty_map_returns_original(self):
        entries = [_entry("48601293", 1.0, "1000")]
        self.assertIs(OrderEditTransaction._remap_data_b_items(entries, {}), entries)

    def test_forced_row_untouched(self):
        entries = [_entry("48601258", 0.0, forced=True)]
        mapped = OrderEditTransaction._remap_data_b_items(entries, {"1000": "1010"})
        self.assertEqual(mapped[0].item, "")
        self.assertTrue(mapped[0].kostl_only)


if __name__ == "__main__":
    unittest.main()
