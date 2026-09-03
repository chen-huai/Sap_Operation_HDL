"""VA02 编辑：item 按 item+物料 双键对比更新的回归测试。

覆盖根因修复（`Property '.text' can not be set`）：命中行只更新金额、绝不改写只读物料；
并验证三条规则：均一致→更新金额；item/物料有一不同→新增；SAP 多余行→提示。
"""

from __future__ import annotations

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

import sap.test.helpers  # noqa: E402,F401  仅为注入 win32com mock
from sap.exceptions import SapUiError  # noqa: E402
from sap.models import OrderData, OrderItemData, SapConfig  # noqa: E402
from sap.session import SapSession  # noqa: E402
from sap.transactions.order import OrderTransaction  # noqa: E402
from sap.transactions.order_edit import OrderEditTransaction  # noqa: E402

CONDITION_BASE = (
    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/"
    "ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/"
)
# 描述列控件名取候选表首项（生产代码会逐个探测，测试固定用第一个即可）。
DESC_COLUMN = OrderEditTransaction._CONDITION_TEXT_CANDIDATES[0]


def _desc(row: int) -> str:
    """条件表"描述"列（[2,row]）控件 ID。"""
    return f"{CONDITION_BASE}{DESC_COLUMN}[2,{row}]"


def _amount(row: int) -> str:
    """条件表金额列（[3,row]）控件 ID。"""
    return f"{CONDITION_BASE}txtKOMV-KBETR[3,{row}]"


def _currency(row: int) -> str:
    """条件表币种列（[4,row]）控件 ID —— 恒与金额列同行。"""
    return f"{CONDITION_BASE}ctxtRV61A-KOEIN[4,{row}]"


# 已存在 item 编辑：价格条件行由描述列定位（默认预设放在第 1 行，与旧行为一致）；
# 新增 item：沿用创建侧第 5 行。
CONDITION_ID = _amount(1)
NEW_CONDITION_ID = (
    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/"
    "ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]"
)
LONG_TEXT_ID = (
    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09/"
    "ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/"
    "cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell"
)
LANG_ID = (
    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09/"
    "ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cmbLV70T-SPRAS"
)


class _Element:
    """可控 SAP 控件；read_only/key_read_only=True 时写 .text/.key 抛错（模拟只读字段）。"""

    def __init__(self, text: str = "", *, read_only: bool = False, key_read_only: bool = False):
        self._text = text
        self._read_only = read_only
        self._key_read_only = key_read_only
        self._key = ""
        self.caretPosition = 0

    @property
    def text(self):
        return self._text

    @text.setter
    def text(self, value):
        if self._read_only:
            raise ValueError("Property '.text' can not be set.")
        self._text = value

    @property
    def key(self):
        return self._key

    @key.setter
    def key(self, value):
        if self._key_read_only:
            raise ValueError("Property '.key' can not be set.")
        self._key = value

    def setFocus(self):
        pass

    def sendVKey(self, _key):
        pass

    def press(self):
        pass

    def select(self):
        pass


class _RawSession:
    """按 element_id 缓存 `_Element`；未知 id 返回空控件。

    `missing` 中的 id 触发 findById 抛错——用于模拟"该控件名在本屏不存在"
    （描述列候选探测依赖这个区分：控件不存在要换候选名，读到空只是扫到表尾）。
    """

    def __init__(
        self,
        preset: dict[str, _Element] | None = None,
        missing: set[str] | None = None,
    ):
        self._cache: dict[str, _Element] = preset or {}
        self._missing: set[str] = missing or set()

    def findById(self, element_id: str) -> _Element:
        if element_id in self._missing:
            raise Exception(f"not found: {element_id}")
        if element_id not in self._cache:
            self._cache[element_id] = _Element()
        return self._cache[element_id]


def _make_tx(preset: dict[str, _Element], missing: set[str] | None = None):
    raw = _RawSession(preset, missing)
    session = SapSession(raw, raw, raw, raw)
    config = SapConfig(
        order_type="ZOR", sales_organization="3002", distribution_channels="10",
        sales_office="1000", sales_group="200", sub_cost_center_cs="1101",
        sub_cost_center_chm="1102", sub_cost_center_phy="1103", cs_code="CS001", sales_code="SA001",
    )
    return OrderEditTransaction(session, config), raw


def _existing_row(item_no: str, material: str, row: int, *, amount: str = "0.00", material_read_only: bool = True):
    """构造一行已存在 item 的控件预设（物料默认只读，模拟已落盘行；金额=概览第5格 NETWR）。"""
    return {
        OrderTransaction._item_id(row): _Element(item_no),
        OrderTransaction._material_id(row): _Element(material, read_only=material_read_only),
        OrderTransaction._net_value_id(row): _Element(amount),
    }


def _price_rows(price_row: int) -> dict[str, _Element]:
    """构造条件表描述列预设：price_row 之前是 Planned Costs，price_row 是 Price。

    价格行行号随"该单创建时有没有录 plan cost"变化——这是本模块要覆盖的核心变量：
    有 plan cost → row0=Planned Costs / row1=Price；没有 → row0=Price。
    """
    preset = {_desc(row): _Element("Planned Costs") for row in range(price_row)}
    preset[_desc(price_row)] = _Element("Price")
    return preset


def _order(*items: OrderItemData) -> OrderData:
    return OrderData(
        sap_no="123456", project_no="PRJ-001", currency_type="CNY",
        exchange_rate=1.0, short_text="T", items=list(items),
    )


class EditItemsMatchTest(unittest.TestCase):
    def test_same_item_and_material_updates_amount_only(self):
        # 现有行 item=10/物料 T75-405-00（物料只读）；ODM 同 item 同物料 → 只更新金额。
        # 条件页现值 100.00 ≠ ODM 5000 → item/物料一致但金额不同，进详情更新金额（不碰物料）。
        preset = {**_existing_row("10", "T75-405-00", 0), **_price_rows(1)}
        preset[CONDITION_ID] = _Element("100.00")
        tx, raw = _make_tx(preset)
        order = _order(OrderItemData(item="10", material_code="T75-405-00", revenue=5000.0))
        diffs: list[str] = []

        result = tx.edit_items(order, diffs)  # 若误写只读物料会在此抛错→失败

        self.assertTrue(result.success, result.message)
        self.assertEqual(raw.findById(OrderTransaction._material_id(0)).text, "T75-405-00")  # 物料未变
        self.assertEqual(raw.findById(CONDITION_ID).text, "5000.00")  # 编辑条件行 [3,1] 已写入
        # 单行汇总：item/物料/金额旧→新（旧值取自条件页 [3,1]）。
        self.assertEqual(diffs, ["item 10 物料 T75-405-00 金额 100.00→5000.00"])

    def test_matched_no_diff_outputs_nothing(self):
        # item+物料一致，进详情比对：条件页金额已等于 ODM、无长文本 → 不写，也不输出（只显示有更新的）。
        preset = {**_existing_row("10", "T75-405-00", 0), **_price_rows(1)}
        preset[CONDITION_ID] = _Element("5000.00")  # 现值已等于 ODM revenue
        tx, raw = _make_tx(preset)
        order = _order(OrderItemData(item="10", material_code="T75-405-00", revenue=5000.0))
        diffs: list[str] = []

        result = tx.edit_items(order, diffs)

        self.assertTrue(result.success, result.message)
        self.assertEqual(raw.findById(CONDITION_ID).text, "5000.00")  # 比对相等，未重写
        self.assertEqual(diffs, [])  # 无变化 → 不输出

    def test_long_text_update_does_not_touch_language(self):
        # 长文本不同 → 更新文本；语言下拉只读(编辑不可改)，若误调 set_key 会抛错→整体失败。
        preset = {**_existing_row("10", "T75-405-00", 0), **_price_rows(1)}
        preset[CONDITION_ID] = _Element("5000.00")  # 金额一致，仅长文本变化
        preset[LONG_TEXT_ID] = _Element("旧描述")
        preset[LANG_ID] = _Element(key_read_only=True)  # 模拟编辑屏语言不可改
        tx, raw = _make_tx(preset)
        order = _order(OrderItemData(
            item="10", material_code="T75-405-00", revenue=5000.0, long_text="新描述",
        ))
        diffs: list[str] = []

        result = tx.edit_items(order, diffs)

        self.assertTrue(result.success, result.message)  # 未触碰语言 → 不报错
        self.assertEqual(raw.findById(LONG_TEXT_ID).text, "新描述")  # 文本已更新
        self.assertEqual(raw.findById(LANG_ID).key, "")  # 语言保持原值
        # 金额无变化不输出，仅文本旧→新。
        self.assertEqual(diffs, ["item 10 物料 T75-405-00 文本 旧描述→新描述"])

    def test_different_material_adds_new_row(self):
        # item 同号但物料不同 → 新增一条；item 号已存在故 SAP 自动分配（不写 POSNR）。
        preset = {**_existing_row("10", "T75-405-00", 0), **_price_rows(1)}
        preset[CONDITION_ID] = _Element("0.00")
        tx, raw = _make_tx(preset)
        order = _order(OrderItemData(item="10", material_code="T20-430-00", revenue=3000.0))
        diffs: list[str] = []

        result = tx.edit_items(order, diffs)

        self.assertTrue(result.success, result.message)
        # 新行落在 next_row=1：物料写入，POSNR 不写（让 SAP 自动分配）。
        self.assertEqual(raw.findById(OrderTransaction._material_id(1)).text, "T20-430-00")
        self.assertEqual(raw.findById(OrderTransaction._item_id(1)).text, "")
        # 新增金额走创建侧条件行 [3,5]，而非编辑行 [3,1]。
        self.assertEqual(raw.findById(NEW_CONDITION_ID).text, "3000.00")
        # 新增一行 + 原 item 10/T75 物料已变成孤儿行（SAP 有 ODM 无）一行。
        self.assertIn("item 10 物料 T20-430-00 新增金额 3000.00", diffs)
        self.assertIn("item 10 物料 T75-405-00 金额 0.00：SAP 有、Excel 无，已跳过", diffs)

    def test_sap_extra_row_warns(self):
        # SAP 有 item 20 但 ODM 表无 → 提示并记 log，不删不改。
        preset = {
            **_existing_row("10", "T75-405-00", 0),
            **_existing_row("20", "T20-405-00", 1),
            **_price_rows(1),
        }
        preset[CONDITION_ID] = _Element("5000.00")  # item 10 金额已一致 → 无更新
        tx, _ = _make_tx(preset)
        order = _order(OrderItemData(item="10", material_code="T75-405-00", revenue=5000.0))
        diffs: list[str] = []

        result = tx.edit_items(order, diffs)

        self.assertTrue(result.success, result.message)
        # item 10 命中但无更新 → 不输出；仅 item 20 多余提示一行。
        self.assertEqual(diffs, ["item 20 物料 T20-405-00 金额 0.00：SAP 有、Excel 无，已跳过"])


class PriceConditionRowTest(unittest.TestCase):
    """价格条件行按描述列定位：行号随"该单有没有 plan cost"变化，不能写死。

    实机结论（2026-09-03，订单 7482678489）：
        创建时录过 plan cost → row0=Planned Costs、row1=Price
        创建时没录 plan cost → row0=Price
    旧实现写死 [3,1]，对后者会写到只读条件行上 → `Property '.text' can not be set.`
    """

    def test_no_plan_cost_writes_row_zero(self):
        # 本次故障场景：没有 plan cost，Price 在 row0 —— 旧实现写 row1 会被 SAP 拒。
        preset = {**_existing_row("10", "T75-405-00", 0), **_price_rows(0)}
        preset[_amount(0)] = _Element("100.00")
        # row1 是只读条件行：若代码仍写死 row1，这里会抛错让用例失败。
        preset[_amount(1)] = _Element("999.00", read_only=True)
        tx, raw = _make_tx(preset)
        order = _order(OrderItemData(item="10", material_code="T75-405-00", revenue=680.0))
        diffs: list[str] = []

        result = tx.edit_items(order, diffs)

        self.assertTrue(result.success, result.message)
        self.assertEqual(raw.findById(_amount(0)).text, "680.00")   # 写在 Price 行
        self.assertEqual(raw.findById(_amount(1)).text, "999.00")   # 只读行未被触碰
        self.assertEqual(diffs, ["item 10 物料 T75-405-00 金额 100.00→680.00"])

    def test_with_plan_cost_writes_row_one(self):
        # 有 plan cost：row0=Planned Costs（绝不能写）、row1=Price。
        preset = {**_existing_row("10", "T75-405-00", 0), **_price_rows(1)}
        preset[_amount(0)] = _Element("50.00")    # plan cost 金额，必须原样保留
        preset[_amount(1)] = _Element("100.00")
        tx, raw = _make_tx(preset)
        order = _order(OrderItemData(item="10", material_code="T75-405-00", revenue=680.0))
        diffs: list[str] = []

        result = tx.edit_items(order, diffs)

        self.assertTrue(result.success, result.message)
        self.assertEqual(raw.findById(_amount(1)).text, "680.00")   # 写在 Price 行
        self.assertEqual(raw.findById(_amount(0)).text, "50.00")    # plan cost 金额未被改
        self.assertEqual(diffs, ["item 10 物料 T75-405-00 金额 100.00→680.00"])

    def test_price_row_absent_skips_and_warns(self):
        # 条件表里没有 Price 行 → 告警跳过、不写任何一行（写到猜出来的行上会改错 plan cost）。
        preset = {**_existing_row("10", "T75-405-00", 0)}
        preset[_desc(0)] = _Element("Planned Costs")
        preset[_amount(0)] = _Element("50.00")
        tx, raw = _make_tx(preset)
        order = _order(OrderItemData(item="10", material_code="T75-405-00", revenue=680.0))
        diffs: list[str] = []

        result = tx.edit_items(order, diffs)

        self.assertTrue(result.success, result.message)
        self.assertTrue(result.warning)
        self.assertEqual(raw.findById(_amount(0)).text, "50.00")    # 一行都没动
        self.assertEqual(
            diffs, ["item 10 物料 T75-405-00：未找到 Price 条件行，金额未写入(待人工核对)"]
        )

    def test_currency_follows_price_row_without_plan_cost(self):
        # 币种列与价格行同行（用户确认的业务约束）：没有 plan cost → 币种也在 row0。
        preset = {**_existing_row("10", "T75-405-00", 0), **_price_rows(0)}
        preset[_amount(0)] = _Element("680.00")           # 金额已一致，只有币种要改
        preset[_currency(0)] = _Element("USD")
        # row1 是只读条件行：若币种写死 row1，这里会抛错让用例失败。
        preset[_currency(1)] = _Element("EUR", read_only=True)
        tx, raw = _make_tx(preset)
        order = _order(OrderItemData(item="10", material_code="T75-405-00", revenue=680.0))
        diffs: list[str] = []

        result = tx.edit_items(order, diffs)   # order.currency_type = CNY

        self.assertTrue(result.success, result.message)
        self.assertEqual(raw.findById(_currency(0)).text, "CNY")   # 币种写在 Price 行
        self.assertEqual(raw.findById(_currency(1)).text, "EUR")   # 只读行未被触碰
        self.assertEqual(diffs, ["item 10 物料 T75-405-00 币种 USD→CNY"])

    def test_currency_follows_price_row_with_plan_cost(self):
        # 有 plan cost → 价格行是 row1，币种也必须落 row1；row0 是 plan cost 行，不能动。
        preset = {**_existing_row("10", "T75-405-00", 0), **_price_rows(1)}
        preset[_amount(1)] = _Element("680.00")
        preset[_currency(0)] = _Element("USD")   # plan cost 行币种，必须原样保留
        preset[_currency(1)] = _Element("USD")
        tx, raw = _make_tx(preset)
        order = _order(OrderItemData(item="10", material_code="T75-405-00", revenue=680.0))
        diffs: list[str] = []

        result = tx.edit_items(order, diffs)

        self.assertTrue(result.success, result.message)
        self.assertEqual(raw.findById(_currency(1)).text, "CNY")   # 币种写在 Price 行
        self.assertEqual(raw.findById(_currency(0)).text, "USD")   # plan cost 行币种未被改
        self.assertEqual(diffs, ["item 10 物料 T75-405-00 币种 USD→CNY"])

if __name__ == "__main__":
    unittest.main()
