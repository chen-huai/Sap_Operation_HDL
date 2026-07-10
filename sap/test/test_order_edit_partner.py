"""VA02 编辑：Sales 伙伴行"从无到有"新增的回归测试。

覆盖根因修复：订单原本无 VE(Sales) 行时，编辑应在伙伴表空行追加一行写入 sales_code，
而非静默跳过（原 `_edit_partners` 仅有"改已存在行"分支）。
"""

from __future__ import annotations

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

import sap.test.helpers  # noqa: E402,F401  仅为注入 win32com mock
from sap.models import OrderData, SapConfig  # noqa: E402
from sap.session import SapSession  # noqa: E402
from sap.transactions.order_edit import OrderEditTransaction  # noqa: E402


def _make_config() -> SapConfig:
    # cs_code 置空以隔离出仅 Sales 路径，避免 GPC/CS 干扰。
    return SapConfig(
        order_type="ZOR",
        sales_organization="3002",
        distribution_channels="10",
        sales_office="1000",
        sales_group="200",
        sub_cost_center_cs="1101",
        sub_cost_center_chm="1102",
        sub_cost_center_phy="1103",
        cs_code="",
        sales_code="SA001",
    )

PARTNER_PREFIX = (
    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/"
    "ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:"
    "SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW"
)


class _Element:
    """可控 SAP 控件：text/key 可读写，焦点/按键为 no-op。"""

    def __init__(self, text: str = "", key: str = ""):
        self.text = text
        self.key = key
        self.caretPosition = 0

    def setFocus(self):
        pass

    def sendVKey(self, _key):
        pass

    def press(self):
        pass

    def select(self):
        pass


class _RawSession:
    """按 element_id 缓存 `_Element` 的最小 raw session；未知 id 返回空控件。"""

    def __init__(self, preset: dict[str, _Element] | None = None):
        self._cache: dict[str, _Element] = preset or {}

    def findById(self, element_id: str) -> _Element:
        if element_id not in self._cache:
            self._cache[element_id] = _Element()
        return self._cache[element_id]


def _make_transaction(preset: dict[str, _Element] | None = None):
    raw = _RawSession(preset)
    session = SapSession(raw, raw, raw, raw)
    return OrderEditTransaction(session, _make_config()), raw


class FindEmptyPartnerRowTest(unittest.TestCase):
    def test_returns_first_empty_row(self):
        # 行0=ZG已占用，行1=空，断言返回 1。
        preset = {
            f"{PARTNER_PREFIX}/cmbGVS_TC_DATA-REC-PARVW[0,0]": _Element(key="ZG"),
            f"{PARTNER_PREFIX}/ctxtGVS_TC_DATA-REC-PARTNER[1,0]": _Element(text="GP001"),
        }
        tx, _ = _make_transaction(preset)
        self.assertEqual(tx._find_empty_partner_row(PARTNER_PREFIX, max_rows=4), 1)


class AddPartnerRowTest(unittest.TestCase):
    def test_adds_ve_row_on_empty_table(self):
        tx, raw = _make_transaction()
        diffs: list[str] = []
        tx._add_partner_row(PARTNER_PREFIX, "VE", "SA001", field="Sales", diffs=diffs)

        parvw = raw.findById(f"{PARTNER_PREFIX}/cmbGVS_TC_DATA-REC-PARVW[0,0]")
        partner = raw.findById(f"{PARTNER_PREFIX}/ctxtGVS_TC_DATA-REC-PARTNER[1,0]")
        self.assertEqual(parvw.key, "VE")
        self.assertEqual(partner.text, "SA001")
        self.assertEqual(diffs, ["Sales:(空)→SA001"])

    def test_empty_value_is_noop(self):
        tx, _ = _make_transaction()
        diffs: list[str] = []
        tx._add_partner_row(PARTNER_PREFIX, "VE", "", field="Sales", diffs=diffs)
        self.assertEqual(diffs, [])


class EditPartnersSalesBranchTest(unittest.TestCase):
    def test_missing_ve_row_triggers_add(self):
        # 伙伴表全空（无 VE 行）→ _edit_partners 应新增 VE 行写入 sales_code。
        tx, raw = _make_transaction()
        order = OrderData(  # global_partner_code 空 → 关闭 GPC 分支，聚焦 Sales
            sap_no="123456",
            project_no="PRJ-001",
            currency_type="CNY",
            exchange_rate=1.0,
            short_text="Test",
            global_partner_code="",
        )
        diffs: list[str] = []
        tx._edit_partners(order, diffs)

        parvw = raw.findById(f"{PARTNER_PREFIX}/cmbGVS_TC_DATA-REC-PARVW[0,0]")
        partner = raw.findById(f"{PARTNER_PREFIX}/ctxtGVS_TC_DATA-REC-PARTNER[1,0]")
        self.assertEqual(parvw.key, "VE")
        self.assertEqual(partner.text, "SA001")
        self.assertIn("Sales:(空)→SA001", diffs)


if __name__ == "__main__":
    unittest.main()
