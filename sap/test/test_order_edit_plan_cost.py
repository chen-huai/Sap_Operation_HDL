"""VA02 编辑：计划成本按 成本中心+类别 匹配更新的回归测试。

覆盖：①打开编辑器容错缺失的 btnSPOP-VAROPTION1 弹窗（原报"找不到 SAP 元素"）；
②中心+类别一致→仅更新金额/时间；③中心或类别不同→新增；④SAP 多余行→提示；
⑤按 item 号（而非位置）定位 SAP 物理行；⑥SAP 无对应 item→成功跳过。
"""

from __future__ import annotations

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

import sap.test.helpers  # noqa: E402,F401  仅为注入 win32com mock
from sap.models import PlanCostEntry, SapConfig  # noqa: E402
from sap.session import SapSession  # noqa: E402
from sap.transactions.order import OrderTransaction  # noqa: E402
from sap.transactions.order_edit import OrderEditTransaction  # noqa: E402

TABLE = "wnd[0]/usr/tblSAPLKKDI1301_TC"


def _herk2(row): return f"{TABLE}/ctxtRK70L-HERK2[3,{row}]"
def _herk3(row): return f"{TABLE}/ctxtRK70L-HERK3[4,{row}]"
def _menge(row): return f"{TABLE}/txtRK70L-MENGE[6,{row}]"


class _Element:
    def __init__(self, text: str = ""):
        self.text = text
        self.key = ""
        self.caretPosition = 0
        self.focused = False

    def setFocus(self): self.focused = True
    def sendVKey(self, _k): pass
    def press(self): pass
    def select(self): pass


class _RawSession:
    """按 id 缓存控件；`missing` 中的 id 触发 findById 抛错（模拟弹窗不存在）。"""

    def __init__(self, preset=None, missing=None):
        self._cache = preset or {}
        self._missing = missing or set()

    def findById(self, element_id: str):
        if element_id in self._missing:
            raise Exception(f"not found: {element_id}")
        if element_id not in self._cache:
            self._cache[element_id] = _Element()
        return self._cache[element_id]


def _make_tx(preset=None, missing=None):
    raw = _RawSession(preset, missing)
    session = SapSession(raw, raw, raw, raw)
    config = SapConfig(
        order_type="ZOR", sales_organization="3002", distribution_channels="10",
        sales_office="1000", sales_group="200", sub_cost_center_cs="1101",
        sub_cost_center_chm="1102", sub_cost_center_phy="1103", cs_code="CS001", sales_code="SA001",
    )
    return OrderEditTransaction(session, config), raw


def _item_row(item_no, material, row):
    """预设 SAP item 概览行（供 _find_item_physical_row 按 item 号定位物理行）。"""
    return {
        OrderTransaction._item_id(row): _Element(item_no),
        OrderTransaction._material_id(row): _Element(material),
    }


def _existing(cost_center, category, amount, row):
    return {_herk2(row): _Element(cost_center), _herk3(row): _Element(category), _menge(row): _Element(amount)}


class EditPlanCostTest(unittest.TestCase):
    # 所有用例都将 VAROPTION1 设为缺失，验证打开编辑器的容错。
    MISSING = {"wnd[1]/usr/btnSPOP-VAROPTION1"}

    def test_open_tolerates_missing_variant_popup(self):
        preset = {**_item_row("10", "M1", 0), **_existing("1100", "FREMDL", "500.00", 0)}
        tx, _ = _make_tx(preset, self.MISSING)
        entry = PlanCostEntry(cost_center="1100", category="FREMDL", amount=500.0)
        diffs: list[str] = []
        result = tx.edit_plan_cost([entry], diffs, target_item="10")
        self.assertTrue(result.success, result.message)  # 弹窗缺失不再报错

    def test_matched_updates_amount_only(self):
        preset = {**_item_row("10", "M1", 0), **_existing("1100", "FREMDL", "100.00", 0)}
        tx, raw = _make_tx(preset, self.MISSING)
        entry = PlanCostEntry(cost_center="1100", category="FREMDL", amount=500.0)
        diffs: list[str] = []
        result = tx.edit_plan_cost([entry], diffs, target_item="10")
        self.assertTrue(result.success, result.message)
        self.assertEqual(raw.findById(_menge(0)).text, "500.00")  # 金额已更新
        self.assertEqual(diffs, ["计划成本 成本中心1100(FREMDL) 金额 100.00 → 500.00"])

    def test_matched_no_diff_outputs_nothing(self):
        preset = {**_item_row("10", "M1", 0), **_existing("1100", "T01AST", "8.00", 0)}
        tx, raw = _make_tx(preset, self.MISSING)
        entry = PlanCostEntry(cost_center="1100", category="T01AST", amount=8.0)
        diffs: list[str] = []
        result = tx.edit_plan_cost([entry], diffs, target_item="10")
        self.assertTrue(result.success, result.message)
        # 无变化 → 不输出任何行（只显示有更新的）。
        self.assertEqual(diffs, [])

    def test_different_center_adds_new_row(self):
        preset = {**_item_row("10", "M1", 0), **_existing("1100", "FREMDL", "100.00", 0)}
        tx, raw = _make_tx(preset, self.MISSING)
        entry = PlanCostEntry(cost_center="2200", category="FREMDL", amount=300.0)
        diffs: list[str] = []
        result = tx.edit_plan_cost([entry], diffs, target_item="10")
        self.assertTrue(result.success, result.message)
        # 新行落 row1，写全字段。
        self.assertEqual(raw.findById(_herk2(1)).text, "2200")
        self.assertEqual(raw.findById(_menge(1)).text, "300.00")
        self.assertIn("计划成本 成本中心2200(FREMDL) 新增金额 300.00", diffs)
        # 原 1100/FREMDL 未被 ODM 命中 → 多余提示。
        self.assertIn("计划成本 成本中心1100(FREMDL) 金额 100.00：SAP 有、Excel 无，已跳过", diffs)

    def test_sap_extra_row_warns(self):
        preset = {
            **_item_row("10", "M1", 0),
            **_existing("1100", "FREMDL", "100.00", 0), **_existing("2200", "T01AST", "8.00", 1),
        }
        tx, _ = _make_tx(preset, self.MISSING)
        entry = PlanCostEntry(cost_center="1100", category="FREMDL", amount=100.0)
        diffs: list[str] = []
        result = tx.edit_plan_cost([entry], diffs, target_item="10")
        self.assertTrue(result.success, result.message)
        # 1100/FREMDL 无变化 → 不输出；仅 2200/T01AST 多余提示（T01AST 标签为"时间"）。
        self.assertEqual(diffs, ["计划成本 成本中心2200(T01AST) 时间 8.00：SAP 有、Excel 无，已跳过"])

    def test_locates_row_by_item_number_not_position(self):
        # SAP item 1000(行0)/2000(行1)；ODM item 2000 应定位到物理行1，而非位置0。
        preset = {
            **_item_row("1000", "M1", 0), **_item_row("2000", "M2", 1),
            **_existing("1100", "FREMDL", "100.00", 0),
        }
        tx, raw = _make_tx(preset, self.MISSING)
        entry = PlanCostEntry(cost_center="1100", category="FREMDL", amount=500.0)
        diffs: list[str] = []
        result = tx.edit_plan_cost([entry], diffs, target_item="2000")
        self.assertTrue(result.success, result.message)
        # 打开编辑器聚焦的是行1物料格（item 2000），证明按 item 号而非位置定位。
        self.assertTrue(raw.findById(OrderTransaction._material_id(1)).focused)
        self.assertFalse(raw.findById(OrderTransaction._material_id(0)).focused)

    def test_item_not_in_sap_skips(self):
        # SAP 只有 item 1000/2000；ODM 计划成本属于 item 4000 → SAP 无对应 → 成功跳过。
        preset = {**_item_row("1000", "M1", 0), **_item_row("2000", "M2", 1)}
        tx, raw = _make_tx(preset, self.MISSING)
        entry = PlanCostEntry(cost_center="1100", category="FREMDL", amount=500.0)
        diffs: list[str] = []
        result = tx.edit_plan_cost([entry], diffs, target_item="4000")
        self.assertTrue(result.success, result.message)  # 不失败、不阻断
        self.assertTrue(result.warning)                  # 但标警告，UI 用区别色提示
        self.assertEqual(diffs, [])                       # 不开编辑器、不产生差异
        self.assertIn("SAP 无对应 item 4000", result.message)
        self.assertIn("已跳过", result.message)


if __name__ == "__main__":
    unittest.main()
