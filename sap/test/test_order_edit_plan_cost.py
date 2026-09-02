"""VA02 编辑：计划成本主键匹配 diff + 删除多余 + 追加新增 的回归测试。

覆盖：①打开编辑器容错缺失的 btnSPOP-VAROPTION1 弹窗（原报"找不到 SAP 元素"）；
②认领同键行仅金额差异才改 MENGE（不重写中心/类别）；③成本中心不同→删旧行 + 末尾新增；
④SAP 多余行（Excel 无）用 Shift+F2 删除；⑤按 item 号（而非位置）定位 SAP 物理行；
⑥SAP 无对应 item→成功跳过；⑦顺序不同但内容一致→零写入零回车；⑧Excel 多出的行追加到末尾。
"""

from __future__ import annotations

import os
import re
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

import sap.test.helpers  # noqa: E402,F401  仅为注入 win32com mock
from sap.models import PlanCostEntry, SapConfig  # noqa: E402
from sap.session import SapSession  # noqa: E402
from sap.transactions.order import OrderTransaction  # noqa: E402
from sap.transactions.order_edit import OrderEditTransaction  # noqa: E402

TABLE = "wnd[0]/usr/tblSAPLKKDI1301_TC"


def _typps(row): return f"{TABLE}/ctxtRK70L-TYPPS[2,{row}]"
def _herk2(row): return f"{TABLE}/ctxtRK70L-HERK2[3,{row}]"
def _herk3(row): return f"{TABLE}/ctxtRK70L-HERK3[4,{row}]"
def _menge(row): return f"{TABLE}/txtRK70L-MENGE[6,{row}]"


class _Element:
    def __init__(self, text: str = "", owner=None, element_id: str = ""):
        self.text = text
        self.key = ""
        self.caretPosition = 0
        self.focused = False
        self.vkeys: list[int] = []
        self._owner = owner
        self._id = element_id

    def setFocus(self):
        self.focused = True
        if self._owner is not None:
            self._owner.focused_id = self._id

    def sendVKey(self, k):
        self.vkeys.append(k)
        # Shift+F2(14) = 删除当前焦点所在行：真实 SAP 会移除该行并把后续行整体上移。
        if k == 14 and self._owner is not None:
            self._owner.delete_focused_plan_cost_row()

    def press(self): pass
    def select(self): pass


class _RawSession:
    """按 id 缓存控件；`missing` 中的 id 触发 findById 抛错（模拟弹窗不存在）。

    额外模拟计划成本编辑器的**删除行为**：Shift+F2 后该行消失、后续行上移。
    不模拟就会让"重读→删多余行"的循环判定不出进展，与真实 SAP 行为脱节。
    """

    _PC_COLUMNS = (_typps, _herk2, _herk3, _menge)

    def __init__(self, preset=None, missing=None):
        self._cache = preset or {}
        self._missing = missing or set()
        self.focused_id = ""

    def findById(self, element_id: str):
        if element_id in self._missing:
            raise Exception(f"not found: {element_id}")
        if element_id not in self._cache:
            self._cache[element_id] = _Element(owner=self, element_id=element_id)
        element = self._cache[element_id]
        # preset 里的控件构造时不知道 owner/id，首次取用时补齐，保证焦点与删除可追踪。
        if element._owner is None:
            element._owner, element._id = self, element_id
        return element

    def _plan_cost_row_count(self, max_rows: int = 50) -> int:
        count = 0
        for row in range(max_rows):
            if not (self._cache.get(_herk2(row)) and self._cache[_herk2(row)].text):
                break
            count += 1
        return count

    def delete_focused_plan_cost_row(self) -> None:
        """删除焦点所在的计划成本行：后续行逐列前移，末行清空。"""
        match = re.search(r"ctxtRK70L-HERK2\[3,(\d+)\]$", self.focused_id or "")
        if not match:
            return
        start = int(match.group(1))
        total = self._plan_cost_row_count()
        if start >= total:
            return
        for row in range(start, total - 1):
            for column in self._PC_COLUMNS:
                self.findById(column(row)).text = self.findById(column(row + 1)).text
        for column in self._PC_COLUMNS:
            self.findById(column(total - 1)).text = ""


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
    """预设 SAP item 概览行（供 OrderTransaction.find_item_row_by_no 按 item 号定位物理行）。"""
    return {
        OrderTransaction._item_id(row): _Element(item_no),
        OrderTransaction._material_id(row): _Element(material),
    }


def _existing(cost_center, category, amount, row):
    return {
        _typps(row): _Element("E"),
        _herk2(row): _Element(cost_center),
        _herk3(row): _Element(category),
        _menge(row): _Element(amount),
    }


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

    def test_amount_change_overwrites_menge(self):
        # 认领同键行：仅金额有差异才改 MENGE，不重写中心/类别，日志带箭头。
        preset = {**_item_row("10", "M1", 0), **_existing("1100", "FREMDL", "100.00", 0)}
        tx, raw = _make_tx(preset, self.MISSING)
        entry = PlanCostEntry(cost_center="1100", category="FREMDL", amount=500.0)
        diffs: list[str] = []
        result = tx.edit_plan_cost([entry], diffs, target_item="10")
        self.assertTrue(result.success, result.message)
        self.assertEqual(raw.findById(_menge(0)).text, "500.00")   # 只改金额
        self.assertEqual(raw.findById(_herk2(0)).text, "1100")     # 中心未被重写
        self.assertEqual(raw.findById(_herk3(0)).text, "FREMDL")   # 类别未被重写
        self.assertIn(0, raw.findById("wnd[0]").vkeys)             # 有差异→提交
        self.assertEqual(diffs, ["计划成本 成本中心1100(FREMDL) 金额 100.00→500.00"])

    def test_center_mismatch_deletes_and_adds(self):
        # 成本中心不同：主键认领不到→删除旧行(1100) + 末尾新增(2200)。
        preset = {**_item_row("10", "M1", 0), **_existing("1100", "FREMDL", "100.00", 0)}
        tx, raw = _make_tx(preset, self.MISSING)
        entry = PlanCostEntry(cost_center="2200", category="FREMDL", amount=300.0)
        diffs: list[str] = []
        result = tx.edit_plan_cost([entry], diffs, target_item="10")
        self.assertTrue(result.success, result.message)
        # 旧行 1100 删除：聚焦 HERK2[3,0] + Shift+F2(vkey14)。
        self.assertTrue(raw.findById(_herk2(0)).focused)
        self.assertIn(14, raw.findById("wnd[0]").vkeys)
        # 新行 2200 追加到"删除完成后"的末尾——旧行已被删掉，当前末尾即 row0（全写四列）。
        # 追加位置每次实时重读行数取得，不用删除前的行数递推。
        self.assertEqual(raw.findById(_typps(0)).text, "E")
        self.assertEqual(raw.findById(_herk2(0)).text, "2200")
        self.assertEqual(raw.findById(_menge(0)).text, "300.00")
        self.assertIn("计划成本 成本中心1100(FREMDL) 金额 100.00：Excel 无，已删除", diffs)
        self.assertIn("计划成本 成本中心2200(FREMDL) 金额 (空)→300.00", diffs)

    def test_deletes_sap_extra_rows(self):
        # row0 金额相等→不重写；SAP 多出的 row1（Excel 无）用 Shift+F2 删除。
        preset = {
            **_item_row("10", "M1", 0),
            **_existing("1100", "FREMDL", "100.00", 0), **_existing("2200", "T01AST", "8.00", 1),
        }
        tx, raw = _make_tx(preset, self.MISSING)
        entry = PlanCostEntry(cost_center="1100", category="FREMDL", amount=100.0)
        diffs: list[str] = []
        result = tx.edit_plan_cost([entry], diffs, target_item="10")
        self.assertTrue(result.success, result.message)
        self.assertEqual(raw.findById(_menge(0)).text, "100.00")   # 金额相等未重写
        self.assertTrue(raw.findById(_herk2(1)).focused)           # 删除聚焦 row1
        self.assertIn(14, raw.findById("wnd[0]").vkeys)
        self.assertEqual(diffs, ["计划成本 成本中心2200(T01AST) 时间 8.00：Excel 无，已删除"])

    def test_same_content_different_order_no_change(self):
        # 顺序不同但内容一致（含前导零中心 + 千分位金额）→ 零写入零回车、无差异。
        preset = {
            **_item_row("10", "M1", 0),
            **_existing("0048601240", "FREMDL", "1,000.00", 0),
            **_existing("0048601294", "FREMDL", "100.00", 1),
        }
        tx, raw = _make_tx(preset, self.MISSING)
        entries = [
            PlanCostEntry(cost_center="48601294", category="FREMDL", amount=100.0),
            PlanCostEntry(cost_center="48601240", category="FREMDL", amount=1000.0),
        ]
        diffs: list[str] = []
        result = tx.edit_plan_cost(entries, diffs, target_item="10")
        self.assertTrue(result.success, result.message)
        self.assertEqual(diffs, [])
        self.assertEqual(result.message, "Plan Cost 无差异")
        self.assertNotIn(0, raw.findById("wnd[0]").vkeys)   # 无金额提交
        self.assertNotIn(14, raw.findById("wnd[0]").vkeys)  # 无删除

    def test_new_rows_appended(self):
        # SAP 1 行匹配（金额相等不写）；Excel 多出的新键行追加到末尾 row1。
        preset = {**_item_row("10", "M1", 0), **_existing("1100", "FREMDL", "100.00", 0)}
        tx, raw = _make_tx(preset, self.MISSING)
        entries = [
            PlanCostEntry(cost_center="1100", category="FREMDL", amount=100.0),
            PlanCostEntry(cost_center="2200", category="T01AST", amount=5.0),
        ]
        diffs: list[str] = []
        result = tx.edit_plan_cost(entries, diffs, target_item="10")
        self.assertTrue(result.success, result.message)
        self.assertEqual(raw.findById(_menge(0)).text, "100.00")   # 匹配行金额未重写
        self.assertEqual(raw.findById(_typps(1)).text, "E")        # 新行追加到 row1
        self.assertEqual(raw.findById(_herk2(1)).text, "2200")
        self.assertEqual(raw.findById(_herk3(1)).text, "T01AST")
        self.assertEqual(raw.findById(_menge(1)).text, "5.00")
        self.assertEqual(diffs, ["计划成本 成本中心2200(T01AST) 时间 (空)→5.00"])

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
