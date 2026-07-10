"""订单价值(AUFTRAGSWERT) 回填/编辑的回归测试。

覆盖本次重构核心（Σ SAP item 未税净值 × 汇率，替代旧的双重汇率口径）：
①净值加和遇空行停止；②扫满上限时告警而非静默少算；
③创建 fill_order_value：≥阈值写入、<阈值不写（SAP 自动带出）；
④编辑 edit_order_value：≥阈值对比写入、<阈值清空脏值自愈、相等不写。

领域规则：阈值 3.5w。<3.5w 时目标留空（SAP 自动填），≥3.5w 时必须自己写换算值。
"""

from __future__ import annotations

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

import sap.test.helpers  # noqa: E402,F401  仅为注入 win32com mock
from sap.models import OrderData, SapConfig  # noqa: E402
from sap.session import SapSession  # noqa: E402
from sap.transactions.order import OrderTransaction  # noqa: E402
from sap.transactions.order_edit import OrderEditTransaction  # noqa: E402


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
    """按 id 缓存控件；未预设的 id 自动生成空控件（text=""）。"""

    def __init__(self, preset=None):
        self._cache = preset or {}

    def findById(self, element_id: str):
        if element_id not in self._cache:
            self._cache[element_id] = _Element()
        return self._cache[element_id]


def _make_config() -> SapConfig:
    # revenue_threshold 取默认 35000（=3.5w）。
    return SapConfig(
        order_type="ZOR", sales_organization="3002", distribution_channels="10",
        sales_office="1000", sales_group="200", sub_cost_center_cs="1101",
        sub_cost_center_chm="1102", sub_cost_center_phy="1103", cs_code="CS001", sales_code="SA001",
    )


def _make_order(rate: float) -> OrderData:
    return OrderData(
        sap_no="1", project_no="P1", currency_type="USD",
        exchange_rate=rate, short_text="t",
    )


def _net_row(item_no: str, net: str, row: int) -> dict:
    """预设 item 概览一行：item 号 + 未税净值(VBAP-NETWR)。"""
    return {
        OrderTransaction._item_id(row): _Element(item_no),
        OrderTransaction._net_value_id(row): _Element(net),
    }


def _make_create_tx(preset=None):
    raw = _RawSession(preset)
    session = SapSession(raw, raw, raw, raw)
    return OrderTransaction(session, _make_config()), raw


def _make_edit_tx(preset=None):
    raw = _RawSession(preset)
    session = SapSession(raw, raw, raw, raw)
    return OrderEditTransaction(session, _make_config()), raw


AUFTRAGSWERT = OrderTransaction._auftragswert_id()


class SumItemNetValuesTest(unittest.TestCase):
    """净值加和 helper：遇空行停止、扫满上限告警。"""

    def test_sums_until_empty_row(self):
        preset = {**_net_row("1000", "10000.00", 0), **_net_row("2000", "5000.00", 1)}
        tx, _ = _make_create_tx(preset)
        total, truncated = tx._sum_item_net_values()
        self.assertEqual(total, 15000.0)
        self.assertFalse(truncated)  # row2 空 → 正常终止

    def test_stops_at_first_empty(self):
        tx, _ = _make_create_tx({})  # 全部为空控件
        total, truncated = tx._sum_item_net_values()
        self.assertEqual(total, 0.0)
        self.assertFalse(truncated)

    def test_truncated_when_all_rows_filled(self):
        preset = {
            **_net_row("1000", "100.00", 0),
            **_net_row("2000", "100.00", 1),
            **_net_row("3000", "100.00", 2),
        }
        tx, _ = _make_create_tx(preset)
        total, truncated = tx._sum_item_net_values(max_rows=3)  # 恰好扫满，未遇空行
        self.assertEqual(total, 300.0)
        self.assertTrue(truncated)  # 可能有未计入 item → 告警


class FillOrderValueTest(unittest.TestCase):
    """创建 VA01：fill_order_value。"""

    def test_above_threshold_writes_value(self):
        # 净值 10000 × 汇率 7 = 70000 ≥ 35000 → 写入换算值。
        preset = _net_row("1000", "10000.00", 0)
        tx, raw = _make_create_tx(preset)
        result = tx.fill_order_value(_make_order(rate=7.0))
        self.assertTrue(result.success, result.message)
        self.assertEqual(raw.findById(AUFTRAGSWERT).text, "70000.00")
        self.assertIn("净值 10000.00 × 7.0", result.message)

    def test_below_threshold_no_write(self):
        # 净值 1000 × 汇率 1 = 1000 < 35000 → 不写，字段保持空（SAP 自动带出）。
        preset = _net_row("1000", "1000.00", 0)
        tx, raw = _make_create_tx(preset)
        result = tx.fill_order_value(_make_order(rate=1.0))
        self.assertTrue(result.success, result.message)
        self.assertEqual(raw.findById(AUFTRAGSWERT).text, "")  # 未写入
        self.assertIn("跳过写入", result.message)

    def test_multi_item_summed_before_threshold(self):
        # 两个 item 净值加和 20000 × 汇率 2 = 40000 ≥ 35000 → 写入。
        preset = {**_net_row("1000", "12000.00", 0), **_net_row("2000", "8000.00", 1)}
        tx, raw = _make_create_tx(preset)
        result = tx.fill_order_value(_make_order(rate=2.0))
        self.assertTrue(result.success, result.message)
        self.assertEqual(raw.findById(AUFTRAGSWERT).text, "40000.00")

    def test_truncated_warns_end_to_end(self):
        # 满 200 行触发截断：warning 置真，且告警文本不被阈值分支的 result.message 覆盖。
        preset = {}
        for row in range(200):
            preset.update(_net_row(str(1000 + row), "200.00", row))
        tx, _ = _make_create_tx(preset)
        result = tx.fill_order_value(_make_order(rate=1.0))  # 200×200=40000 ≥ 阈值
        self.assertTrue(result.success, result.message)
        self.assertTrue(result.warning)
        self.assertIn("超过扫描上限", result.message)


class EditOrderValueTest(unittest.TestCase):
    """编辑 VA02：edit_order_value（仅差异才写）。"""

    def test_above_threshold_sets_value(self):
        # 净值 10000 × 7 = 70000；字段原为空 → 写入并记差异。
        preset = _net_row("1000", "10000.00", 0)
        tx, raw = _make_edit_tx(preset)
        diffs: list[str] = []
        result = tx.edit_order_value(_make_order(rate=7.0), diffs)
        self.assertTrue(result.success, result.message)
        self.assertEqual(raw.findById(AUFTRAGSWERT).text, "70000.00")
        self.assertEqual(diffs, ["订单价值:→70000.00"])

    def test_below_threshold_clears_dirty_value(self):
        # 净值 1000 × 1 = 1000 < 35000 → 目标为空；历史双重汇率脏值 35,675.00 被清空自愈。
        preset = {**_net_row("1000", "1000.00", 0), AUFTRAGSWERT: _Element("35,675.00")}
        tx, raw = _make_edit_tx(preset)
        diffs: list[str] = []
        result = tx.edit_order_value(_make_order(rate=1.0), diffs)
        self.assertTrue(result.success, result.message)
        self.assertEqual(raw.findById(AUFTRAGSWERT).text, "")  # 脏值已清空
        self.assertEqual(diffs, ["订单价值:35,675.00→"])

    def test_below_threshold_already_empty_no_diff(self):
        # 净值小额且字段本就为空 → 无差异、不写。
        preset = {**_net_row("1000", "1000.00", 0), AUFTRAGSWERT: _Element("")}
        tx, _ = _make_edit_tx(preset)
        diffs: list[str] = []
        result = tx.edit_order_value(_make_order(rate=1.0), diffs)
        self.assertTrue(result.success, result.message)
        self.assertEqual(diffs, [])
        self.assertEqual(result.message, "订单价值无差异")

    def test_no_diff_when_equal(self):
        # 现值与计算值等价（仅千分位显示差异）→ 归一后相等，不写、不记差异。
        preset = {**_net_row("1000", "10000.00", 0), AUFTRAGSWERT: _Element("70,000.00")}
        tx, _ = _make_edit_tx(preset)
        diffs: list[str] = []
        result = tx.edit_order_value(_make_order(rate=7.0), diffs)
        self.assertTrue(result.success, result.message)
        self.assertEqual(diffs, [])

    def test_truncated_warns_end_to_end(self):
        # 预设满 200 行（默认上限）全部非空 → edit_order_value 应标警告并追加提示，
        # 且不会因阈值分支的差异写入而丢失该提示。
        preset = {AUFTRAGSWERT: _Element("")}
        for row in range(200):
            preset.update(_net_row(str(1000 + row), "200.00", row))
        tx, _ = _make_edit_tx(preset)
        diffs: list[str] = []
        result = tx.edit_order_value(_make_order(rate=1.0), diffs)  # 200×200=40000 ≥ 阈值
        self.assertTrue(result.success, result.message)
        self.assertTrue(result.warning)
        self.assertIn("超过扫描上限", diffs[-1])


if __name__ == "__main__":
    unittest.main()
