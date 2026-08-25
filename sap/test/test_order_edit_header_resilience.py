"""VA02 编辑：抬头段的弹窗排空与分段独立容错回归测试。

覆盖 2026-08-25 实机故障：改售达方后 SAP 连弹十余个确认框，`_MAX_DIALOG_ROUNDS=8`
排不完 → 残留模态窗让 press(btnBT_HEAD) 不生效（不报错）→ select_tab(tabpT\\01) 报
"找不到 SAP 元素" → 原 edit_header 的大 try 把币种之后的**五段一起吞掉** →
buyer/CS/Sales 从未更新。

两条修复各自的回归点：
    ① 排空上限给足 + 跑满不再静默（`_confirm_sold_to_dialogs` 返回 bool）；
    ② 详情屏各段独立容错——某段抛异常不阻断其余段，伙伴段照样执行。
"""

from __future__ import annotations

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

import sap.test.helpers  # noqa: E402,F401  仅为注入 win32com mock
from sap.models import OrderData, SapConfig  # noqa: E402
from sap.session import SapSession, SapUiError  # noqa: E402
from sap.transactions.order_edit import OrderEditTransaction  # noqa: E402

HEAD_TAB_T01 = "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01"
BT_HEAD = "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD"
POPUP_WINDOW = "wnd[1]"
VAROPTION1 = "wnd[1]/usr/btnSPOP-VAROPTION1"


def _make_config() -> SapConfig:
    return SapConfig(
        order_type="ZOR",
        sales_organization="3002",
        distribution_channels="10",
        sales_office="1000",
        sales_group="200",
        sub_cost_center_cs="1101",
        sub_cost_center_chm="1102",
        sub_cost_center_phy="1103",
        cs_code="CS001",
        sales_code="SA001",
    )


class _Element:
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


class _PopupSession:
    """带弹窗计数的 raw session 桩。

    `popups` 条待排空弹窗：每次访问 wnd[1] 系控件消耗一条，耗尽后 wnd[1] 抛异常
    （模拟无弹窗）。弹窗未耗尽期间 `detail_ready` 为假——模拟"模态窗挡住 btnBT_HEAD
    导致进不了抬头详情屏"，`tabpT\\01` 也随之不可访问。
    """

    def __init__(self, popups: int = 0, *, detail_after_clear: bool = True):
        self.remaining = popups
        self._detail_after_clear = detail_after_clear
        self._cache: dict[str, _Element] = {}
        self.press_bt_head_calls = 0

    def findById(self, element_id: str):
        if element_id.startswith(POPUP_WINDOW):
            if self.remaining <= 0:
                raise Exception(f"popup gone: {element_id}")
            self.remaining -= 1
            return _Element()
        if element_id == HEAD_TAB_T01:
            # 弹窗未排空 → 详情屏进不去；排空后按 detail_after_clear 决定是否就绪。
            if self.remaining > 0 or not self._detail_after_clear:
                raise Exception("not on header detail screen")
            return self._element(element_id)
        if element_id == BT_HEAD:
            self.press_bt_head_calls += 1
        return self._element(element_id)

    def _element(self, element_id: str) -> _Element:
        if element_id not in self._cache:
            self._cache[element_id] = _Element()
        return self._cache[element_id]


def _make_transaction(raw):
    session = SapSession(raw, raw, raw, raw)
    return OrderEditTransaction(session, _make_config())


def _make_order() -> OrderData:
    return OrderData(
        sap_no="654321",
        project_no="PRJ-001",
        currency_type="CNY",
        exchange_rate=1.0,
        short_text="Test",
        global_partner_code="GP001",
    )


class ConfirmSoldToDialogsTest(unittest.TestCase):
    def test_clears_more_than_eight_popups(self):
        # 原上限 8 的实机故障场景：12 个弹窗必须全部排空。
        raw = _PopupSession(popups=12)
        tx = _make_transaction(raw)
        self.assertTrue(tx._confirm_sold_to_dialogs())
        self.assertEqual(raw.remaining, 0)

    def test_no_popup_returns_immediately(self):
        raw = _PopupSession(popups=0)
        tx = _make_transaction(raw)
        self.assertTrue(tx._confirm_sold_to_dialogs())

    def test_exhausting_rounds_reports_false(self):
        # 弹窗数超过兜底上限 → 返回 False，供调用方记「未排空」而非静默放过。
        raw = _PopupSession(popups=OrderEditTransaction._MAX_DIALOG_ROUNDS + 5)
        tx = _make_transaction(raw)
        self.assertFalse(tx._confirm_sold_to_dialogs())

    def test_raw_com_error_does_not_escape(self):
        """press()/sendVKey() 抛的原始 COM 异常不得穿透排空循环。

        注意穿透点在**控件方法调用**而非查找：session.find() 的 except Exception 已把
        findById 的异常统一包装成 SapUiError，真正会漏的是 find 成功之后的 `.press()`
        / `.sendVKey()`——它们不在 find 的 try 内。故桩要让控件方法抛，而非 findById 抛。
        """
        class _ComErrorElement(_Element):
            def press(self):
                raise RuntimeError("COM error 0x80004005: control disabled")

            def sendVKey(self, _key):
                raise RuntimeError("COM error 0x80004005: window busy")

        class _ComErrorSession(_PopupSession):
            def findById(self, element_id):
                if element_id.startswith(POPUP_WINDOW):
                    if self.remaining <= 0:
                        raise Exception("popup gone")
                    self.remaining -= 1
                    return _ComErrorElement()
                return super().findById(element_id)

        tx = _make_transaction(_ComErrorSession(popups=3))
        # 每轮两次控件调用各抛一次 COM 异常 → 都被吞掉、循环判定为"无弹窗"正常退出。
        self.assertTrue(tx._confirm_sold_to_dialogs())


class EnterHeaderDetailTest(unittest.TestCase):
    def test_enters_after_clearing_popups(self):
        raw = _PopupSession(popups=12)
        tx = _make_transaction(raw)
        diffs: list[str] = []
        self.assertTrue(tx._enter_header_detail(diffs))
        self.assertEqual(diffs, [])

    def test_records_failure_when_screen_never_ready(self):
        # 弹窗排空了但仍进不去详情屏 → 不抛异常，记明确原因返回 False。
        raw = _PopupSession(popups=0, detail_after_clear=False)
        tx = _make_transaction(raw)
        diffs: list[str] = []
        self.assertFalse(tx._enter_header_detail(diffs))
        self.assertEqual(diffs, ["抬头详情屏:进入失败(弹窗未排空或屏态异常)"])
        # 重试轮数用尽，press 次数等于上限（每轮一次）。
        self.assertEqual(raw.press_bt_head_calls, OrderEditTransaction._MAX_HEADER_ENTRY_RETRIES)


class EditHeaderSegmentIsolationTest(unittest.TestCase):
    """一段异常不得阻断其余段——本次故障的核心回归点。"""

    def test_currency_failure_does_not_block_partners(self):
        raw = _PopupSession(popups=0)
        tx = _make_transaction(raw)

        calls: list[str] = []

        def _boom(_order, _diffs):
            calls.append("币种")
            raise SapUiError("找不到 SAP 元素: tabpT\\01")

        def _record(name):
            def _handler(_order, _diffs):
                calls.append(name)
            return _handler

        tx._edit_currency = _boom
        tx._edit_partners = _record("伙伴")
        tx._edit_short_text = _record("售达方文本")
        tx._edit_submission = _record("Submission")
        tx._edit_data_a = _record("Data A")
        tx._edit_data_b_header = _record("Data B抬头")

        diffs: list[str] = []
        result = tx.edit_header(_make_order(), diffs)

        # 币种炸了，后面五段必须照样跑完。
        self.assertEqual(
            calls, ["币种", "伙伴", "售达方文本", "Submission", "Data A", "Data B抬头"]
        )
        # 不判 fail（其余段已生效），但标 warning 且 message 点出失败段。
        self.assertTrue(result.success)
        self.assertTrue(result.warning)
        self.assertIn("失败段: 币种:", result.message)

    def test_all_segments_ok_is_not_warning(self):
        raw = _PopupSession(popups=0)
        tx = _make_transaction(raw)
        for name in (
            "_edit_currency", "_edit_partners", "_edit_short_text",
            "_edit_submission", "_edit_data_a", "_edit_data_b_header",
        ):
            setattr(tx, name, lambda _order, _diffs: None)

        result = tx.edit_header(_make_order(), [])
        self.assertTrue(result.success)
        self.assertFalse(result.warning)

    def test_detail_screen_unreachable_fails_whole_step(self):
        # 详情屏进不去 → 所有段都无控件可操作，此时才该整段判失败。
        raw = _PopupSession(popups=0, detail_after_clear=False)
        tx = _make_transaction(raw)
        called: list[str] = []
        tx._edit_partners = lambda _order, _diffs: called.append("伙伴")

        result = tx.edit_header(_make_order(), [])
        self.assertFalse(result.success)
        self.assertIn("抬头详情屏进入失败", result.message)
        self.assertEqual(called, [])

    def test_sold_to_failure_still_runs_detail_segments(self):
        # 售达方段异常也不该阻断详情屏各段（例如控件 ID 变更）。
        raw = _PopupSession(popups=0)
        tx = _make_transaction(raw)
        called: list[str] = []

        def _boom(_order, _diffs):
            raise SapUiError("sold-to control missing")

        tx._edit_sold_to = _boom
        for name in (
            "_edit_currency", "_edit_short_text", "_edit_submission",
            "_edit_data_a", "_edit_data_b_header",
        ):
            setattr(tx, name, lambda _order, _diffs: None)
        tx._edit_partners = lambda _order, _diffs: called.append("伙伴")

        result = tx.edit_header(_make_order(), [])
        self.assertEqual(called, ["伙伴"])
        self.assertTrue(result.warning)
        self.assertIn("售达方:", result.message)


class DataBHeaderReadOnlyTest(unittest.TestCase):
    """Data B 抬头两项走字段级容错：一个只读不得连累另一个。

    实测（2026-08-25 log: `失败段: Data B抬头:Property '.text' can not be set.`）：
    SAP 按客户属性把 IC 交易类型等字段设为只读，写入抛非 SapUiError 的属性错误。
    """

    T14_BASE = (
        "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312"
    )
    ECD_ID = f"{T14_BASE}/ctxtZAUFTD-VORAUS_AUFENDE"
    IC_ID = f"{T14_BASE}/ctxtZAUFTD-IC_TRANSAKTION"

    class _ReadOnlyElement(_Element):
        """只读控件：可读，写 text 抛 SAP 的属性错误。"""

        def __init__(self, text: str = ""):
            object.__setattr__(self, "_text", text)
            self.key = ""
            self.caretPosition = 0

        @property
        def text(self) -> str:
            return self._text

        @text.setter
        def text(self, _value):
            raise RuntimeError("Property '.text' can not be set.")

    def _make_session(self, readonly_ids: set[str], preset_text: dict[str, str]):
        readonly = readonly_ids
        texts = preset_text
        outer = self

        class _Session(_PopupSession):
            def findById(self, element_id):
                if element_id in readonly:
                    if element_id not in self._cache:
                        self._cache[element_id] = outer._ReadOnlyElement(
                            texts.get(element_id, "")
                        )
                    return self._cache[element_id]
                # 初值只在控件首次创建时注入——每次 findById 都填会把写入的空值覆盖回去。
                first_time = element_id not in self._cache
                elem = super().findById(element_id)
                if first_time and element_id in texts:
                    elem.text = texts[element_id]
                return elem

        return _Session(popups=0)

    def test_readonly_ic_does_not_block_ecd(self):
        # IC 只读、ECD 可写 → ECD 正常同步，IC 记「只读无法修改」，不抛异常。
        raw = self._make_session({self.IC_ID}, {self.IC_ID: "O1"})
        tx = _make_transaction(raw)
        order = _make_order()
        order.ecd = "20261231"

        diffs: list[str] = []
        tx._edit_data_b_header(order, diffs)

        self.assertEqual(raw.findById(self.ECD_ID).text, "20261231")
        self.assertEqual(len(diffs), 2)
        self.assertEqual(diffs[0], "ECD:(空)→20261231")
        self.assertIn("IC交易类型:只读无法修改", diffs[1])

    def test_readonly_ecd_does_not_block_ic(self):
        # ECD 只读 → 不阻断 IC 段（原实现第一个字段抛异常就整段结束）。
        raw = self._make_session({self.ECD_ID}, {})
        tx = _make_transaction(raw)
        order = _make_order()
        order.ecd = "20261231"

        diffs: list[str] = []
        tx._edit_data_b_header(order, diffs)

        self.assertIn("ECD:只读无法修改", diffs[0])
        # IC 目标为空、现值也为空 → 无差异不写，故只有 ECD 一条 diff。
        self.assertEqual(len(diffs), 1)

    def test_both_writable_records_both(self):
        raw = self._make_session(set(), {self.IC_ID: "O1"})
        tx = _make_transaction(raw)
        order = _make_order()
        order.ecd = "20261231"

        diffs: list[str] = []
        tx._edit_data_b_header(order, diffs)

        self.assertEqual(diffs, ["ECD:(空)→20261231", "IC交易类型:O1→(空)"])
        self.assertEqual(raw.findById(self.IC_ID).text, "")


if __name__ == "__main__":
    unittest.main()
