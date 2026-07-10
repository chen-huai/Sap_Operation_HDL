"""should_fill_ic_transaction 规则单测（Data_B_TUV → VA01 IC_TRANSAKTION=O1）。

自包含，不依赖 sap/test/helpers.py（其已与当前 models 失步）。
"""

from __future__ import annotations

import sys
from unittest.mock import MagicMock

# 允许在没有真实 pywin32 的环境里导入 sap 模块。
if "win32com" not in sys.modules:
    sys.modules["win32com"] = MagicMock()
    sys.modules["win32com.client"] = MagicMock()

from sap.models import OrderData, SapConfig
from sap.rules import should_fill_ic_transaction


def _make_config(data_b_tuv):
    return SapConfig(
        order_type="ZOR",
        sales_organization="3002",
        distribution_channels="10",
        sales_office="1000",
        sales_group="240",
        sub_cost_center_cs="1101",
        sub_cost_center_chm="1102",
        sub_cost_center_phy="1103",
        cs_code="CS001",
        sales_code="SA001",
        data_b_tuv=data_b_tuv,
    )


def _make_order(sap_no):
    return OrderData(
        sap_no=sap_no,
        project_no="PRJ-001",
        currency_type="CNY",
        exchange_rate=1.0,
        short_text="Test",
    )


class TestShouldFillIcTransaction:
    def test_sap_no_in_data_b_tuv_returns_true(self):
        config = _make_config(["9000220", "9000350"])
        assert should_fill_ic_transaction(_make_order("9000220"), config) is True

    def test_sap_no_not_in_data_b_tuv_returns_false(self):
        config = _make_config(["9000220", "9000350"])
        assert should_fill_ic_transaction(_make_order("123456"), config) is False

    def test_empty_data_b_tuv_returns_false(self):
        config = _make_config([])
        assert should_fill_ic_transaction(_make_order("9000220"), config) is False
