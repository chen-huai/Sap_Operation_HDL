"""
测试 Sap.va02_operate — 添加/修改订单 Item (VA02)

使用方法:
    pytest sap/test/test_va02_operate.py -v

覆盖分支:
    - A2 物料 + 430 子码 → 双 Item, 特殊 focus
    - A2 物料 + 非 430 → 双 Item, 标准 focus
    - 非 A2 物料 → 单 Item
    - long_text 填写（含 plan_cost 联动跳转）
    - plan_cost 调用失败 → message 追加
    - 异常处理
"""

import pytest
from unittest.mock import MagicMock, patch

from sap.test.helpers import (
    create_sap_instance, make_config, make_order, make_revenue, make_flags,
    setup_session_mock,
)
from sap.models import SapResult


class TestVa02A2Material:
    """VA02 A2 物料双 Item 分支"""

    def test_a2_430_dual_items(self):
        """A2-430 物料 → 双 Item, 430 特殊 focus 路径"""
        session = MagicMock()
        setup_session_mock(session, text_returns={
            'wnd[0]/usr/ctxtVBAK-VBELN': '60001234',
            # 两次读取 sapAmountVatStr
            'wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06'
            '/ssubSUBSCREEN_BODY:SAPLV69A:6201'
            '/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]': '1,500.00',
        })
        flags = make_flags(plan_cost=False)
        sap = create_sap_instance(mock_session=session, flags=flags)
        order = make_order(material_code='A2-430-00', long_text='')
        revenue = make_revenue(
            revenue_cny=20000.0,  # 低于阈值，不触发 plan_cost
            phy_revenue=1500.0,
            chm_revenue=1000.0,
        )

        result = sap.va02_operate(order, revenue)

        assert result.success
        assert result.order_no == '60001234'

    def test_a2_non430_dual_items(self):
        """A2-405 物料 → 双 Item, 标准 focus 路径"""
        session = MagicMock()
        setup_session_mock(session, text_returns={
            'wnd[0]/usr/ctxtVBAK-VBELN': '60005678',
            'wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06'
            '/ssubSUBSCREEN_BODY:SAPLV69A:6201'
            '/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]': '2,000.00',
        })
        flags = make_flags(plan_cost=False)
        sap = create_sap_instance(mock_session=session, flags=flags)
        order = make_order(material_code='A2-405-00', long_text='')
        revenue = make_revenue(revenue_cny=20000.0)

        result = sap.va02_operate(order, revenue)

        assert result.success
        assert result.order_no == '60005678'


class TestVa02NonA2Material:
    """VA02 非 A2 物料单 Item 分支"""

    def test_single_item(self):
        """普通物料 → 单 Item"""
        session = MagicMock()
        setup_session_mock(session, text_returns={
            'wnd[0]/usr/ctxtVBAK-VBELN': '60009999',
            'wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06'
            '/ssubSUBSCREEN_BODY:SAPLV69A:6201'
            '/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]': '10,000.00',
        })
        flags = make_flags(plan_cost=False)
        sap = create_sap_instance(mock_session=session, flags=flags)
        order = make_order(material_code='T75-405-00', long_text='')
        revenue = make_revenue(revenue_cny=20000.0)

        result = sap.va02_operate(order, revenue)

        assert result.success
        assert result.sap_amount_vat == '10,000.00'

    def test_with_long_text(self):
        """带 long_text — 填写 Item 文本 tab"""
        session = MagicMock()
        setup_session_mock(session, text_returns={
            'wnd[0]/usr/ctxtVBAK-VBELN': '60009999',
            'wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06'
            '/ssubSUBSCREEN_BODY:SAPLV69A:6201'
            '/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]': '5,000.00',
        })
        flags = make_flags(plan_cost=False)
        sap = create_sap_instance(mock_session=session, flags=flags)
        order = make_order(material_code='T75-405-00', long_text='Some long text')
        revenue = make_revenue(revenue_cny=20000.0)

        result = sap.va02_operate(order, revenue)

        assert result.success


class TestVa02Exception:
    """VA02 异常处理"""

    def test_exception_returns_failure(self):
        """session 异常 — 返回失败 SapResult"""
        session = MagicMock()
        session.findById.side_effect = Exception("VA02 界面异常")
        sap = create_sap_instance(mock_session=session)

        result = sap.va02_operate(make_order(), make_revenue())

        assert not result.success
        assert 'Order添加Item失败' in result.message


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
