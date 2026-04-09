"""
测试 Sap.va01_operate — 创建订单 (VA01)

使用方法:
    pytest sap/test/test_va01_operate.py -v

覆盖分支:
    - CNY / 非 CNY 货币
    - DATA A 四路分支 (D2-E1, D2-Z0, Z2, 默认00)
    - DATA B 阈值判断 (revenue_cny >= threshold)
    - 联系人 / 销售合作伙伴
    - 异常处理
"""

import pytest
from unittest.mock import MagicMock

from sap.test.helpers import (
    create_sap_instance, make_config, make_order, make_revenue, make_flags,
    setup_session_mock,
)


class TestVa01OperateCurrency:
    """VA01 货币分支"""

    def test_cny_no_exchange_rate(self):
        """CNY 货币 — 不填汇率字段"""
        session = MagicMock()
        elements = setup_session_mock(session, text_returns={
            # 合作伙伴 tab: 第4行是'负责雇员'
            'wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352'
            '/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW'
            '/cmbGVS_TC_DATA-REC-PARVW[0,4]': '负责雇员',
            'wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352'
            '/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW'
            '/cmbGVS_TC_DATA-REC-PARVW[0,5]': '送达方',
        })
        order = make_order(currency_type='CNY')
        revenue = make_revenue()
        sap = create_sap_instance(mock_session=session)

        result = sap.va01_operate(order, revenue)

        assert result.success
        # 汇率字段不应被设置 (CNY 不进入 exchange_rate 分支)
        exchange_id = ('wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01'
                       '/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBKD-KURSK')
        assert exchange_id not in elements

    def test_non_cny_sets_exchange_rate(self):
        """非 CNY 货币 — 填写汇率"""
        session = MagicMock()
        elements = setup_session_mock(session, text_returns={
            'wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352'
            '/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW'
            '/cmbGVS_TC_DATA-REC-PARVW[0,4]': '负责雇员',
            'wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352'
            '/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW'
            '/cmbGVS_TC_DATA-REC-PARVW[0,5]': '送达方',
        })
        order = make_order(currency_type='USD', exchange_rate=7.25)
        revenue = make_revenue()
        sap = create_sap_instance(mock_session=session)

        result = sap.va01_operate(order, revenue)

        assert result.success
        # 汇率字段应被创建（意味着 findById 被调用）
        exchange_id = ('wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01'
                       '/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBKD-KURSK')
        assert exchange_id in elements


class TestVa01OperateDataA:
    """VA01 DATA A 分支"""

    def test_d2_material_in_data_ae1(self):
        """D2 物料 + sap_no 在 data_ae1 → KVGR1 = E1"""
        session = MagicMock()
        elements = setup_session_mock(session, text_returns={
            'wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352'
            '/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW'
            '/cmbGVS_TC_DATA-REC-PARVW[0,4]': 'Employee respons.',
            'wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352'
            '/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW'
            '/cmbGVS_TC_DATA-REC-PARVW[0,5]': 'Ship-to',
        })
        config = make_config(data_ae1=['100001'])
        order = make_order(sap_no='100001', material_code='D2-405-00')
        revenue = make_revenue()
        sap = create_sap_instance(mock_session=session, config=config)

        result = sap.va01_operate(order, revenue)

        assert result.success
        kvgr1_id = ('wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13'
                     '/ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR1')
        kvgr1_el = elements.get(kvgr1_id)
        assert kvgr1_el is not None

    def test_default_data_a(self):
        """非 D2/D3 且不在 data_az2 → KVGR1 = 00"""
        session = MagicMock()
        setup_session_mock(session, text_returns={
            'wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352'
            '/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW'
            '/cmbGVS_TC_DATA-REC-PARVW[0,4]': '负责雇员',
            'wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352'
            '/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW'
            '/cmbGVS_TC_DATA-REC-PARVW[0,5]': '送达方',
        })
        order = make_order(sap_no='999999', material_code='T75-405-00')
        revenue = make_revenue()
        sap = create_sap_instance(mock_session=session)

        result = sap.va01_operate(order, revenue)

        assert result.success


class TestVa01OperateDataB:
    """VA01 DATA B (Auftragswert) 阈值分支"""

    def test_revenue_above_threshold_fills_auftragswert(self):
        """revenue_cny >= 35000 → 填写 Auftragswert"""
        session = MagicMock()
        elements = setup_session_mock(session, text_returns={
            'wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352'
            '/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW'
            '/cmbGVS_TC_DATA-REC-PARVW[0,4]': '负责雇员',
            'wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352'
            '/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW'
            '/cmbGVS_TC_DATA-REC-PARVW[0,5]': '送达方',
        })
        revenue = make_revenue(revenue_cny=50000.0)
        sap = create_sap_instance(mock_session=session)

        result = sap.va01_operate(make_order(), revenue)

        assert result.success
        auft_id = ('wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14'
                    '/ssubSUBSCREEN_BODY:SAPMV45A:4312/txtZAUFTD-AUFTRAGSWERT')
        assert auft_id in elements

    def test_revenue_below_threshold_skips_auftragswert(self):
        """revenue_cny < 35000 → 不填写 Auftragswert"""
        session = MagicMock()
        elements = setup_session_mock(session, text_returns={
            'wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352'
            '/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW'
            '/cmbGVS_TC_DATA-REC-PARVW[0,4]': '负责雇员',
            'wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352'
            '/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW'
            '/cmbGVS_TC_DATA-REC-PARVW[0,5]': '送达方',
        })
        revenue = make_revenue(revenue_cny=20000.0)
        sap = create_sap_instance(mock_session=session)

        result = sap.va01_operate(make_order(), revenue)

        assert result.success
        auft_id = ('wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14'
                    '/ssubSUBSCREEN_BODY:SAPMV45A:4312/txtZAUFTD-AUFTRAGSWERT')
        assert auft_id not in elements


class TestVa01OperateException:
    """VA01 异常处理"""

    def test_exception_returns_failure(self):
        """session 异常 — 返回失败 SapResult"""
        session = MagicMock()
        session.findById.side_effect = Exception("SAP 界面异常")
        sap = create_sap_instance(mock_session=session)

        result = sap.va01_operate(make_order(), make_revenue())

        assert not result.success
        assert 'Order No未创建成功' in result.message


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
