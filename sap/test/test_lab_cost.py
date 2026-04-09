"""
测试 Sap.lab_cost — 填写 Data B 成本中心和人工成本

使用方法:
    pytest sap/test/test_lab_cost.py -v

覆盖分支:
    - A2/D2/D3 物料 → 双部门 (CHM + PHY)
    - T20/430 物料 → 单部门 (PHY)
    - 其他物料 → 单部门 (CHM)
    - 异常处理
"""

import pytest
from unittest.mock import MagicMock

from sap.test.helpers import (
    create_sap_instance, make_order, make_revenue, setup_session_mock,
)


class TestLabCostBranches:
    """lab_cost 三路物料分支"""

    def test_a2_material_dual_department(self):
        """A2 物料 → CHM + PHY 双部门成本中心和费用"""
        session = MagicMock()
        elements = setup_session_mock(session)
        sap = create_sap_instance(mock_session=session)
        order = make_order(material_code='A2-405-00')
        revenue = make_revenue(chm_cost=3000.0, phy_cost=2000.0)

        result = sap.lab_cost(order, revenue)

        assert result.success
        # 验证 CHM 和 PHY 成本中心都被设置
        chm_id = ('wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14'
                   '/ssubSUBSCREEN_BODY:SAPMV45A:4312'
                   '/tblSAPMV45AZULEISTENDE/ctxtTABL-KOSTL[0,0]')
        phy_id = ('wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14'
                   '/ssubSUBSCREEN_BODY:SAPMV45A:4312'
                   '/tblSAPMV45AZULEISTENDE/ctxtTABL-KOSTL[0,1]')
        assert chm_id in elements
        assert phy_id in elements

    def test_d2_material_dual_department(self):
        """D2 物料 → 同 A2，走双部门分支"""
        session = MagicMock()
        elements = setup_session_mock(session)
        sap = create_sap_instance(mock_session=session)
        order = make_order(material_code='D2-405-00')

        result = sap.lab_cost(order, make_revenue())

        assert result.success
        # D2 走 A2/D2/D3 分支，应设置两个成本中心
        phy_cost_id = ('wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14'
                       '/ssubSUBSCREEN_BODY:SAPMV45A:4312'
                       '/tblSAPMV45AKOSTENSAETZE/txtTABD-FESTPREIS[5,1]')
        assert phy_cost_id in elements

    def test_t20_material_single_phy(self):
        """T20 物料 → 单部门 PHY"""
        session = MagicMock()
        elements = setup_session_mock(session)
        sap = create_sap_instance(mock_session=session)
        order = make_order(material_code='T20-430-00')

        result = sap.lab_cost(order, make_revenue())

        assert result.success
        # 只有一行成本中心 (PHY)，不应有第二行
        single_id = ('wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14'
                      '/ssubSUBSCREEN_BODY:SAPMV45A:4312'
                      '/tblSAPMV45AZULEISTENDE/ctxtTABL-KOSTL[0,0]')
        second_id = ('wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14'
                      '/ssubSUBSCREEN_BODY:SAPMV45A:4312'
                      '/tblSAPMV45AZULEISTENDE/ctxtTABL-KOSTL[0,1]')
        assert single_id in elements
        assert second_id not in elements

    def test_other_material_single_chm(self):
        """其他物料 (如 T75) → 单部门 CHM"""
        session = MagicMock()
        elements = setup_session_mock(session)
        sap = create_sap_instance(mock_session=session)
        order = make_order(material_code='T75-405-00')

        result = sap.lab_cost(order, make_revenue())

        assert result.success
        single_id = ('wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14'
                      '/ssubSUBSCREEN_BODY:SAPMV45A:4312'
                      '/tblSAPMV45AZULEISTENDE/ctxtTABL-KOSTL[0,0]')
        assert single_id in elements

    def test_exception_returns_failure(self):
        """session 异常 — 返回失败结果"""
        session = MagicMock()
        session.findById.side_effect = Exception("填写异常")
        sap = create_sap_instance(mock_session=session)

        result = sap.lab_cost(make_order(), make_revenue())

        assert not result.success
        assert 'Data B未填写' in result.message


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
