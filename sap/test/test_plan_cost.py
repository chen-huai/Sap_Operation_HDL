"""
测试 Sap.plan_cost — 填写成本计划 (Plan Cost)

使用方法:
    pytest sap/test/test_plan_cost.py -v

覆盖分支:
    - 未触发 (plan_cost=False 且 revenue_cny < threshold) → 直接返回成功
    - D2/D3 物料路径 (CS + CHM LAB + PHY LAB + FREMDL)
    - A2 物料 + 430 子码
    - A2 物料 + 非 430 子码
    - 通用物料 (T75/T20 等)
    - 异常处理
"""

import pytest
from unittest.mock import MagicMock

from sap.test.helpers import (
    create_sap_instance, make_order, make_revenue, make_flags,
    setup_session_mock,
)


class TestPlanCostNotTriggered:
    """plan_cost 未触发"""

    def test_flags_off_and_below_threshold(self):
        """plan_cost=False 且 revenue_cny < threshold → 不执行，返回成功"""
        session = MagicMock()
        sap = create_sap_instance(mock_session=session, flags=make_flags(plan_cost=False))
        revenue = make_revenue(revenue_cny=20000.0)  # 低于 35000 阈值

        result = sap.plan_cost(make_order(), revenue)

        assert result.success
        # 不应有任何 session 操作
        session.findById.assert_not_called()


class TestPlanCostD2D3:
    """plan_cost D2/D3 物料路径"""

    def test_d2_full_path(self):
        """D2 物料 — CS + CHM LAB + PHY LAB + FREMDL 全路径"""
        session = MagicMock()
        setup_session_mock(session)
        flags = make_flags(plan_cost=True, cs=True, chm=True, phy=True)
        sap = create_sap_instance(mock_session=session, flags=flags)
        order = make_order(material_code='D2-405-00', cost=3000.0)
        revenue = make_revenue(
            revenue_cny=50000.0,
            chm_cs_cost=1000.0,
            phy_cs_cost=800.0,
            chm_lab_cost=2000.0,
            phy_lab_cost=1500.0,
        )

        result = sap.plan_cost(order, revenue)

        assert result.success

    def test_d3_cs_only(self):
        """D3 物料 — 仅 CS 标志开启"""
        session = MagicMock()
        setup_session_mock(session)
        flags = make_flags(plan_cost=True, cs=True, chm=False, phy=False)
        sap = create_sap_instance(mock_session=session, flags=flags)
        order = make_order(material_code='D3-405-00', cost=0.0)
        revenue = make_revenue(
            revenue_cny=50000.0,
            chm_cs_cost=500.0,
            phy_cs_cost=500.0,
        )

        result = sap.plan_cost(order, revenue)

        assert result.success


class TestPlanCostA2:
    """plan_cost A2 物料路径"""

    def test_a2_430_subcode(self):
        """A2-430 物料 — 430 特殊路径，含 FREMDL"""
        session = MagicMock()
        setup_session_mock(session)
        flags = make_flags(plan_cost=True, cs=True, chm=True, phy=True)
        sap = create_sap_instance(mock_session=session, flags=flags)
        order = make_order(material_code='A2-430-00', cost=2000.0)
        revenue = make_revenue(
            revenue_cny=50000.0,
            phy_cs_cost=800.0,
            phy_lab_cost=1500.0,
            chm_cs_cost=600.0,
            chm_lab_cost=1200.0,
        )

        result = sap.plan_cost(order, revenue)

        assert result.success

    def test_a2_non430_subcode(self):
        """A2-405 物料 — 非 430 路径，含 FREMDL"""
        session = MagicMock()
        setup_session_mock(session)
        flags = make_flags(plan_cost=True, cs=True, chm=True, phy=True)
        sap = create_sap_instance(mock_session=session, flags=flags)
        order = make_order(material_code='A2-405-00', cost=1500.0)
        revenue = make_revenue(
            revenue_cny=50000.0,
            phy_cs_cost=800.0,
            phy_lab_cost=1500.0,
            chm_cs_cost=600.0,
            chm_lab_cost=1200.0,
        )

        result = sap.plan_cost(order, revenue)

        assert result.success


class TestPlanCostGeneral:
    """plan_cost 通用物料路径"""

    def test_t75_material(self):
        """T75 物料 — 通用路径，CHM 成本中心"""
        session = MagicMock()
        setup_session_mock(session)
        flags = make_flags(plan_cost=True, cs=True, chm=True, phy=True)
        sap = create_sap_instance(mock_session=session, flags=flags)
        order = make_order(material_code='T75-405-00', cost=1000.0)
        revenue = make_revenue(
            revenue_cny=50000.0,
            cs_cost=1800.0,
            lab_cost=3500.0,
        )

        result = sap.plan_cost(order, revenue)

        assert result.success

    def test_t20_material(self):
        """T20 物料 — 通用路径，PHY 成本中心"""
        session = MagicMock()
        setup_session_mock(session)
        flags = make_flags(plan_cost=True, cs=True, chm=False, phy=True)
        sap = create_sap_instance(mock_session=session, flags=flags)
        order = make_order(material_code='T20-441-00', cost=500.0)
        revenue = make_revenue(
            revenue_cny=50000.0,
            cs_cost=1000.0,
            lab_cost=2000.0,
        )

        result = sap.plan_cost(order, revenue)

        assert result.success


class TestPlanCostException:
    """plan_cost 异常处理"""

    def test_exception_returns_failure(self):
        """session 异常 — 返回失败结果"""
        session = MagicMock()
        session.findById.side_effect = Exception("Plan Cost 界面异常")
        flags = make_flags(plan_cost=True)
        sap = create_sap_instance(mock_session=session, flags=flags)

        result = sap.plan_cost(make_order(), make_revenue())

        assert not result.success
        assert 'plan cost未添加成功' in result.message


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
