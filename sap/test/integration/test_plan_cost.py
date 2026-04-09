"""
集成测试: plan_cost 填写成本计划

前提条件:
    - SAP GUI 已登录
    - 当前 SAP 屏幕处于订单 Item 编辑状态（通常在 VA02 之后）
    - test_config.py: ORDER, REVENUE 已填写
    - FLAGS.plan_cost = True 或 REVENUE.revenue_cny >= 阈值

运行: pytest sap/test/integration/test_plan_cost.py -v -s
"""

import pytest


class TestPlanCost:
    def test_plan_cost(self, sap_live, order, revenue):
        """填写 Plan Cost"""
        result = sap_live.plan_cost(order, revenue)
        assert result.success, f'plan_cost 失败: {result.message}'
        print(f'plan_cost 执行成功')


if __name__ == '__main__':
    pytest.main([__file__, '-v', '-s'])
