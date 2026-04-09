"""
集成测试: lab_cost 填写 Data B 成本

前提条件:
    - SAP GUI 已登录
    - 当前 SAP 屏幕处于订单编辑状态（通常在 VA01 之后）
    - test_config.py: ORDER, REVENUE 已填写

运行: pytest sap/test/integration/test_lab_cost.py -v -s
"""

import pytest


class TestLabCost:
    def test_lab_cost(self, sap_live, order, revenue):
        """填写 Data B 成本中心和人工成本"""
        result = sap_live.lab_cost(order, revenue)
        assert result.success, f'lab_cost 失败: {result.message}'
        print(f'lab_cost 执行成功')


if __name__ == '__main__':
    pytest.main([__file__, '-v', '-s'])
