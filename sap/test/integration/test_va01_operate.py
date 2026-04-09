"""
集成测试: VA01 创建订单

前提条件:
    - SAP GUI 已登录
    - test_config.py: ORDER, REVENUE 已填写

运行: pytest sap/test/integration/test_va01_operate.py -v -s
"""

import pytest


class TestVa01Operate:
    def test_va01_create_order(self, sap_live, order, revenue):
        """VA01 创建订单 — 验证 SAP 操作无报错"""
        result = sap_live.va01_operate(order, revenue)
        assert result.success, f'VA01 失败: {result.message}'
        print(f'VA01 执行成功')


if __name__ == '__main__':
    pytest.main([__file__, '-v', '-s'])
