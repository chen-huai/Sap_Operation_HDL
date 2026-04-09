"""
集成测试: open_va02 打开订单

前提条件:
    - SAP GUI 已登录
    - test_config.py: EXISTING_ORDER_NO 已填写

运行: pytest sap/test/integration/test_open_va02.py -v -s
"""

import pytest


class TestOpenVa02:
    def test_open_order(self, sap_live, existing_order_no):
        """打开已有订单到 VA02 屏幕"""
        result = sap_live.open_va02(existing_order_no)
        assert result.success, f'打开订单失败: {result.message}'
        print(f'打开订单 {existing_order_no} 成功')


if __name__ == '__main__':
    pytest.main([__file__, '-v', '-s'])
