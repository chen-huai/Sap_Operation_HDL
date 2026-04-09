"""
集成测试: unlock_or_lock_order 解锁/锁定订单

前提条件:
    - SAP GUI 已登录
    - test_config.py: EXISTING_ORDER_NO 已填写
    - 订单已通过 open_va02 打开

运行: pytest sap/test/integration/test_unlock_or_lock_order.py -v -s
"""

import pytest


class TestUnlockOrLockOrder:
    def test_unlock(self, sap_live, existing_order_no):
        """解锁订单（先打开订单）"""
        open_result = sap_live.open_va02(existing_order_no)
        assert open_result.success, f'打开订单失败: {open_result.message}'

        result = sap_live.unlock_or_lock_order('Unlock')
        assert result.success, f'Unlock 失败: {result.message}'
        print(f'订单 {existing_order_no} 解锁成功')

    def test_lock(self, sap_live, existing_order_no):
        """锁定订单（先打开订单）"""
        open_result = sap_live.open_va02(existing_order_no)
        assert open_result.success, f'打开订单失败: {open_result.message}'

        result = sap_live.unlock_or_lock_order('Lock')
        assert result.success, f'Lock 失败: {result.message}'
        print(f'订单 {existing_order_no} 锁定成功')


if __name__ == '__main__':
    pytest.main([__file__, '-v', '-s'])
