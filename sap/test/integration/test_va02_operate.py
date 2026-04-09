"""
集成测试: VA02 添加/修改订单 Item

前提条件:
    - SAP GUI 已登录
    - test_config.py: ORDER, REVENUE, EXISTING_ORDER_NO 已填写
    - 或者: 刚执行完 VA01（SAP 屏幕保留了订单号）

运行: pytest sap/test/integration/test_va02_operate.py -v -s
"""

import pytest


class TestVa02Operate:
    def test_va02_add_item_after_va01(self, sap_live, order, revenue):
        """VA02 添加 Item（紧接 VA01 之后运行，SAP 屏幕已有订单号）"""
        result = sap_live.va02_operate(order, revenue)
        assert result.success, f'VA02 失败: {result.message}'
        print(f'VA02 执行成功, 订单号: {result.order_no}, '
              f'SAP 含税金额: {result.sap_amount_vat}')

    def test_va02_add_item_with_existing_order(self, sap_live, existing_order_no, order, revenue):
        """VA02 添加 Item（先打开已有订单）"""
        # 前置: 加载订单到 VA02 屏幕
        open_result = sap_live.open_va02(existing_order_no)
        assert open_result.success, f'打开订单失败: {open_result.message}'

        result = sap_live.va02_operate(order, revenue)
        assert result.success, f'VA02 失败: {result.message}'
        print(f'VA02 执行成功, 订单号: {result.order_no}, '
              f'SAP 含税金额: {result.sap_amount_vat}')


if __name__ == '__main__':
    pytest.main([__file__, '-v', '-s'])
