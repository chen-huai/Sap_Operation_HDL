"""
集成测试: SAP 连接验证

前提条件: SAP GUI 已登录
运行: pytest sap/test/integration/test_connection.py -v -s
"""

import pytest


class TestConnection:
    def test_sap_connected(self, sap_live):
        """验证 SAP GUI 连接正常"""
        assert sap_live.result.success
        assert sap_live.session is not None
        print(f'SAP 连接成功, session = {sap_live.session}')


if __name__ == '__main__':
    pytest.main([__file__, '-v', '-s'])
