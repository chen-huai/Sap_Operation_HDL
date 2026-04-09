"""
集成测试: end_sap 结束 SAP 连接

注意: 运行此测试后 SAP 连接将断开，其他测试无法继续使用

运行: pytest sap/test/integration/test_end_sap.py -v -s
"""

import pytest
from sap.function import Sap
from sap.test.integration.test_config import SAP_CONFIG, FLAGS


class TestEndSap:
    def test_end_sap(self):
        """结束 SAP 连接 — 所有属性置空"""
        # 创建独立连接（不使用 sap_live fixture，避免影响其他测试）
        sap = Sap(config=SAP_CONFIG, flags=FLAGS)
        if not sap.result.success:
            pytest.skip(f'SAP 未连接: {sap.result.message}')

        sap.end_sap()

        assert sap.session is None
        assert sap.connection is None
        assert sap.application is None
        assert sap.SapGuiAuto is None
        print('SAP 连接已断开')


if __name__ == '__main__':
    pytest.main([__file__, '-v', '-s'])
