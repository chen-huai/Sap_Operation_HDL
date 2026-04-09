"""
测试 Sap.end_sap — 结束 SAP 连接

使用方法:
    pytest sap/test/test_end_sap.py -v
"""

import pytest

from sap.test.helpers import create_sap_instance


class TestEndSap:
    """end_sap 连接清理"""

    def test_all_attributes_set_to_none(self):
        """调用后 session/connection/application/SapGuiAuto 全部置空"""
        sap = create_sap_instance()

        assert sap.session is not None
        assert sap.connection is not None

        sap.end_sap()

        assert sap.session is None
        assert sap.connection is None
        assert sap.application is None
        assert sap.SapGuiAuto is None


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
