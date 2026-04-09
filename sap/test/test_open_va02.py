"""
测试 Sap.open_va02 — 打开订单 (VA02 屏幕)

使用方法:
    pytest sap/test/test_open_va02.py -v
"""

import pytest
from unittest.mock import MagicMock

from sap.test.helpers import create_sap_instance, setup_session_mock


class TestOpenVa02:
    """open_va02 打开订单"""

    def test_success(self):
        """正常打开订单 — 事务码和订单号正确设置"""
        session = MagicMock()
        elements = setup_session_mock(session)
        sap = create_sap_instance(mock_session=session)

        result = sap.open_va02('60001234')

        assert result.success
        # 验证事务码
        okcd = elements.get('wnd[0]/tbar[0]/okcd')
        assert okcd is not None
        # 验证订单号
        order_field = elements.get('wnd[0]/usr/ctxtVBAK-VBELN')
        assert order_field is not None

    def test_exception_returns_failure(self):
        """session 异常 — 返回失败结果，包含订单号"""
        session = MagicMock()
        session.findById.side_effect = Exception("无法打开")
        sap = create_sap_instance(mock_session=session)

        result = sap.open_va02('60001234')

        assert not result.success
        assert '60001234' in result.message
        assert '未开启' in result.message


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
