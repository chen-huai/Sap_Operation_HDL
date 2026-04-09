"""
测试 Sap.vf03_operate — 查看形式发票 (VF03)

使用方法:
    pytest sap/test/test_vf03_operate.py -v
"""

import pytest
from unittest.mock import MagicMock

from sap.test.helpers import create_sap_instance, setup_session_mock


class TestVf03Operate:
    """vf03_operate 形式发票查看"""

    def test_success_returns_proforma_no(self):
        """正常执行 — 返回 proforma_no"""
        session = MagicMock()
        setup_session_mock(session, text_returns={
            'wnd[0]/usr/ctxtVBRK-VBELN': '9000012345',
        })
        sap = create_sap_instance(mock_session=session)

        result = sap.vf03_operate()

        assert result.success
        assert result.proforma_no == '9000012345'

    def test_exception_returns_failure(self):
        """session 异常 — 返回失败结果"""
        session = MagicMock()
        session.findById.side_effect = Exception("VF03 界面异常")
        sap = create_sap_instance(mock_session=session)

        result = sap.vf03_operate()

        assert not result.success
        assert '形式发票查看失败' in result.message


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
