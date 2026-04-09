"""
测试 Sap.login_hour_gui — 登录工时 GUI 系统 (ZRU1)

使用方法:
    pytest sap/test/test_login_hour_gui.py -v

覆盖分支:
    - 登录成功
    - 员工 ID 无效 (3 种文本: "doesn't exist", "does not exist", "不存在")
    - 异常处理
"""

import pytest
from unittest.mock import MagicMock, patch

from sap.test.helpers import create_sap_instance, make_hour, setup_session_mock


class TestLoginHourGuiSuccess:
    """login_hour_gui 成功"""

    @patch('sap.function.time.sleep')
    def test_login_success(self, mock_sleep):
        """正常登录 — 状态栏无错误关键字"""
        session = MagicMock()
        setup_session_mock(session, text_returns={
            'wnd[0]/sbar/pane[0]': '就绪',
        })
        sap = create_sap_instance(mock_session=session)
        hour = make_hour()

        result = sap.login_hour_gui(hour)

        assert result.success
        mock_sleep.assert_called_once_with(1)


class TestLoginHourGuiInvalidEmployee:
    """login_hour_gui 员工 ID 无效"""

    @patch('sap.function.time.sleep')
    def test_doesnt_exist(self, mock_sleep):
        """状态栏含 "doesn't exist" → 失败"""
        session = MagicMock()
        setup_session_mock(session, text_returns={
            'wnd[0]/sbar/pane[0]': "Personnel number doesn't exist",
        })
        sap = create_sap_instance(mock_session=session)

        result = sap.login_hour_gui(make_hour())

        assert not result.success
        assert '员工ID无效' in result.message

    @patch('sap.function.time.sleep')
    def test_does_not_exist(self, mock_sleep):
        """状态栏含 "does not exist" → 失败"""
        session = MagicMock()
        setup_session_mock(session, text_returns={
            'wnd[0]/sbar/pane[0]': 'Personnel number does not exist',
        })
        sap = create_sap_instance(mock_session=session)

        result = sap.login_hour_gui(make_hour())

        assert not result.success

    @patch('sap.function.time.sleep')
    def test_chinese_not_exist(self, mock_sleep):
        """状态栏含 '不存在' → 失败"""
        session = MagicMock()
        setup_session_mock(session, text_returns={
            'wnd[0]/sbar/pane[0]': '人员编号不存在',
        })
        sap = create_sap_instance(mock_session=session)

        result = sap.login_hour_gui(make_hour())

        assert not result.success


class TestLoginHourGuiException:
    """login_hour_gui 异常"""

    def test_exception_returns_failure(self):
        """session 异常 — 返回失败结果"""
        session = MagicMock()
        session.findById.side_effect = Exception("ZRU1 界面异常")
        sap = create_sap_instance(mock_session=session)

        result = sap.login_hour_gui(make_hour())

        assert not result.success
        assert 'Hour界面失败' in result.message


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
