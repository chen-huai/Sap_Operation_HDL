"""
测试 Sap.__init__ — SAP GUI 连接初始化

使用方法:
    pytest sap/test/test_init.py -v
"""

import pytest
from unittest.mock import patch, MagicMock

from sap.test.helpers import make_config, make_flags
from sap.function import Sap


class _FakeDispatch:
    """用于 type() == CDispatch 检查的假 COM 对象"""
    pass


def _build_fake_chain():
    """构建 SapGuiAuto → application → connection → session 完整链"""
    fake_gui = _FakeDispatch()
    fake_app = _FakeDispatch()
    fake_conn = _FakeDispatch()
    fake_sess = _FakeDispatch()

    fake_gui.GetScriptingEngine = fake_app
    fake_app.Children = lambda idx: fake_conn
    fake_conn.Children = lambda idx: fake_sess

    return fake_gui, fake_sess


class TestSapInit:
    """Sap.__init__ 连接初始化"""

    @patch('sap.function.win32com.client')
    def test_connection_success(self, mock_client):
        """全链路连接成功 — session 可用"""
        fake_gui, fake_sess = _build_fake_chain()
        mock_client.GetObject.return_value = fake_gui
        mock_client.CDispatch = _FakeDispatch

        sap = Sap(make_config(), make_flags())

        assert sap.result.success
        assert sap.session is fake_sess
        mock_client.GetObject.assert_called_once_with("SAPGUI")

    @patch('sap.function.win32com.client')
    def test_sapgui_type_check_fail(self, mock_client):
        """SapGuiAuto 类型检查失败 — 提前返回, 无 application"""
        mock_client.GetObject.return_value = "not_dispatch"
        mock_client.CDispatch = _FakeDispatch

        sap = Sap(make_config(), make_flags())

        # type("str") != _FakeDispatch → 提前 return
        assert not hasattr(sap, 'application')

    @patch('sap.function.win32com.client')
    def test_application_type_check_fail(self, mock_client):
        """application 类型检查失败 — SapGuiAuto 置空"""
        fake_gui = _FakeDispatch()
        fake_gui.GetScriptingEngine = "not_a_dispatch"  # 非 _FakeDispatch 类型
        mock_client.GetObject.return_value = fake_gui
        mock_client.CDispatch = _FakeDispatch

        sap = Sap(make_config(), make_flags())

        assert sap.SapGuiAuto is None

    @patch('sap.function.win32com.client')
    def test_connection_exception(self, mock_client):
        """连接异常 — result 标记失败"""
        mock_client.GetObject.side_effect = Exception("SAP未启动")

        sap = Sap(make_config(), make_flags())

        assert not sap.result.success
        assert 'SAP未启动' in sap.result.message


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
