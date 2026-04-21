"""会话层测试。"""

from unittest.mock import MagicMock, patch

import pytest

from sap.exceptions import SapConnectionError, SapUiError
from sap.session import SapSession
from sap.test.helpers import create_raw_session


class _FakeDispatch:
    pass


def _build_fake_chain():
    sap_gui_auto = _FakeDispatch()
    application = _FakeDispatch()
    connection = _FakeDispatch()
    session = _FakeDispatch()

    sap_gui_auto.GetScriptingEngine = application
    application.Children = lambda _: connection
    connection.Children = lambda _: session
    return sap_gui_auto, session


class TestSapSessionConnect:
    @patch("sap.session.win32com.client")
    def test_connect_success(self, mock_client):
        fake_gui, fake_session = _build_fake_chain()
        mock_client.GetObject.return_value = fake_gui
        mock_client.CDispatch = _FakeDispatch

        session = SapSession.connect()

        assert session.raw is fake_session

    @patch("sap.session.win32com.client")
    def test_connect_invalid_dispatch_raises(self, mock_client):
        mock_client.GetObject.return_value = "bad"
        mock_client.CDispatch = _FakeDispatch

        with pytest.raises(SapConnectionError):
            SapSession.connect()


class TestSapSessionOperations:
    def test_find_missing_element_raises_ui_error(self):
        raw = MagicMock()
        raw.findById.side_effect = Exception("not found")
        session = SapSession(MagicMock(), MagicMock(), MagicMock(), raw)

        with pytest.raises(SapUiError):
            session.find("wnd[0]/bad")

    def test_try_send_vkey_returns_false_on_missing_window(self):
        raw = create_raw_session()
        session = SapSession(MagicMock(), MagicMock(), MagicMock(), raw)
        raw.findById.side_effect = Exception("missing")

        assert not session.try_send_vkey(0, window_id="wnd[1]")

    def test_close_clears_internal_references(self):
        session = SapSession(MagicMock(), MagicMock(), MagicMock(), MagicMock())
        session.close()
        assert session.raw is None
