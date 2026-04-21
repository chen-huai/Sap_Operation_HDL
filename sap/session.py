"""SAP GUI 会话封装。"""

from __future__ import annotations

import time

import win32com.client

from sap.exceptions import SapConnectionError, SapUiError


class SapSession:
    """对 SAP GUI COM Session 的薄封装。"""

    def __init__(self, sap_gui_auto, application, connection, session):
        """保存整个 COM 链路，便于后续关闭或扩展。"""
        self._sap_gui_auto = sap_gui_auto
        self._application = application
        self._connection = connection
        self._session = session

    @property
    def raw(self):
        """暴露原始 session，供少数必须直接访问 COM API 的场景使用。"""
        return self._session

    @classmethod
    def connect(cls, connection_index: int = 0, session_index: int = 0) -> "SapSession":
        """连接到指定的 SAP connection/session。"""
        try:
            # SAP GUI -> application -> connection -> session 是固定的 COM 访问链路。
            sap_gui_auto = win32com.client.GetObject("SAPGUI")
            cls._ensure_dispatch(sap_gui_auto, "SAPGUI")

            application = sap_gui_auto.GetScriptingEngine
            cls._ensure_dispatch(application, "GetScriptingEngine")

            connection = application.Children(connection_index)
            cls._ensure_dispatch(connection, f"Children({connection_index})")

            session = connection.Children(session_index)
            cls._ensure_dispatch(session, f"Session({session_index})")
            return cls(sap_gui_auto, application, connection, session)
        except SapConnectionError:
            raise
        except Exception as exc:
            raise SapConnectionError(str(exc)) from exc

    @staticmethod
    def _ensure_dispatch(obj, name: str) -> None:
        """校验 COM 对象类型，避免把半初始化对象继续向下传。"""
        if type(obj) != win32com.client.CDispatch:
            raise SapConnectionError(f"{name} 不是有效的 SAP COM 对象")

    def find(self, element_id: str):
        """按 SAP 元素 ID 查找控件。"""
        try:
            return self._session.findById(element_id)
        except Exception as exc:
            raise SapUiError(f"找不到 SAP 元素: {element_id}") from exc

    def read_text(self, element_id: str) -> str:
        """读取控件文本。"""
        return self.find(element_id).text

    def set_text(self, element_id: str, value) -> None:
        """写入控件文本。"""
        self.find(element_id).text = value

    def set_key(self, element_id: str, value: str) -> None:
        """写入下拉框 key 值。"""
        self.find(element_id).key = value

    def set_selected(self, element_id: str, value: bool) -> None:
        """写入复选框选中状态。"""
        self.find(element_id).selected = value

    def press(self, element_id: str) -> None:
        """点击按钮。"""
        self.find(element_id).press()

    def send_vkey(self, key: int, *, window_id: str = "wnd[0]") -> None:
        """向指定窗口发送 SAP 虚拟按键。"""
        self.find(window_id).sendVKey(key)

    def try_send_vkey(self, key: int, *, window_id: str = "wnd[0]") -> bool:
        # 某些弹窗是条件出现的，这里提供“尽力而为”的键盘操作。
        try:
            self.send_vkey(key, window_id=window_id)
            return True
        except SapUiError:
            return False

    def focus(self, element_id: str, caret_position: int | None = None) -> None:
        """给控件设置焦点，并在需要时移动光标。"""
        element = self.find(element_id)
        element.setFocus()
        if caret_position is not None:
            element.caretPosition = caret_position

    def set_selection_indexes(self, element_id: str, start: int, end: int) -> None:
        """设置文本控件的选区范围。"""
        self.find(element_id).setSelectionIndexes(start, end)

    def select_tab(self, element_id: str) -> None:
        """切换到指定页签。"""
        self.find(element_id).select()

    def read_status(self) -> str:
        """读取状态栏文本。"""
        return self.read_text("wnd[0]/sbar/pane[0]")

    def sleep(self, seconds: float) -> None:
        """保留显式等待入口，避免业务层直接依赖 time.sleep。"""
        time.sleep(seconds)

    def close(self) -> None:
        """断开本地引用，释放 COM 链路。"""
        self._session = None
        self._connection = None
        self._application = None
        self._sap_gui_auto = None
