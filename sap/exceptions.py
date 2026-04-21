"""SAP 模块异常定义。"""


class SapError(Exception):
    """SAP 模块基类异常。"""


class SapConnectionError(SapError):
    """SAP GUI 连接异常。"""


class SapUiError(SapError):
    """SAP GUI 控件操作异常。"""
