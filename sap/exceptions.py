"""SAP 模块异常定义。"""


class SapError(Exception):
    """SAP 模块基类异常。"""


class SapConnectionError(SapError):
    """SAP GUI 连接异常。"""


class SapUiError(SapError):
    """SAP GUI 控件操作异常。"""


class SapWriteError(SapError):
    """SAP GUI 控件**写入**被拒（只读字段等），消息带控件 ID 与目标值。

    刻意**不继承 SapUiError**：全仓库大量 `except SapUiError` 用于"控件不存在→跳过"
    的判定，写入被拒是完全不同的语义（控件在、就是不让改），混进去会让只读字段被
    静默当成"控件不存在"跳过。继承 SapError 而非 SapUiError 后，现有那些 `except
    Exception` 的容错点行为不变（原本抛的 COM 异常同样不是 SapUiError），只是消息
    从裸的 `Property '.text' can not be set.` 变成能直接定位到控件的完整信息。
    """
