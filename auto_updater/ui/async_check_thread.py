# -*- coding: utf-8 -*-
"""
异步更新检查线程

提供基于 QThread 的版本检查，避免阻塞 GUI 主线程。
与 AsyncDownloadThread 同构，结果通过信号回主线程。
"""

import logging
from typing import Optional

from PyQt5.QtCore import QThread, pyqtSignal

logger = logging.getLogger(__name__)


class UpdateCheckThread(QThread):
    """异步执行版本检查的 QThread。

    Signals:
        check_finished(bool, object, str, object):
            (has_update, remote_version, local_version, error)
            remote_version / error 可能为 None，故用 object 类型承载。
    """

    check_finished = pyqtSignal(bool, object, str, object)

    def __init__(
        self,
        auto_updater,
        is_silent: bool = True,
        force_check: bool = False,
        parent: Optional[object] = None,
    ):
        super().__init__(parent)
        self.auto_updater = auto_updater
        self.is_silent = is_silent
        self.force_check = force_check

    def run(self) -> None:
        """子线程执行版本检查，所有异常吞掉转为信号。"""
        try:
            has_update, remote_version, local_version, error = (
                self.auto_updater.check_for_updates(
                    force_check=self.force_check,
                    is_silent=self.is_silent,
                )
            )
            self.check_finished.emit(
                bool(has_update),
                remote_version,
                local_version or "",
                error,
            )
        except Exception as e:
            logger.debug(f"异步检查更新异常（静默）: {e}")
            self.check_finished.emit(False, None, "", str(e))
