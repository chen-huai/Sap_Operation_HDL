# -*- coding: utf-8 -*-
"""
自动更新UI管理器
提供独立的更新功能UI接口，实现UI与业务逻辑的完全解耦

使用方法：
    # 在主程序中
    updater = AutoUpdater(main_window)
    updater.setup_update_ui(menu_bar)

    # 或者直接使用
    ui_manager = UpdateUIManager(updater, main_window)
    ui_manager.setup_update_menu(menu_bar)
"""
import logging
import os
from typing import Optional, Any
from PyQt5.QtWidgets import (
    QMainWindow,
    QMessageBox,
    QMenuBar,
    QMenu,
    QAction,
    QWidget,
    QProgressDialog,
)
from PyQt5.QtCore import QObject, Qt, pyqtSignal

from .dialogs import UpdateProgressDialog, AboutDialog
from .resources import UpdateUIText
from .async_check_thread import UpdateCheckThread

# 获取模块专用日志记录器
logger = logging.getLogger(__name__)

class UpdateUIManager(QObject):
    """
    自动更新UI管理器

    负责管理所有与更新功能相关的UI组件，提供标准化的接口供主程序调用。
    实现UI与业务逻辑的完全解耦，便于维护和移植。
    """

    # 信号定义
    update_started = pyqtSignal(str)  # 更新开始，参数：版本号
    update_finished = pyqtSignal(bool, str)  # 更新完成，参数：是否成功、错误信息
    update_progress = pyqtSignal(int, int, int)  # 更新进度，参数：已下载、总量、百分比

    def __init__(self, auto_updater: Any, parent: Optional[QWidget] = None):
        """
        初始化UI管理器

        Args:
            auto_updater: AutoUpdater实例
            parent: 父窗口组件

        Raises:
            TypeError: 当auto_updater不是AutoUpdater实例时
            ValueError: 当parent不是QWidget实例时
        """
        super().__init__(parent)

        # 参数验证
        if auto_updater is None:
            raise TypeError("auto_updater参数不能为None")

        if parent is not None and not hasattr(parent, 'menuBar'):
            logger.warning("parent对象缺少menuBar方法，某些功能可能不可用")

        self.auto_updater = auto_updater
        self.parent = parent

        # UI组件缓存
        self._update_action = None
        self._version_action = None
        self._about_dialog = None
        self._update_dialog = None
        self._help_menu = None

        # 状态标志
        self._is_initialized = False
        self._is_menu_setup = False

        # 启动检查异步线程（持引用避免被 GC）
        self._check_thread: Optional[UpdateCheckThread] = None

        # 手点更新异步状态（与启动检查独立，避免相互干扰）
        self._manual_check_thread: Optional[UpdateCheckThread] = None
        self._manual_check_dialog: Optional[QProgressDialog] = None
        self._manual_check_cancelled: bool = False

        logger.info("UpdateUIManager 初始化完成")

    def setup_update_menu(self, menu_bar: QMenuBar, menu_title: str = None) -> None:
        """
        设置更新菜单

        Args:
            menu_bar: 主窗口的菜单栏
            menu_title: 菜单标题，默认使用UpdateUIText中的标题

        Raises:
            TypeError: 当menu_bar不是QMenuBar实例时
            RuntimeError: 当菜单设置失败时
        """
        if menu_bar is None:
            raise TypeError("menu_bar参数不能为None")

        if self._is_menu_setup:
            logger.warning("更新菜单已经设置过，跳过重复设置")
            return

        try:
            # 使用默认菜单标题
            if menu_title is None:
                menu_title = UpdateUIText.HELP_MENU_TITLE

            logger.debug(f"开始设置更新菜单，标题: {menu_title}")

            # 查找现有的actionUpdate对象（在主窗口UI中）
            existing_update_action = self._find_existing_action(self.parent, "actionUpdate")
            existing_version_action = self._find_existing_action(self.parent, "actionbangbenv1_0_0")

            if existing_update_action:
                # 使用现有的actionUpdate对象，连接我们的功能
                self._update_action = existing_update_action
                # 使用lambda强制传递force_check=True，确保手动更新时跳过频率检查
                self._update_action.triggered.connect(lambda: self.check_for_updates_with_ui(force_check=True))
                self._update_action.setStatusTip(UpdateUIText.UPDATE_MENU_TOOLTIP)
                logger.debug("使用现有actionUpdate对象并连接信号")
            else:
                # 如果找不到现有对象，创建新的（这种情况应该不会发生）
                logger.warning("未找到现有actionUpdate对象，创建新的")
                help_menu = self._find_or_create_menu(menu_bar, menu_title)
                if help_menu is None:
                    raise RuntimeError("无法创建或查找帮助菜单")

                self._update_action = QAction(UpdateUIText.UPDATE_MENU_TEXT, self.parent)
                self._update_action.setObjectName("actionUpdate")
                # 使用lambda强制传递force_check=True，确保手动更新时跳过频率检查
                self._update_action.triggered.connect(lambda: self.check_for_updates_with_ui(force_check=True))
                self._update_action.setStatusTip(UpdateUIText.UPDATE_MENU_TOOLTIP)
                help_menu.addAction(self._update_action)

            if existing_version_action:
                # 使用现有的版本动作
                self._version_action = existing_version_action
                self._version_action.triggered.connect(self.show_about_dialog)
                self._version_action.setStatusTip(UpdateUIText.VERSION_MENU_TOOLTIP)
                logger.debug("使用现有版本动作对象并连接信号")

            # 缓存菜单引用
            self._help_menu = menu_bar.findChild(QMenu, menu_title) if menu_title else None
            self._is_menu_setup = True

            logger.info(f"更新菜单设置完成，信号连接成功")

        except Exception as e:
            logger.error(f"设置更新菜单失败: {e}")
            # 清理可能创建的部分组件
            self._cleanup_menu_items()
            raise RuntimeError(f"设置更新菜单失败: {str(e)}")

    def _find_existing_action(self, parent, object_name: str):
        """
        在父对象中查找现有的动作对象

        Args:
            parent: 父对象（主窗口）
            object_name: 动作对象名称

        Returns:
            QAction or None: 找到的动作对象，如果未找到则返回None
        """
        try:
            if parent is None:
                return None

            # 在父对象中查找指定名称的动作
            for child in parent.children():
                if hasattr(child, 'objectName') and child.objectName() == object_name:
                    logger.debug(f"找到现有动作对象: {object_name}")
                    return child

            # 如果直接查找失败，尝试在菜单中查找
            if hasattr(parent, 'menuBar'):
                menu_bar = parent.menuBar()
                for menu in menu_bar.findChildren(QMenu):
                    for action in menu.actions():
                        if hasattr(action, 'objectName') and action.objectName() == object_name:
                            logger.debug(f"在菜单中找到动作对象: {object_name}")
                            return action

            logger.debug(f"未找到现有动作对象: {object_name}")
            return None

        except Exception as e:
            logger.error(f"查找现有动作对象失败: {e}")
            return None

    def check_for_updates_with_ui(self, force_check: bool = True) -> None:
        """
        检查更新（带UI交互）— 异步执行，避免 GUI 主线程阻塞

        Args:
            force_check: 是否强制检查更新

        Note:
            改造说明（2026-05-20）：
            - 原实现在主线程同步调用 self.auto_updater.check_for_updates() →
              requests.get() 网络异常时 GUI 卡死最长 ≈ 166 秒（HTTPAdapter Retry +
              @network_retry 双层重试 + 频率限制 sleep(60)）
            - 现改为 UpdateCheckThread 异步执行 + QProgressDialog 忙等 + 可取消
            - 结果通过 check_finished 信号回主线程 _on_manual_check_finished 处理
        """
        try:
            logger.info("actionUpdate按钮被点击，开始执行更新检查")

            if not self.auto_updater:
                logger.error("auto_updater对象为None")
                QMessageBox.warning(
                    self.parent,
                    UpdateUIText.UPDATE_UNAVAILABLE_TITLE,
                    UpdateUIText.UPDATE_UNAVAILABLE_MESSAGE
                )
                return

            # 防重入：手点检查线程仍在跑则直接返回，避免重复创建对话框
            if (
                self._manual_check_thread is not None
                and self._manual_check_thread.isRunning()
            ):
                logger.debug("手点更新检查仍在进行中，忽略重复点击")
                return

            # 重置取消标志
            self._manual_check_cancelled = False

            # 显示检查进度（如果父窗口有文本浏览器）
            self._show_status_message(UpdateUIText.CHECKING_UPDATE_MESSAGE)

            # 创建忙等对话框（range=0,0 即 indeterminate 旋转动画）
            dialog = QProgressDialog(
                UpdateUIText.CHECKING_UPDATE_MESSAGE,
                "取消",
                0,
                0,
                self.parent,
            )
            dialog.setWindowTitle(UpdateUIText.CHECK_UPDATE_TITLE)
            dialog.setWindowModality(Qt.ApplicationModal)
            # 去掉标题栏的 ? 帮助按钮，保持简洁
            dialog.setWindowFlags(
                dialog.windowFlags() & ~Qt.WindowContextHelpButtonHint
            )
            dialog.setMinimumDuration(0)  # 立即显示，不等 4 秒默认延迟
            dialog.setAutoClose(False)
            dialog.setAutoReset(False)
            dialog.canceled.connect(self._on_manual_check_cancelled)

            # 创建并启动异步检查线程
            thread = UpdateCheckThread(
                auto_updater=self.auto_updater,
                is_silent=False,
                force_check=force_check,
                parent=self,
            )
            thread.check_finished.connect(self._on_manual_check_finished)
            thread.finished.connect(thread.deleteLater)
            thread.finished.connect(self._clear_manual_check_refs)

            self._manual_check_thread = thread
            self._manual_check_dialog = dialog

            dialog.show()
            thread.start()

        except Exception as e:
            logger.error(f"检查更新异常: {e}")
            # 异常路径兜底清理
            self._safe_close_manual_dialog()
            QMessageBox.critical(
                self.parent,
                UpdateUIText.ERROR_TITLE,
                f"{UpdateUIText.CHECK_UPDATE_ERROR_MESSAGE}\n{str(e)}"
            )

    def _on_manual_check_finished(
        self,
        has_update: bool,
        remote_version,
        local_version: str,
        error,
    ) -> None:
        """
        手点更新的异步检查结果处理槽（主线程执行）

        Args:
            has_update: 是否发现新版本
            remote_version: 远程版本号（可能为 None）
            local_version: 本地版本号
            error: 错误信息（可能为 None）
        """
        # 不论何种结果，先关闭忙等对话框
        self._safe_close_manual_dialog()

        # 用户已点取消：静默丢弃结果，不再弹任何对话框
        if self._manual_check_cancelled:
            logger.debug("手点检查已被取消，丢弃异步结果")
            return

        try:
            if error:
                QMessageBox.warning(
                    self.parent,
                    UpdateUIText.CHECK_UPDATE_FAILED_TITLE,
                    f"{UpdateUIText.CHECK_UPDATE_FAILED_MESSAGE}\n{error}",
                )
                self._show_status_message(
                    f"{UpdateUIText.CHECK_UPDATE_FAILED_MESSAGE}: {error}"
                )
                return

            if has_update and remote_version:
                # 发现新版本
                reply = QMessageBox.question(
                    self.parent,
                    UpdateUIText.NEW_VERSION_FOUND_TITLE,
                    UpdateUIText.NEW_VERSION_FOUND_MESSAGE.format(
                        remote_version=remote_version,
                        local_version=local_version,
                    ),
                    QMessageBox.Yes | QMessageBox.No,
                )
                if reply == QMessageBox.Yes:
                    self.start_update_process(remote_version)
                else:
                    self._show_status_message(UpdateUIText.UPDATE_CANCELLED_MESSAGE)
            else:
                # 已是最新版本
                QMessageBox.information(
                    self.parent,
                    UpdateUIText.CHECK_UPDATE_TITLE,
                    UpdateUIText.LATEST_VERSION_MESSAGE,
                )
                self._show_status_message(UpdateUIText.LATEST_VERSION_MESSAGE)

        except Exception as e:
            logger.error(f"处理手点检查结果异常: {e}")
            QMessageBox.critical(
                self.parent,
                UpdateUIText.ERROR_TITLE,
                f"{UpdateUIText.CHECK_UPDATE_ERROR_MESSAGE}\n{str(e)}",
            )

    def _on_manual_check_cancelled(self) -> None:
        """
        用户点击「取消」按钮：标记取消、关闭对话框、状态提示

        Note:
            QThread 在 requests.get() 中无法真正中断，
            线程会继续在后台跑完，结果到达后被 _on_manual_check_finished 静默丢弃
        """
        try:
            logger.info("用户取消了手点更新检查")
            self._manual_check_cancelled = True

            # 声明式中断请求（即使无效也无害）
            if (
                self._manual_check_thread is not None
                and self._manual_check_thread.isRunning()
            ):
                self._manual_check_thread.requestInterruption()

            self._safe_close_manual_dialog()
            self._show_status_message("已取消更新检查")

        except Exception as e:
            logger.debug(f"取消手点检查处理异常（静默）: {e}")

    def _clear_manual_check_refs(self) -> None:
        """线程 finished 信号触发后清空引用，便于下次点击重新启动。"""
        self._manual_check_thread = None
        self._manual_check_dialog = None

    def _safe_close_manual_dialog(self) -> None:
        """安全关闭手点检查的忙等对话框（幂等）。"""
        try:
            dialog = self._manual_check_dialog
            if dialog is not None:
                # 断开 canceled 信号，避免 close() 触发递归
                try:
                    dialog.canceled.disconnect(self._on_manual_check_cancelled)
                except (TypeError, RuntimeError):
                    pass
                dialog.close()
                dialog.deleteLater()
                self._manual_check_dialog = None
        except Exception as e:
            logger.debug(f"关闭手点检查对话框异常（静默）: {e}")

    def show_about_dialog(self) -> None:
        """显示关于对话框"""
        try:
            if not self._about_dialog:
                self._about_dialog = AboutDialog(self.auto_updater, self.parent)

            self._about_dialog.show()
            self._about_dialog.raise_()
            self._about_dialog.activateWindow()

        except Exception as e:
            logger.error(f"显示关于对话框失败: {e}")
            QMessageBox.critical(
                self.parent,
                UpdateUIText.ERROR_TITLE,
                f"{UpdateUIText.SHOW_ABOUT_ERROR_MESSAGE}\n{str(e)}"
            )

    def start_update_process(self, version: str) -> None:
        """
        启动更新过程

        Args:
            version: 要更新到的版本号
        """
        try:
            if not self.auto_updater:
                return

            # 创建更新进度对话框
            if not self._update_dialog:
                self._update_dialog = UpdateProgressDialog(self.parent)

            # 开始更新
            self._update_dialog.start_update(version, self.auto_updater)
            self._update_dialog.exec_()

            # 发送更新开始信号
            self.update_started.emit(version)

        except Exception as e:
            logger.error(f"启动更新过程失败: {e}")
            QMessageBox.critical(
                self.parent,
                UpdateUIText.UPDATE_FAILED_TITLE,
                f"{UpdateUIText.START_UPDATE_ERROR_MESSAGE}\n{str(e)}"
            )

    def startup_update_check(self) -> None:
        """
        启动时的静默更新检查
        只在生产环境执行，不显示UI提示
        """
        try:
            if not self.auto_updater:
                return

            import sys
            if not getattr(sys, 'frozen', False):
                return  # 开发环境不执行启动检查

            # 延迟执行，避免影响启动速度
            from PyQt5.QtCore import QTimer
            QTimer.singleShot(3000, self._perform_startup_check)

        except Exception as e:
            logger.error(f"启动更新检查设置失败: {e}")

    def _perform_startup_check(self) -> None:
        """
        启动后台线程执行版本检查（不阻塞 GUI 主线程）

        改造说明：
        - 原实现在主线程同步调用 requests.get()，网络异常时 GUI 卡死最长 ≈ 166 秒
        - 现改为 UpdateCheckThread 异步执行；结果通过信号回主线程槽 _on_startup_check_finished 处理
        """
        try:
            # 防止重复启动：若上次线程未结束，跳过本次
            if self._check_thread is not None and self._check_thread.isRunning():
                logger.debug("启动检查线程仍在运行，跳过本次")
                return

            thread = UpdateCheckThread(
                auto_updater=self.auto_updater,
                is_silent=True,
                force_check=False,
                parent=self,
            )
            thread.check_finished.connect(self._on_startup_check_finished)
            # 线程结束后自动清理，避免内存泄漏
            thread.finished.connect(thread.deleteLater)
            thread.finished.connect(self._clear_check_thread_ref)

            self._check_thread = thread
            thread.start()

        except Exception as e:
            # 启动检查异常也静默处理，避免打扰用户
            logger.debug(f"启动更新检查异常（静默）: {e}")

    def _on_startup_check_finished(
        self,
        has_update: bool,
        remote_version,
        local_version: str,
        error,
    ) -> None:
        """
        启动检查结果处理槽（主线程执行）

        Args:
            has_update: 是否发现新版本
            remote_version: 远程版本号（可能为 None）
            local_version: 本地版本号
            error: 错误信息（可能为 None）
        """
        try:
            # 只有在有更新且无错误时才提示用户
            if has_update and error is None and remote_version:
                reply = QMessageBox.question(
                    self.parent,
                    UpdateUIText.NEW_VERSION_FOUND_TITLE,
                    UpdateUIText.NEW_VERSION_FOUND_MESSAGE.format(
                        remote_version=remote_version,
                        local_version=local_version,
                    ),
                    QMessageBox.Yes | QMessageBox.No,
                )

                if reply == QMessageBox.Yes:
                    self.start_update_process(remote_version)

            # 静默忽略其他情况：无更新 / 间隔未到 / 网络错误
            else:
                logger.debug(
                    f"启动检查无操作 (has_update={has_update}, error={error})"
                )

        except Exception as e:
            logger.debug(f"启动检查结果处理异常（静默）: {e}")

    def _clear_check_thread_ref(self) -> None:
        """线程结束后清空引用，便于下次创建。"""
        self._check_thread = None

    def _find_or_create_menu(self, menu_bar: QMenuBar, title: str) -> QMenu:
        """
        查找或创建菜单

        Args:
            menu_bar: 菜单栏
            title: 菜单标题

        Returns:
            QMenu: 菜单对象
        """
        # 查找现有菜单
        for action in menu_bar.actions():
            if action.text() == title:
                return action.menu()

        # 创建新菜单
        menu = QMenu(title, self.parent)
        menu_bar.addMenu(menu)
        return menu

    def _get_version_display_text(self) -> str:
        """获取版本显示文本"""
        try:
            if self.auto_updater:
                return f"{UpdateUIText.VERSION_PREFIX}{self.auto_updater.config.current_version}"
            else:
                from .. import get_version
                return f"{UpdateUIText.VERSION_PREFIX}{get_version()}"
        except Exception:
            return f"{UpdateUIText.VERSION_PREFIX}未知"

    def _show_status_message(self, message: str) -> None:
        """
        显示状态消息

        Args:
            message: 要显示的消息
        """
        try:
            if hasattr(self.parent, 'textBrowser'):
                self.parent.textBrowser.append(message)
            elif hasattr(self.parent, 'statusBar'):
                self.parent.statusBar().showMessage(message, 5000)  # 显示5秒
        except Exception as e:
            logger.debug(f"显示状态消息失败: {e}")

    def _cleanup_menu_items(self) -> None:
        """清理已创建的菜单项"""
        try:
            if self._help_menu:
                if self._update_action and self._update_action in self._help_menu.actions():
                    self._help_menu.removeAction(self._update_action)
                if self._version_action and self._version_action in self._help_menu.actions():
                    self._help_menu.removeAction(self._version_action)

            self._update_action = None
            self._version_action = None
            self._help_menu = None
            self._is_menu_setup = False

            logger.debug("菜单项清理完成")

        except Exception as e:
            logger.error(f"清理菜单项失败: {e}")

    def cleanup(self) -> None:
        """清理资源"""
        try:
            # 清理对话框
            if self._about_dialog:
                self._about_dialog.close()
                self._about_dialog = None

            if self._update_dialog:
                self._update_dialog.close()
                self._update_dialog = None

            # 关闭手点检查的忙等对话框
            self._safe_close_manual_dialog()

            # 标记手点检查为已取消，丢弃后续线程回调
            self._manual_check_cancelled = True

            # 手点检查线程：声明式中断，不强制 wait，避免退出阻塞
            if (
                self._manual_check_thread is not None
                and self._manual_check_thread.isRunning()
            ):
                try:
                    self._manual_check_thread.requestInterruption()
                except Exception as e:
                    logger.debug(f"中断手点检查线程异常（静默）: {e}")

            # 清理菜单项
            self._cleanup_menu_items()

            # 重置状态标志
            self._is_initialized = False
            self._is_menu_setup = False

            logger.info("UpdateUIManager 资源清理完成")

        except Exception as e:
            logger.error(f"清理UI管理器资源失败: {e}")

    def is_menu_setup(self) -> bool:
        """
        检查菜单是否已经设置

        Returns:
            bool: 菜单是否已设置
        """
        return self._is_menu_setup

    def get_menu_actions(self) -> list:
        """
        获取更新相关的菜单动作

        Returns:
            list: 包含更新动作和版本动作的列表
        """
        actions = []
        if self._update_action:
            actions.append(self._update_action)
        if self._version_action:
            actions.append(self._version_action)
        return actions

    def enable_update_menu(self, enabled: bool = True) -> None:
        """
        启用或禁用更新菜单

        Args:
            enabled: 是否启用更新功能
        """
        try:
            if self._update_action:
                self._update_action.setEnabled(enabled)
                logger.debug(f"更新菜单已{'启用' if enabled else '禁用'}")

        except Exception as e:
            logger.error(f"设置更新菜单状态失败: {e}")