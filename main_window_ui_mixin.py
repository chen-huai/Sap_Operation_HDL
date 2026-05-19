import sys
import os
import re
import time
import math
import pandas as pd
import csv
import copy
import numpy as np
import win32com.client
import datetime
import shutil
import logging

from PyQt5.QtWidgets import QApplication, QFileDialog, QMainWindow, QMessageBox, QVBoxLayout, QPushButton, QAction, QLabel
from PyQt5.QtCore import QDate
from PyQt5.QtGui import QIcon, QFontDatabase
from Get_Data import *
from PDF_Parser_Utils import extract_company_name, extract_revenue, extract_fapiao_no, parse_pdf_fields, PDF_Operate
from Data_Table import *
from Logger import *
from Excel_Field_Mapper import excel_field_mapper
from theme_manager_theme import ThemeManager
from Revenue_Operate import *
from auto_updater.config_constants import CURRENT_VERSION
from auto_updater import AutoUpdater, UI_AVAILABLE
from sap import CostOptions, OrderData, OrderService, PartnerOptions, RevenueData, SapConfig, SapSession

class MainWindowUiMixin:
    @staticmethod
    def format_file_size(size_bytes: int) -> str:
        """
        格式化文件大小显示

        Args:
            size_bytes: 文件大小（字节）

        Returns:
            格式化后的文件大小字符串
        """
        if size_bytes <= 0:
            return "0 B"

        size_names = ["B", "KB", "MB", "GB", "TB"]
        i = int(math.floor(math.log(size_bytes, 1024)))
        p = math.pow(1024, i)
        s = round(size_bytes / p, 2)
        return f"{s} {size_names[i]}"
    def init_theme_action(self):
        theme_action = QAction(QIcon('theme_icon.png'), 'Toggle Theme', self)
        theme_action.setStatusTip('Toggle Theme')
        theme_action.triggered.connect(self.toggle_theme)

        # 将 action 添加到菜单（如果有的话）
        if hasattr(self, 'menuBar'):
            view_menu = self.menuBar().addMenu('Theme')
            view_menu.addAction(theme_action)
    def toggle_theme(self):
        self.theme_manager.set_random_theme()
    def setup_auto_update(self):
        """设置自动更新功能 - 标准集成方案"""
        logger = logging.getLogger(__name__)
        try:
            logger.info("开始初始化自动更新功能...")

            if UI_AVAILABLE:
                # 初始化更新器
                self.auto_updater = AutoUpdater(self)
                logger.info("AutoUpdater实例创建成功")

                # 设置更新UI - 集成到帮助菜单
                success = self.auto_updater.setup_update_ui(
                    self.menuBar(), "帮助(&H)"
                )

                if success:
                    logger.info("自动更新功能集成成功")

                    # 可选配置
                    # self.auto_updater.enable_auto_check = True
                    # self.auto_updater.check_interval = 24 * 3600

                else:
                    logger.warning("自动更新功能集成失败")
                    self.auto_updater = None
            else:
                logger.warning("UI模块不可用，跳过自动更新功能")
                self.auto_updater = None

        except ImportError as e:
            logger.error(f"导入自动更新模块失败: {e}")
            self.auto_updater = None

        except Exception as e:
            logger.error(f"自动更新器初始化失败: {e}")
            self.auto_updater = None
    def showAuthorMessage(self):
        # 关于作者
        QMessageBox.about(self, "关于",
                          "人生苦短，码上行乐。\n\n\n        ----Frank Chen")
    def showVersion(self):
        # 显示版本信息 - 动态从配置常量获取
        QMessageBox.about(self, "版本",
                          f"版本: {CURRENT_VERSION}")
    def update_status_bar(self):
        """更新状态栏显示版本信息（右下角）"""
        try:
            # 清除现有的永久小部件（如果存在）
            if hasattr(self, '_version_widget'):
                self.status_bar.removeWidget(self._version_widget)

            # 创建版本标签
            version_text = f"版本: {CURRENT_VERSION}"
            self._version_widget = QLabel(version_text)
            self._version_widget.setStyleSheet("color: #666; font-size: 12px; padding: 0 5px;")

            # 添加到状态栏右侧
            self.status_bar.addPermanentWidget(self._version_widget)

            # 清除主消息区域
            self.status_bar.showMessage("")
        except Exception as e:
            print(f"更新状态栏失败: {e}")
            self.status_bar.showMessage("状态栏初始化失败")
    def check_for_updates_startup(self):
        """兼容保留；启动检查已由 auto_updater 内部统一处理。"""
        try:
            if hasattr(self, 'auto_updater') and self.auto_updater:
                self.auto_updater.ui_manager.startup_update_check()
            else:
                print("自动更新器未初始化")

        except Exception as e:
            print(f"启动更新检查异常: {e}")
    def update_status_bar_with_update(self, remote_version=None):
        """更新状态栏显示更新提示"""
        try:
            if remote_version:
                version_text = f"发现新版本: {remote_version} (当前: {CURRENT_VERSION})"
            else:
                version_text = f"版本: {CURRENT_VERSION}"
            self.status_bar.showMessage(version_text)
        except Exception as e:
            print(f"更新状态栏失败: {e}")
            self.status_bar.showMessage("状态栏初始化失败")
    def show_update_notification(self):
        """显示更新通知"""
        try:
            # 检查最新的版本信息
            has_update, remote_version, local_version, error_msg = self.auto_updater.check_for_updates(force_check=True)

            if not has_update:
                # 显示已是最新版本提示
                QMessageBox.information(
                    self,
                    "检查更新",
                    f"当前版本 {local_version} 已是最新版本！",
                    QMessageBox.Yes
                )
                # 恢复状态栏显示当前版本
                self.update_status_bar()
                return

            # 创建更新对话框
            reply = QMessageBox.question(
                self,
                "发现新版本",
                f"发现新版本 {remote_version}，当前版本: {local_version}\n\n是否现在下载更新？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes
            )

            if reply == QMessageBox.Yes:
                # 显示下载进度
                self.start_update_process()

        except Exception as e:
            print(f"显示更新通知失败: {e}")
            QMessageBox.information(self, "提示", f"检查更新失败: {str(e)}", QMessageBox.Yes)
    def start_update_process(self):
        """启动更新过程"""
        try:
            # 获取更新信息
            has_update, remote_version, local_version, error_msg = self.auto_updater.check_for_updates(force_check=True)

            if not has_update:
                QMessageBox.information(self, "提示", "当前已是最新版本", QMessageBox.Yes)
                return

            # 下载进度对话框
            progress_dialog = QMessageBox()
            progress_dialog.setWindowTitle("下载更新")
            progress_dialog.setText(f"准备下载版本 {remote_version}...")
            progress_dialog.setStandardButtons(QMessageBox.NoButton)
            progress_dialog.setModal(True)  # 设置为模态对话框
            progress_dialog.setIcon(QMessageBox.Information)  # 设置信息图标

            # 居中显示对话框
            progress_dialog.setGeometry(
                self.x() + (self.width() - self.UIConstants.PROGRESS_DIALOG_WIDTH) // 2,
                self.y() + (self.height() - self.UIConstants.PROGRESS_DIALOG_HEIGHT) // 2,
                self.UIConstants.PROGRESS_DIALOG_WIDTH,
                self.UIConstants.PROGRESS_DIALOG_HEIGHT
            )

            progress_dialog.show()

            # 启动下载
            success, download_path, error = self.auto_updater.download_update(
                remote_version,
                lambda downloaded, total, percentage: self.update_download_progress(downloaded, total, percentage, progress_dialog)
            )

            progress_dialog.hide()

            if success:
                # 询问是否执行更新
                reply = QMessageBox.question(
                    self,
                    "下载完成",
                    f"更新文件已下载到:\n{download_path}\n\n是否现在安装更新？",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.Yes
                )

                if reply == QMessageBox.Yes:
                    success, error = self.auto_updater.execute_update(download_path, remote_version)
                    if success:
                        QMessageBox.information(self, "更新完成", "应用已更新，将在下次启动时生效", QMessageBox.Yes)
                        # 可选择重启应用
                        reply = QMessageBox.question(
                            self, "更新成功",
                            "更新完成！是否重启应用？",
                            QMessageBox.Yes | QMessageBox.No,
                            QMessageBox.Yes
                        )
                        if reply == QMessageBox.Yes:
                            import subprocess
                            subprocess.Popen([sys.executable, sys.argv[0]], creationflags=subprocess.CREATE_NEW_CONSOLE)
                    else:
                        QMessageBox.warning(self, "更新失败", f"更新失败: {error}", QMessageBox.Yes)
            else:
                QMessageBox.warning(self, "下载失败", f"下载失败: {error}", QMessageBox.Yes)

        except Exception as e:
            QMessageBox.warning(self, "更新异常", f"更新过程异常: {str(e)}", QMessageBox.Yes)
    def update_download_progress(self, downloaded: int, total: int, percentage: float, progress_dialog: QMessageBox):
        """
        更新下载进度显示（优化版本）

        Args:
            downloaded: 已下载字节数
            total: 总文件大小
            percentage: 下载百分比 (0-100)
            progress_dialog: 进度对话框对象
        """
        try:
            # 参数验证
            if progress_dialog is None:
                return

            if not all(isinstance(x, (int, float)) for x in [downloaded, total, percentage]):
                return

            # 确保百分比在有效范围内
            percentage = min(100, max(0, percentage))

            # 性能优化：控制更新频率
            import time
            current_time = time.time()
            if current_time - self._last_update_time < self.UIConstants.UPDATE_INTERVAL:
                return
            self._last_update_time = current_time

            # 使用优化的文件大小格式化方法
            downloaded_str = self.format_file_size(int(downloaded))
            total_str = self.format_file_size(int(total))

            # 更新进度对话框的文本
            if percentage >= 100:
                progress_text = f"下载完成！\n已下载: {downloaded_str}"
            else:
                progress_text = f"正在下载更新... {percentage:.1f}%\n已下载: {downloaded_str} / {total_str}"

            progress_dialog.setText(progress_text)

            # 优化的界面更新：安全获取 QApplication 实例
            try:
                from PyQt5.QtWidgets import QApplication
                app = QApplication.instance()
                if app:
                    QApplication.processEvents()
            except Exception:
                # 如果无法获取 QApplication 实例，静默处理
                pass

        except Exception as e:
            # 记录错误日志，不影响下载过程
            import logging
            logging.warning(f"更新进度显示失败: {str(e)}")
    def closeEvent(self, event):
        """应用退出时清理资源"""
        logger = logging.getLogger(__name__)
        try:
            # 清理自动更新器资源
            if hasattr(self, 'auto_updater') and self.auto_updater:
                self.auto_updater.cleanup()
                logger.info("自动更新器资源已清理")

            logger.info("应用程序退出")
            event.accept()

        except Exception as e:
            logger.error(f"清理资源时出错: {e}")
            event.accept()  # 确保应用能够正常退出

