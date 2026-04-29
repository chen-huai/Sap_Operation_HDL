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
import chicon  # 引用图标
# from PyQt5 import QtCore, QtGui, QtWidgets
# from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtWidgets import QApplication, QFileDialog, QMainWindow, QMessageBox, QVBoxLayout, QPushButton, QAction, QLabel
from PyQt5.QtCore import QDate
from PyQt5.QtGui import QIcon, QFontDatabase
from Get_Data import *
# File_Operate 已合并，方法已内联到 MyMainWindow
# PDF_Operate 已合并到 PDF_Parser_Utils
from PDF_Parser_Utils import extract_company_name, extract_revenue, extract_fapiao_no, parse_pdf_fields, PDF_Operate
from Sap_Function import *
from Sap_Operate_Ui import Ui_MainWindow
from Data_Table import *
from Logger import *
from Excel_Field_Mapper import excel_field_mapper
from theme_manager_theme import ThemeManager
from Revenue_Operate import *
from auto_updater.config_constants import CURRENT_VERSION
from auto_updater import AutoUpdater, UI_AVAILABLE
from sap import CostOptions, OrderData, OrderService, PartnerOptions, RevenueData, SapConfig, SapSession
from main_window_ui_mixin import MainWindowUiMixin
from config_mixin import ConfigMixin
from sap_order_mixin import SapOrderMixin
from odm_invoice_mixin import OdmInvoiceMixin
from hour_mixin import HourMixin
import logging

# 延迟导入qt_material以避免警告
# qt_material 将在需要时由 ThemeManager 导入
import shutil
import builtins
from PyQt5 import QtCore








class MyMainWindow(MainWindowUiMixin, ConfigMixin, SapOrderMixin, OdmInvoiceMixin, HourMixin, QMainWindow, Ui_MainWindow):

    class UIConstants:
        PROGRESS_DIALOG_WIDTH = 400
        PROGRESS_DIALOG_HEIGHT = 200
        UPDATE_INTERVAL = 0.5  # UI更新间隔（秒）
        NETWORK_TIMEOUT = 5


    def __init__(self, parent=None):
        super(MyMainWindow, self).__init__(parent)
        self.setupUi(self)

        self.theme_manager = ThemeManager(QApplication.instance())
        self.init_theme_action()

        self.setGeometry(100, 100, 300, 200)

        self.theme_manager = ThemeManager(QApplication.instance())

        layout = QVBoxLayout()

        toggle_button = QPushButton("Toggle Theme")
        toggle_button.clicked.connect(self.theme_manager.toggle_theme)
        layout.addWidget(toggle_button)
        


        self.actionExport.triggered.connect(self.exportConfig)
        self.actionImport.triggered.connect(self.importConfig)
        self.actionExit.triggered.connect(self.close)
        self.actionHelp.triggered.connect(self.showVersion)
        self.actionAuthor.triggered.connect(self.showAuthorMessage)
        self.theme_manager.set_theme("default")  # 设置默认主题

        # 设置自动更新功能
        self.setup_auto_update()
        self.pushButton_11.clicked.connect(self.sap_operate)
        # self.pushButton_11.clicked.connect(self.sapOperate)
        self.pushButton_12.clicked.connect(self.textBrowser.clear)
        self.pushButton_20.clicked.connect(self.textBrowser_2.clear)
        self.pushButton_16.clicked.connect(self.getFileUrl)
        self.pushButton_18.clicked.connect(self.getODMDataFileUrl)
        self.pushButton_23.clicked.connect(self.getCombineFileUrl)
        self.pushButton_24.clicked.connect(self.getLogFileUrl)
        self.pushButton_17.clicked.connect(self.odmDataToSap)
        self.pushButton_19.clicked.connect(self.odmCombineData)
        self.pushButton_25.clicked.connect(self.orderMergeProject)
        self.pushButton_36.clicked.connect(self.splitOdmData)
        self.pushButton_34.clicked.connect(self.textBrowser_3.clear)
        self.pushButton_35.clicked.connect(self.invoiceRenameOperate)
        self.pushButton_33.clicked.connect(self.getPdfFiles)
        self.pushButton_48.clicked.connect(self.addMsg)
        self.pushButton_49.clicked.connect(self.viewOdmData)
        self.pushButton_59.clicked.connect(self.viewBillingListData)
        self.pushButton_50.clicked.connect(self.getEleInvoiceFiles)
        self.pushButton_51.clicked.connect(self.electronicInvoice)
        self.pushButton_58.clicked.connect(self.getBillingListFile)
        self.lineEdit_15.textChanged.connect(self.lineEditChange)
        self.doubleSpinBox_2.valueChanged.connect(self.getAmountVat)
        self.checkBox_9.toggled.connect(lambda: self.pdfNameRule('Invoice No', 'invoice'))
        self.checkBox_20.toggled.connect(lambda: self.pdfNameRule('Invoice No', 'Electron'))
        self.checkBox_10.toggled.connect(lambda: self.pdfNameRule('Company Name', 'invoice'))
        self.checkBox_22.toggled.connect(lambda: self.pdfNameRule('Company Name', 'Electron'))
        self.checkBox_12.toggled.connect(lambda: self.pdfNameRule('Order No', 'invoice'))
        self.checkBox_21.toggled.connect(lambda: self.pdfNameRule('Order No', 'Electron'))
        self.checkBox_11.toggled.connect(lambda: self.pdfNameRule('Project No', 'invoice'))
        self.checkBox_23.toggled.connect(lambda: self.pdfNameRule('FaPiao No', 'Electron'))
        self.checkBox_24.toggled.connect(lambda: self.pdfNameRule('Revenue', 'Electron'))
        self.checkBox_25.toggled.connect(lambda: self.pdfNameRule('CS', 'invoice'))
        self.checkBox_25.toggled.connect(lambda: self.pdfNameRule('CS', 'Electron'))
        self.checkBox_26.toggled.connect(lambda: self.pdfNameRule('Client Contact Name', 'invoice'))
        self.pushButton_56.clicked.connect(lambda: self.orderUnlockOrLock('Unlock'))
        self.pushButton_57.clicked.connect(lambda: self.orderUnlockOrLock('Lock'))
        self.pushButton_63.clicked.connect(lambda: self.get_hour_file_url(self.lineEdit_30))
        self.pushButton_71.clicked.connect(lambda: self.get_hour_file_url(self.lineEdit_38))
        self.pushButton_76.clicked.connect(lambda: self.get_hour_file_url(self.lineEdit_37))
        self.pushButton_78.clicked.connect(lambda: self.get_hour_file_url(self.lineEdit_31))
        self.pushButton_72.clicked.connect(self.get_hour_combine_file)
        self.pushButton_77.clicked.connect(self.get_department_hour)
        self.pushButton_73.clicked.connect(self.get_average_person_hour)
        # self.pushButton_73.clicked.connect(self.get_person_hour)
        self.pushButton_79.clicked.connect(self.hourOperate)
        self.pushButton_80.clicked.connect(self.clear_hour_gui)
        self.pushButton_81.clicked.connect(lambda: self.open_file(self.lineEdit_30.text()))
        self.pushButton_82.clicked.connect(lambda: self.open_file(self.lineEdit_37.text()))
        self.pushButton_83.clicked.connect(lambda: self.open_file(self.lineEdit_38.text()))
        self.pushButton_84.clicked.connect(lambda: self.open_file(self.lineEdit_31.text()))
        self.filesUrl = []

        # 初始化状态栏
        self.status_bar = self.statusBar()
        self.update_status_bar()

        # 性能优化：初始化更新控制变量
        self._last_update_time = 0


if __name__ == "__main__":
    # ================================================================
    # 步骤1: 自动完成更新（如果是从下载目录启动的新版本）
    # ================================================================
    try:
        from auto_updater.auto_complete import auto_complete_update_if_needed

        def update_callback(success, message):
            """更新完成回调函数"""
            if success:
                print(f"✓ 后台更新完成: {message}")
                print(f"✓ 下次启动将使用主目录的程序")
                # 可选：显示更新成功通知
            else:
                print(f"⚠ 后台更新: {message}")

        # 如果是从下载目录启动的新版本，会自动在后台完成文件替换
        # 这个调用会在后台线程运行，不会阻塞程序启动
        auto_complete_update_if_needed(update_callback)

    except Exception as e:
        # 如果自动完成功能不可用，仅记录错误，不影响程序启动
        print(f"自动完成更新检查失败: {e}")

    # ================================================================
    # 步骤2: 正常启动应用程序
    # ================================================================
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)
    app = QApplication(sys.argv)
    myWin = MyMainWindow()
    myTable = MyTableWindow()
    builtins.myWin = myWin
    builtins.myTable = myTable
    myWin.show()
    myWin.getConfig()
    sys.exit(app.exec_())
