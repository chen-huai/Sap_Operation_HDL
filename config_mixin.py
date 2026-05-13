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
import builtins

from PyQt5.QtWidgets import QApplication, QFileDialog, QMainWindow, QMessageBox, QVBoxLayout, QPushButton, QAction, QLabel
from PyQt5.QtCore import QDate
from PyQt5.QtGui import QIcon, QFontDatabase
from Get_Data import *
from PDF_Parser_Utils import extract_company_name, extract_revenue, extract_fapiao_no, parse_pdf_fields, PDF_Operate
from Sap_Function import *
from Data_Table import *
from Logger import *
from Excel_Field_Mapper import excel_field_mapper
from theme_manager_theme import ThemeManager
from Revenue_Operate import *
from auto_updater.config_constants import CURRENT_VERSION
from auto_updater import AutoUpdater, UI_AVAILABLE
from sap import CostOptions, OrderData, OrderService, PartnerOptions, RevenueData, SapConfig, SapSession

def _publish_runtime_globals(source):
    for name in (
        'configFileUrl', 'desktopUrl', 'now', 'last_time', 'today',
        'oneWeekday', 'fileUrl', 'configContent', 'username', 'role',
        'staff_dict', 'monthAbbrev',
    ):
        if name in source:
            setattr(builtins, name, source[name])


class ConfigMixin:
    def getConfig(self):
        # 初始化，获取或生成配置文件
        global configFileUrl
        global desktopUrl
        global now
        global last_time
        global today
        global oneWeekday
        global fileUrl

        date = datetime.datetime.now() + datetime.timedelta(days=1)
        now = int(time.strftime('%Y'))
        last_time = now - 1
        today = time.strftime('%Y.%m.%d')
        oneWeekday = (datetime.datetime.now() + datetime.timedelta(days=7)).strftime('%Y.%m.%d')
        desktopUrl = os.path.join(os.path.expanduser("~"), 'Desktop')
        configFileUrl = '%s/config' % desktopUrl
        _publish_runtime_globals(globals())
        configFile = os.path.exists('%s/config_sap_HDL.csv' % configFileUrl)
        # print(desktopUrl,configFileUrl,configFile)
        if not configFile:  # 判断是否存在文件夹如果不存在则创建为文件夹
            reply = QMessageBox.question(self, '信息', '确认是否要创建配置文件', QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            if reply == QMessageBox.Yes:
                if not os.path.exists(configFileUrl):
                    os.makedirs(configFileUrl)
                self.__class__.createConfigContent(self)
                self.__class__.getConfigContent(self)
                self.textBrowser.append("创建并导入配置成功")
            else:
                exit()
        else:
            self.__class__.getConfigContent(self)
    def getConfigContent(self):
        # 配置文件
        csvFile = pd.read_csv('%s/config_sap_HDL.csv' % configFileUrl, names=['A', 'B', 'C'])
        global configContent
        global username
        global role
        global staff_dict
        configContent = {}
        staff_dict = {}
        # configContent[configContent.get('Business_Department','CS')] = []
        # configContent[configContent.get('Lab_1','PHY')] = []
        # configContent[configContent.get('Lab_2','CHM')] = []
        username = list(csvFile['A'])
        number = list(csvFile['B'])
        role = list(csvFile['C'])
        for i in range(len(username)):
            configContent['%s' % username[i]] = number[i]
            if role[i] == configContent.get('Business_Department', 'CS'):
                # 使用 setdefault 确保键存在且为列表类型
                staff_dict.setdefault(configContent.get('Business_Department', 'CS'), []).append(username[i])
            if role[i] == configContent.get('Lab_1', 'PHY'):
                # 使用 setdefault 确保键存在且为列表类型
                staff_dict.setdefault(configContent.get('Lab_1', 'PHY'), []).append(username[i])
            if role[i] == configContent.get('Lab_2', 'CHM'):
                # 使用 setdefault 确保键存在且为列表类型
                staff_dict.setdefault(configContent.get('Lab_2', 'CHM'), []).append(username[i])
        _publish_runtime_globals(globals())
        self.__class__.csItem(self)
        self.__class__.salesItem(self)
        self.__class__.getDefaultInformation(self)
        self.__class__.getInvoiceMsg(self)

        try:
            self.textBrowser_2.append("配置获取成功")
        except AttributeError:
            QMessageBox.information(self, "提示信息", "已获取配置文件内容", QMessageBox.Yes)
        else:
            pass
    def createConfigContent(self):
        global monthAbbrev
        months = "JanFebMarAprMayJunJulAugSepOctNovDec"
        n = time.strftime('%m')
        pos = (int(n) - 1) * 3
        monthAbbrev = months[pos:pos + 3]

        configContent = [
            ['特殊开票', '内容', '备注'],
            ['SAP_Date_URL', 'N:\\XM Softlines\\6. Personel\\5. Personal\\Supporting Team\\收样\\3.Sap\\ODM Data - XM',
             '文件数据路径'],
            ['Invoice_File_URL',
             'N:\\XM Softlines\\6. Personel\\5. Personal\\Supporting Team\\收样\\3.Sap\\ODM Data - XM\\2.特殊开票',
             '特殊开票文件路径'],
            ['Invoice_File_Name', '特殊开票要求2022.xlsx', '特殊开票文件名称'],
            ['Data数据处理', '内容', '备注'],
            ['Row Data', 'Client Contact Name', '以;分隔，横向添加数据'],
            ['Column Data', 'Project No.;Currency;Amount with VAT;Reference No.', '以;分隔，纵向添加数据'],
            ['Row Check', 0, '是否默认被选中,1选中，0未选中'],
            ['Column Check', 0, '是否默认被选中,1选中，0未选中'],
            ['Combine Key', "CS;Sales;Currency;Material Code;Invoices' name (Chinese);Buyer(GPC);Month;Exchange Rate",
             '以;分隔，数据透视字段'],
            ['SAP登入信息', '内容', '备注'],
            ['Login_msg', 'DR-0486-01->601-240', '订单类型-销售组织-分销渠道-销售办事处-销售组'],
            ['Business_Department', 'CS', '业务部门,名称会用于后续'],
            ['Lab_1', 'PHY', '代表实验室，会用于后续'],
            ['Lab_2', 'CHM', '代表实验室，会用于后续'],
            ['T20', 'PHY', '代表实验室，会用于后续'],
            ['T75', 'CHM', '代表实验室，会用于后续'],
            ['Hourly Rate', '金额', '备注'],
            ['CS_Hourly_Rate', 300, '客服时薪'],
            ['PHY_Hourly_Rate', 300, '物理时薪'],
            ['CHM_Hourly_Rate', 300, '化学时薪'],
            ['成本中心', '编号', '备注'],
            ['CS_Selected', 1, '是否默认被选中,1选中，0未选中'],
            ['PHY_Selected', 1, '是否默认被选中,1选中，0未选中'],
            ['CHM_Selected', 1, '是否默认被选中,1选中，0未选中'],
            ['CS_Cost_Center', '48601240', 'CS成本中心'],
            ['CHM_Cost_Center', '48601293', 'CHM成本中心'],
            ['PHY_Cost_Center', '48601294', 'PHY成本中心'],
            ['计划成本', '数值', '备注'],
            ['Plan_Cost_Parameter', 0.9, '实际的90%，预留10%利润'],
            ['Significant_Digits', 2, '保留几位小数'],
            ['实验室成本比例', '数值', '备注'],
            ['CHM_Cost_Parameter', 0.3, '给到CHM30%'],
            ['PHY_Cost_Parameter', 0.3, '给到PHY30%'],
            # 新增分配规则参数
            ['405_Item_1000', 0.5, '405分配规则'],
            ['405_Item_2000', 0.5, '405分配规则'],
            ['441_Item_1000', 0.8, '441分配规则'],
            ['441_Item_2000', 0.2, '441分配规则'],
            ['430_Item_1000', 0.8, '430分配规则'],
            ['430_Item_2000', 0.2, '430分配规则'],
            # 新增特殊MC规则
            ['T20-430-A2', 'PHY_1000/CHM_2000', '1000/2000对应的lab，强制1000设置在前，2000在后'],
            ['T20-430-A2_mc', 'T20-430-00/T75-430-00', '1000/2000对应的mc'],
            ['T75-441-A2', 'CHM_1000/PHY_2000', '1000/2000对应的lab，强制1000设置在前，2000在后'],
            ['T75-441-A2_mc', 'T75-441-00/T20-441-00', '1000/2000对应的mc'],
            ['T75-405-A2', 'CHM_1000/PHY_2000', '1000/2000对应的lab，强制1000设置在前，2000在后'],
            ['T75-405-A2_mc', 'T75-405-00/T20-405-00', '1000/2000对应的mc'],
            ['T75-405-D2', 'CHM_1000/PHY_2000', '1000/2000对应的lab，计算hour后强制都转为1000'],
            ['T75-405-D2_mc', 'T75-405-D2/T75-405-D2', '1000/2000对应的mc'],
            ['T75-405-D3', 'CHM_1000/PHY_2000', '1000/2000对应的lab，计算hour后强制都转为1000'],
            ['T75-405-D3_mc', 'T75-405-D3/T75-405-D3', '1000/2000对应的mc'],
            ['T75-441-D2', 'CHM_1000/PHY_2000', '1000/2000对应的lab，计算hour后强制都转为1000'],
            ['T75-441-D2_mc', 'T75-441-D2/T75-441-D2', '1000/2000对应的mc'],
            ['T75-441-D3', 'CHM_1000/PHY_2000', '1000/2000对应的lab，计算hour后强制都转为1000'],
            ['T75-441-D3_mc', 'T75-441-D3/T75-441-D3', '1000/2000对应的mc'],
            # 新增公共参数
            ['Max_Hour', 8, '最大工作时长'],
            ['Hours_Combine_Key', "Order Number;Material Code;Primary CS",'以;分隔，数据透视字段'],
            ['Hour_Files_Import_URL', "N:\\XM Softlines\\6. Personel\\5. Personal\\Supporting Team\\2.财务\\2.SAP\\1.ODM Data - XM\\3.Hours",'Invoice文件导入路径'],
            ['Hour_Files_Export_URL', "N:\\XM Softlines\\6. Personel\\5. Personal\\Supporting Team\\2.财务\\2.SAP\\1.ODM Data - XM\\3.Hours",'Invoice文件导入路径'],
            ['Hour_Field_Mapping', "{'staff_id': 'staff_id','week': 'week','order_no': 'order_no','allocated_hours': 'allocated_hours','office_time':'office_time','material_code': 'material_code','item': 'item','allocated_day': 'allocated_day','staff_name': 'staff_name'}", '对应字段映射'],
            ['DATA A数据填写', '判断依据', '备注'],
            ['Data_A_E1', '5010815347;5010427355;5010913488;5010685589;5010829635;5010817524', 'Data A录E1,新添加用;隔开即可'],
            ['Data_A_Z2', '5010908478;5010823259', 'Data A录Z2,新添加用;隔开即可'],
            ['SAP操作', '内容', '备注'],
            ['Cost_VAT_Selected', 1, '是否默认被选中,1选中，0未选中'],
            ['NVA01_Selected', 1, '是否默认被选中,1选中，0未选中'],
            ['NVA02_Selected', 1, '是否默认被选中,1选中，0未选中'],
            ['NVF01_Selected', 0, '是否默认被选中,1选中，0未选中'],
            ['NVF03_Selected', 0, '是否默认被选中,1选中，0未选中'],
            ['DataB_Selected', 1, '是否默认被选中,1选中，0未选中'],
            ['Plan_Cost_Selected', 25, '每月超过几号自动选中（不包含）'],
            ['Save_Selected', 1, '是否默认被选中,1选中，0未选中'],
            ['Every_Selected', 1, '是否默认被选中,1选中，0未选中'],
            ['Contact_Selected', 0, '是否默认被选中,1选中，0未选中'],
            ['管理操作', '内容', '备注'],
            ['Billing_List_URL',
             'N:\\XM Softlines\\6. Personel\\5. Personal\\Supporting Team\\2.财务\\3.Billing存档\\4.XM-billing list\\2023',
             'Billing list文件默认路径'],
            ['Add_CS_Msg_Selected', 1, '是否默认被选中,1选中，0未选中'],
            ['Invoice_No_Selected', 1, '是否默认被选中,1选中，0未选中'],
            ['Invoice_Start_Num', 4, 'Invoice的起始数字'],
            ['Invoice_Num', 9, 'Invoice的总位数'],
            ['Company_Name_Selected', 1, '是否默认被选中,1选中，0未选中'],
            ['Order_No_Selected', 0, '是否默认被选中,1选中，0未选中'],
            ['Invoice_Contact_Selected', 0, '是否默认被选中,1选中，0未选中'],
            ['Order_Start_Num', 7, 'Order的起始数字'],
            ['Order_Num', 9, 'Order的总位数'],
            ['Project_No_Selected', 0, '是否默认被选中,1选中，0未选中'],
            ['Invoice_Name', 'CS + Invoice No + Company Name', 'Invoice文件名称默认规则'],
            ['Invoice_Files_Import_URL', desktopUrl, 'Invoice文件导入路径'],
            ['Invoice_Files_Export_URL', 'N:\\XM Softlines\\1. Project\\3. Finance\\02. WIP', 'Invoice文件导出路径'],
            ['Ele_Invoice_No_Selected', 1, '是否默认被选中,1选中，0未选中'],
            ['Ele_Invoice_Start_Num', 486, '电子发票的起始数字'],
            ['Ele_Invoice_Num', 9, 'Invoice的总位数'],
            ['Ele_Order_No_Selected', 0, '是否默认被选中,1选中，0未选中'],
            ['Ele_Order_Start_Num', 7486, '电子发票的起始数字'],
            ['Ele_Order_Num', 9, 'Order的总位数'],
            ['Ele_Company_Name_Selected', 1, '是否默认被选中,1选中，0未选中'],
            ['Ele_Revenue_Selected', 1, '是否默认被选中,1选中，0未选中'],
            ['Ele_Fapiao_No_Selected', 0, '是否默认被选中,1选中，0未选中'],
            ['Ele_Invoice_Name', 'CS + Company Name + Invoice No + Revenue', '电子发票文件名称默认规则'],
            ['Ele_Invoice_Files_Import_URL', 'N:\\Company Data\\FCO\\11.全电发票', '全电发票路径'],
            ['Ele_Invoice_Files_Export_URL', 'N:\\XM Softlines\\1. Project\\3. Finance\\02. WIP\\全电发票 2023\\10',
             '全电发票导出路径'],
            ['名称', '编号', '角色'],
            ['chen, frank', '6375108', 'CS'],
            ['chen, frank', '6375108', 'Sales'],
        ]
        config = np.array(configContent)
        df = pd.DataFrame(config)
        df.to_csv('%s/config_sap_HDL.csv' % configFileUrl, index=0, header=0, encoding='utf_8_sig')
        self.textBrowser_2.append("配置文件创建成功")
        QMessageBox.information(self, "提示信息",
                                "默认配置文件已经创建好，\n如需修改请在用户桌面查找config文件夹中config_sap_HDL.csv，\n将相应的文件内容替换成用户需求即可，修改后记得重新导入配置文件。",
                                QMessageBox.Yes)
    def exportConfig(self):
        # 重新导出默认配置文件
        reply = QMessageBox.question(self, '信息', '确认是否要创建默认配置文件', QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        if reply == QMessageBox.Yes:
            self.__class__.createConfigContent(self)
        else:
            QMessageBox.information(self, "提示信息", "没有创建默认配置文件，保留原有的配置文件", QMessageBox.Yes)
    def importConfig(self):
        # 重新导入配置文件
        reply = QMessageBox.question(self, '信息', '确认是否要导入配置文件', QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        if reply == QMessageBox.Yes:
            self.__class__.getConfigContent(self)
        else:
            QMessageBox.information(self, "提示信息", "没有重新导入配置文件，将按照原有的配置文件操作", QMessageBox.Yes)
    def getDefaultInformation(self):
        # 默认登录界面信息
        try:
            # data处理
            self.checkBox_17.setChecked(int(configContent['Row Check']))
            self.checkBox_18.setChecked(int(configContent['Column Check']))
            self.lineEdit_23.setText(configContent['Row Data'])
            self.lineEdit_24.setText(configContent['Column Data'])
            self.lineEdit_15.setText(configContent['Combine Key'])
            self.lineEdit_16.setText(configContent['Combine Key'])
            # login信息
            loginMsgList = configContent['Login_msg'].split('-')
            self.lineEdit_10.setText(loginMsgList[0])
            self.lineEdit_11.setText(loginMsgList[1])
            self.lineEdit_12.setText(loginMsgList[2])
            self.lineEdit_13.setText(loginMsgList[3])
            self.lineEdit_14.setText(loginMsgList[4])
            # 每小时成本
            self.doubleSpinBox_5.setValue(float(format(float(configContent['CS_Hourly_Rate']), '.2f')))
            self.doubleSpinBox_6.setValue(float(format(float(configContent['CHM_Hourly_Rate']), '.2f')))
            self.doubleSpinBox_8.setValue(float(format(float(configContent['PHY_Hourly_Rate']), '.2f')))
            # 成本中心
            self.checkBox_13.setChecked(int(configContent['CS_Selected']))
            self.checkBox_14.setChecked(int(configContent['CHM_Selected']))
            self.checkBox_15.setChecked(int(configContent['PHY_Selected']))
            self.lineEdit_18.setText(configContent['CS_Cost_Center'])
            self.lineEdit_19.setText(configContent['CHM_Cost_Center'])
            self.lineEdit_20.setText(configContent['PHY_Cost_Center'])
            # 计划成本
            self.doubleSpinBox_7.setValue(float(format(float(configContent['Plan_Cost_Parameter']), '.2f')))
            self.spinBox_5.setValue(int(configContent['Significant_Digits']))
            # 实验室分配比例
            self.doubleSpinBox_9.setValue(float(format(float(configContent['CHM_Cost_Parameter']), '.2f')))
            self.doubleSpinBox_10.setValue(float(format(float(configContent['PHY_Cost_Parameter']), '.2f')))
            # DATA A选择
            self.lineEdit_21.setText(configContent['Data_A_E1'])
            self.lineEdit_22.setText(configContent['Data_A_Z2'])
            # COST是否含税
            self.checkBox_27.setChecked(int(configContent['Cost_VAT_Selected']))
            # SAP操作
            self.checkBox.setChecked(int(configContent['NVA01_Selected']))
            self.checkBox_2.setChecked(int(configContent['NVA02_Selected']))
            self.checkBox_3.setChecked(int(configContent['NVF01_Selected']))
            self.checkBox_4.setChecked(int(configContent['NVF03_Selected']))
            self.checkBox_7.setChecked(int(configContent['DataB_Selected']))
            self.checkBox_6.setChecked(int(configContent['Save_Selected']))
            self.checkBox_16.setChecked(int(configContent['Every_Selected']))
            self.checkBox_19.setChecked(int(configContent['Contact_Selected']))
            if int(configContent['Plan_Cost_Selected']) < int(today.split('.')[-1]):
                self.checkBox_8.setChecked(True)
            # admin操作
            self.checkBox_25.setChecked(int(configContent['Add_CS_Msg_Selected']))
            self.checkBox_9.setChecked(int(configContent['Invoice_No_Selected']))
            self.spinBox.setValue(int(configContent['Invoice_Start_Num']))
            self.spinBox_2.setValue(int(configContent['Invoice_Num']))
            self.checkBox_10.setChecked(int(configContent['Company_Name_Selected']))
            self.checkBox_12.setChecked(int(configContent['Order_No_Selected']))
            self.checkBox_26.setChecked(int(configContent['Invoice_Contact_Selected']))
            self.spinBox_3.setValue(int(configContent['Order_Start_Num']))
            self.spinBox_4.setValue(int(configContent['Order_Num']))
            self.checkBox_11.setChecked(int(configContent['Project_No_Selected']))
            self.lineEdit_17.setText(configContent['Invoice_Name'])
            self.checkBox_20.setChecked(int(configContent['Ele_Invoice_No_Selected']))
            self.spinBox_8.setValue(int(configContent['Ele_Invoice_Start_Num']))
            self.spinBox_6.setValue(int(configContent['Ele_Invoice_Num']))
            self.checkBox_21.setChecked(int(configContent['Ele_Order_No_Selected']))
            self.spinBox_9.setValue(int(configContent['Ele_Order_Start_Num']))
            self.spinBox_7.setValue(int(configContent['Ele_Order_Num']))
            self.checkBox_22.setChecked(int(configContent['Ele_Company_Name_Selected']))
            self.checkBox_23.setChecked(int(configContent['Ele_Fapiao_No_Selected']))
            self.checkBox_24.setChecked(int(configContent['Ele_Revenue_Selected']))
            self.lineEdit_27.setText(configContent['Ele_Invoice_Name'])
            # hour界面操作
            self.spinBox_10.setValue(int(configContent['Max_Hour']))
            self.lineEdit_39.setText(configContent['Hours_Combine_Key'])
            today_hours = datetime.date.today()
            first_day = today_hours.replace(day=1)
            self.dateEdit.setDate(QDate(first_day.year, first_day.month, first_day.day))  # 当月第一天
            self.dateEdit_2.setDate(QDate.currentDate())  # 当天日期
            self.doubleSpinBox_14.setValue(float(format(float(configContent['CS_Hourly_Rate']), '.2f')))
            self.doubleSpinBox_16.setValue(float(format(float(configContent['CHM_Hourly_Rate']), '.2f')))
            self.doubleSpinBox_15.setValue(float(format(float(configContent['PHY_Hourly_Rate']), '.2f')))
            self.doubleSpinBox_11.setValue(float(format(float(configContent['Plan_Cost_Parameter']), '.2f')))
            self.doubleSpinBox_13.setValue(float(format(float(configContent['CHM_Cost_Parameter']), '.2f')))
            self.doubleSpinBox_12.setValue(float(format(float(configContent['PHY_Cost_Parameter']), '.2f')))
            self.spinBox_11.setValue(int(configContent['Significant_Digits']))
            self.lineEdit_28.setText(configContent['Lab_1'])
            self.lineEdit_29.setText(configContent['Lab_2'])
        except Exception as msg:
            self.textBrowser_2.append("错误信息：%s" % msg)
            self.textBrowser_2.append('----------------------------------')
            QApplication.processEvents()
            reply = QMessageBox.question(self, '信息', '错误信息：%s。\n是否要重新创建配置文件' % msg, QMessageBox.Yes | QMessageBox.No,
                                         QMessageBox.Yes)
            if reply == QMessageBox.Yes:
                self.__class__.createConfigContent(self)
                self.textBrowser.append("创建并导入配置成功")
                self.textBrowser_2.append('----------------------------------')
                QApplication.processEvents()
    def csItem(self):
        self.comboBox_2.clear()
        self.comboBox_2.addItem('')
        nameList = username
        i = 0
        for each in nameList:
            if role[i] == 'CS':
                self.comboBox_2.addItem(each)
            i += 1
            QApplication.processEvents()
    def salesItem(self):
        self.comboBox_3.clear()
        self.comboBox_3.addItem('')
        self.comboBox_3.addItem('')
        nameList = username
        i = 0
        for each in nameList:
            if role[i] == 'Sales':
                self.comboBox_3.addItem(each)
            i += 1
            QApplication.processEvents()
    def getAmountVat(self):
        amount = float(self.doubleSpinBox_2.text())
        self.doubleSpinBox_4.setValue(amount * 1.06)
    def getGuiData(self):
        guiData = {}
        guiData['sapNo'] = self.lineEdit.text()
        guiData['projectNo'] = self.lineEdit_2.text()
        guiData['currencyType'] = self.comboBox.currentText()
        guiData['exchangeRate'] = float(self.doubleSpinBox.text())
        guiData['globalPartnerCode'] = self.lineEdit_3.text()
        guiData['csName'] = self.comboBox_2.currentText()
        if guiData['csName'] != '':
            guiData['csCode'] = configContent[guiData['csName']]
        guiData['salesName'] = self.comboBox_3.currentText()
        if guiData['salesName'] != '':
            guiData['salesCode'] = configContent[guiData['salesName']]
        guiData['amount'] = float(self.doubleSpinBox_2.text())
        if self.checkBox_27.isChecked():
            guiData['cost'] = float(self.doubleSpinBox_3.text())/1.06
        else:
            guiData['cost'] = float(self.doubleSpinBox_3.text())
        guiData['amountVat'] = float(self.doubleSpinBox_4.text())
        guiData['csHourlyRate'] = float(self.doubleSpinBox_5.text())
        guiData['chmHourlyRate'] = float(self.doubleSpinBox_6.text())
        guiData['phyHourlyRate'] = float(self.doubleSpinBox_8.text())
        guiData['shortText'] = self.lineEdit_5.text()
        guiData['materialCode'] = getattr(self, 'current_material_code', '')
        guiData['longText'] = getattr(self, 'current_long_text', '')
        guiData['planCostRate'] = float(self.doubleSpinBox_7.text())
        guiData['significantDigits'] = int(self.spinBox_5.text())
        guiData['chmCostRate'] = float(self.doubleSpinBox_9.text())
        guiData['phyCostRate'] = float(self.doubleSpinBox_10.text())
        guiData['dataAE1'] = self.lineEdit_21.text().split(';')
        guiData['dataAZ2'] = self.lineEdit_22.text().split(';')
        guiData['invoiceStsrtNum'] = int(self.spinBox.text())
        guiData['invoiceBits'] = int(self.spinBox_2.text())
        guiData['orderStsrtNum'] = int(self.spinBox_3.text())
        guiData['orderBits'] = int(self.spinBox_4.text())
        guiData['pdfName'] = self.lineEdit_17.text()
        guiData['orderType'] = self.lineEdit_10.text()
        guiData['salesOrganization'] = self.lineEdit_11.text()
        guiData['distributionChannels'] = self.lineEdit_12.text()
        guiData['salesOffice'] = self.lineEdit_13.text()
        guiData['salesGroup'] = self.lineEdit_14.text()
        guiData['csCostCenter'] = self.lineEdit_18.text()
        guiData['chmCostCenter'] = self.lineEdit_19.text()
        guiData['phyCostCenter'] = self.lineEdit_20.text()
        if self.checkBox.isChecked():
            guiData['va01Check'] = True
        else:
            guiData['va01Check'] = False

        if self.checkBox_2.isChecked():
            guiData['va02Check'] = True
        else:
            guiData['va02Check'] = False

        if self.checkBox_3.isChecked():
            guiData['vf01Check'] = True
        else:
            guiData['vf01Check'] = False

        if self.checkBox_4.isChecked():
            guiData['vf03Check'] = True
        else:
            guiData['vf03Check'] = False

        if self.checkBox_6.isChecked():
            guiData['saveCheck'] = True
        else:
            guiData['saveCheck'] = False

        if self.checkBox_7.isChecked():
            guiData['labCostCheck'] = True
        else:
            guiData['labCostCheck'] = False

        if self.checkBox_8.isChecked():
            guiData['planCostCheck'] = True
        else:
            guiData['planCostCheck'] = False

        if self.checkBox_13.isChecked():
            guiData['csCheck'] = True
        else:
            guiData['csCheck'] = False

        if self.checkBox_14.isChecked():
            guiData['chmCheck'] = True
        else:
            guiData['chmCheck'] = False

        if self.checkBox_15.isChecked():
            guiData['phyCheck'] = True
        else:
            guiData['phyCheck'] = False

        if self.checkBox_16.isChecked():
            guiData['everyCheck'] = True
        else:
            guiData['everyCheck'] = False

        if self.checkBox_17.isChecked():
            guiData['rowCheck'] = True
        else:
            guiData['rowCheck'] = False

        if self.checkBox_18.isChecked():
            guiData['columnCheck'] = True
        else:
            guiData['columnCheck'] = False

        if self.checkBox_19.isChecked():
            guiData['contactCheck'] = True
        else:
            guiData['contactCheck'] = False

        return guiData
    def getAdminGuiData(self):
        guiAdminData = {}
        # guiAdminData['BillingList'] = self.lineEdit_25.text()
        guiAdminData['invoiceStsrtNum'] = int(self.spinBox.text())
        guiAdminData['invoiceBits'] = int(self.spinBox_2.text())
        guiAdminData['orderStsrtNum'] = int(self.spinBox_3.text())
        guiAdminData['orderBits'] = int(self.spinBox_4.text())
        guiAdminData['pdfName'] = self.lineEdit_17.text()
        guiAdminData['eleInvoiceStsrtNum'] = int(self.spinBox_8.text())
        guiAdminData['eleOrderStsrtNum'] = int(self.spinBox_9.text())
        guiAdminData['eleInvoiceBits'] = int(self.spinBox_6.text())
        guiAdminData['eleOrderBits'] = int(self.spinBox_7.text())
        guiAdminData['fapiaoName'] = self.lineEdit_27.text()
        return guiAdminData
    def getHourGuiData(self):
        guiHourData = {}
        guiHourData['Hours_Combine_Key'] = self.lineEdit_39.text()
        guiHourData['Max_Hour'] = int(self.spinBox_10.text())
        guiHourData['CS_Hourly_Rate'] = float(self.doubleSpinBox_14.text())
        guiHourData['CHM_Hourly_Rate'] = float(self.doubleSpinBox_16.text())
        guiHourData['PHY_Hourly_Rate'] = float(self.doubleSpinBox_15.text())
        guiHourData['Plan_Cost_Parameter'] = float(self.doubleSpinBox_11.text())
        guiHourData['Significant_Digits'] = float(self.spinBox_11.text())
        guiHourData['CHM_Cost_Parameter'] = float(self.doubleSpinBox_13.text())
        guiHourData['PHY_Cost_Parameter'] = float(self.doubleSpinBox_12.text())
        guiHourData['Lab_1'] = self.lineEdit_28.text()
        guiHourData['Lab_2'] = self.lineEdit_29.text()
        guiHourData['Business_Department'] = self.lineEdit_26.text()
        guiHourData['Start_Date'] = self.dateEdit.date().toString("yyyy.MM.dd")
        guiHourData['End_Date'] = self.dateEdit_2.date().toString("yyyy.MM.dd")
        return guiHourData

