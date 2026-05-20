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
from runtime_globals import configContent, desktopUrl, myTable, myWin, today

class OdmInvoiceMixin:
    def getFile(self, path):
        selectBatchFile = QFileDialog.getOpenFileName(self, '选择ODM导出文件',
                                                      '%s\\%s' % (path, today),
                                                      'files(*.docx;*.xls*;*.csv)')
        fileUrl = selectBatchFile[0]
        return fileUrl
    def getFileUrl(self):
        fileUrl = self.__class__.getFile(self, configContent['SAP_Date_URL'])
        if fileUrl:
            self.lineEdit_6.setText(fileUrl)
            QApplication.processEvents()
        else:
            self.textBrowser.append("请重新选择ODM文件")
            QMessageBox.information(self, "提示信息", "请重新选择ODM文件", QMessageBox.Yes)
    def viewOdmData(self):
        fileUrl = self.lineEdit_6.text()
        odm_data_obj = Get_Data()
        df = odm_data_obj.getFileData(fileUrl)
        myTable.createTable(df)
        myTable.showMaximized()
    def getFileName(self, fileUrl, fileName, fileType):
        nowTime = time.strftime('%Y-%m-%d %H.%M.%S')
        fileName = fileUrl + '/' + nowTime + ' - ' + fileName + '.' + fileType
        return fileName
    def createFolder(self, url):
        isExists = os.path.exists(url)
        if not isExists:
            os.makedirs(url)
    def pdfNameRule(self, msg, flag):
        guiData = self.__class__.getAdminGuiData(self)
        if flag == 'invoice':
            pdfName = guiData['pdfName']
        else:
            pdfName = guiData['fapiaoName']
        pdfNameList = pdfName.split(' + ')
        changedPdfName = ''
        if msg in pdfNameList:
            pdfNameList.remove(msg)
            for each in pdfNameList:
                if changedPdfName != '':
                    changedPdfName += ' + '
                changedPdfName += each
        else:
            changedPdfName += pdfName
            if changedPdfName != '':
                changedPdfName += ' + '
            changedPdfName += msg
        if flag == 'invoice':
            self.lineEdit_17.setText(changedPdfName)
        else:
            self.lineEdit_27.setText(changedPdfName)
        QApplication.processEvents()
        return changedPdfName
    def getPdfFiles(self):
        # 获取invoice文件
        selectBatchFile = QFileDialog.getOpenFileNames(self, '选择文件', '%s' % configContent['Invoice_Files_Import_URL'],
                                                       'files(*.pdf)')
        self.filesUrl = selectBatchFile[0]
        if self.filesUrl != []:
            self.textBrowser_3.append('选中文件:')
            self.textBrowser_3.append('\n'.join(self.filesUrl))
            self.textBrowser_3.append('----------------------------------')
        else:
            self.textBrowser_3.append('无选中文件')
            self.textBrowser_3.append('----------------------------------')
        QApplication.processEvents()
        return self.filesUrl
    def invoiceRenameOperate(self):
        # invoice中的pdf重新命名
        fileUrls = self.filesUrl
        flag = 'Y'
        if fileUrls == []:
            reply = QMessageBox.question(self, '信息', '没有选中文件，是否重新选择文件', QMessageBox.Yes | QMessageBox.No,
                                         QMessageBox.Yes)
            if reply == QMessageBox.Yes:
                fileUrls = self.__class__.getPdfFiles(self)
                if fileUrls == []:
                    flag = 'N'
            else:
                QMessageBox.information(self, "提示信息", "没有选中文件，请重新选择文件", QMessageBox.Yes)
                flag = 'N'
        if flag == 'Y':
            try:
                if self.checkBox_25.isChecked() and self.lineEdit_25.text() == '':
                    self.textBrowser_3.append('未选择Billing List文件')
                    self.textBrowser_3.append('----------------------------------')
                    QApplication.processEvents()
                else:
                    guiData = self.__class__.getAdminGuiData(self)
                    pdfOperate = PDF_Operate
                    self.textBrowser_3.append('导出文件夹：%s' % configContent['Invoice_Files_Export_URL'])
                    self.textBrowser_3.append('导出文件名称：')
                    if self.checkBox_25.isChecked():
                        billing_df = myWin.getBillingListData()
                    i = 1
                    # 新增加log
                    log_file_name = 'Invoice %s.xlsx' % time.strftime('%Y-%m-%d %H.%M.%S')
                    Log_file = os.path.join(configContent['Invoice_Files_Export_URL'], log_file_name)
                    log_obj = Logger(Log_file, ['Update', 'Invoice No', 'File Name', 'Company Name', 'CS', 'Project No', 'Customer Name', 'Client Contact Name', 'Remark'])
                    has_error = False
                    for fileUrl in fileUrls:
                        try:
                            log_list = {}

                            self.textBrowser_3.append('第%s份文件：' % i)
                            msg = {}
                            msg['Invoice No'] = ''
                            with open(fileUrl, 'rb') as pdfFile:
                                fileCon = pdfOperate.readPdf(pdfFile)
                                fileNum = 0
                                for fileCon[fileNum] in fileCon:
                                    if re.match('.*P. R. China', fileCon[fileNum]) or re.match('.*P.R. China', fileCon[fileNum]) or re.match('Pleasequotethisnumberonallinquiriesandpayments', fileCon[fileNum]) or re.match('请在项目咨询或付款时提示此帐单号', fileCon[fileNum]) or re.match(
                                'Please quote this number on all inquiries and payments.', fileCon[fileNum]):
                                        # DG还是+1，包含invoice no；NB是invoice no后是公司名称
                                        if '487' in str(guiData['invoiceStsrtNum']):
                                            msg['Company Name'] = fileCon[fileNum + 2].replace(
                                                'Please quote this number on all inquiries and payments.', '').replace(
                                                'Invoice No.', '')
                                        else:
                                            # XM+1，不包含invoice no；DG还是+1，包含invoice no；
                                            msg['Company Name'] = fileCon[fileNum + 1].replace(
                                                'Please quote this number on all inquiries and payments.', '').replace(
                                                'Invoice No.', '')
                                    elif re.match('请在项目咨询或付款时提示此帐单号', fileCon[fileNum]):
                                        msg['Company Name'] = fileCon[fileNum + 2].replace(
                                            'Please quote this number on all inquiries and payments.', '').replace(
                                            'Invoice No.', '')
                                    elif re.match(r'%s\d{%s}' % (guiData['invoiceStsrtNum'], int(guiData['invoiceBits']) - len(
                                            str(guiData['invoiceStsrtNum']))),
                                                  fileCon[fileNum]):
                                        msg['Invoice No'] = fileCon[fileNum]
                                    elif re.search(r'\d{2}.\d{3}.\d{2}.\d{4,5}', fileCon[fileNum]):
                                        res = fileCon[fileNum].split(' ')
                                        for each in res:
                                            if re.search(r'\d{2}.\d{3}.\d{2}.\d{4,5}', each):
                                                msg['Project No'] = re.findall(r'\d{2}.\d{3}.\d{2}.\d{4,5}..*$', each)[0]
                                            elif re.search(
                                                    r'%s\d{%s}' % (guiData['invoiceStsrtNum'], int(guiData['invoiceBits']) - len(str(guiData['invoiceStsrtNum']))),
                                                    each) and msg['Invoice No'] == '':
                                                msg['Invoice No'] = each
                                    elif re.search(r'%s\d{%s}' % (
                                            guiData['orderStsrtNum'],
                                            int(guiData['orderBits']) - len(str(guiData['orderStsrtNum']))),
                                                   fileCon[fileNum]):
                                        res = fileCon[fileNum].split(' ')
                                        if len(res[1]) == int(guiData['orderBits']):
                                            msg['Order No'] = res[1]
                                    elif 'Client Contact Name' in fileCon[fileNum] or 'ClientContactName' in fileCon[fileNum]:
                                        msg['Client Contact Name'] = fileCon[fileNum].replace('Client Contact Name:', '').replace('ClientContactName:', '').replace('/', '&')
                                    fileNum += 1
                                if 'Project No' not in msg:
                                    msg['Project No'] =''
                                if 'Client Contact Name' not in msg:
                                    msg['Client Contact Name'] = ''
                                if self.checkBox_25.isChecked():
                                    cs = list(billing_df[billing_df['Final Invoice No.'].astype('int64') == int(msg['Invoice No'])]['CS'])
                                    # NB的Company Name需要在这边操作
                                    customerName = list(
                                        billing_df[
                                            billing_df['Final Invoice No.'].astype('int64') == int(msg['Invoice No'])][
                                            'Customer Name'])
                                    if cs == []:
                                        msg['CS'] = ''
                                        msg['Customer Name'] = ''
                                    else:
                                        msg['CS'] = cs[0]
                                        msg['Customer Name'] = customerName[0]
                                    if 'Company Name' not in msg:
                                        msg['Company Name'] = msg['Customer Name']
                                if msg['Invoice No'] in msg['Company Name']:
                                    msg['Company Name'] = msg['Company Name'].replace(' %s' % msg['Invoice No'], '')
                                pdfNameRule = guiData['pdfName'].split(' + ')
                                outputFlieName = ''
                                for eachName in pdfNameRule:
                                    if outputFlieName != '':
                                        outputFlieName += '-'
                                    if eachName == 'Invoice No':
                                        outputFlieName += msg['Invoice No']
                                    elif eachName == 'Company Name':
                                        outputFlieName += msg['Company Name']
                                    elif eachName == 'Order No':
                                        outputFlieName += msg['Order No']
                                    elif eachName == 'Project No':
                                        outputFlieName += msg['Project No']
                                    elif eachName == 'CS':
                                        outputFlieName += msg['CS']
                                    elif eachName == 'Client Contact Name':
                                        outputFlieName += msg['Client Contact Name']
                                outputFlie = outputFlieName + '.pdf'
                                # 替换非法字符后的文件名字符串
                                outputFlie = PDF_Operate.sanitize_filename(outputFlie)
                                log_list['Invoice No'] = msg['Invoice No']
                                log_list['Company Name'] = msg['Company Name']
                                log_list['Project No'] = msg['Project No']
                                log_list['Client Contact Name'] = msg['Client Contact Name']
                                log_list['File Name'] = outputFlie
                                if 'CS' in msg:
                                    log_list['CS'] = msg['CS']
                                    log_list['Customer Name'] = msg['Customer Name']
                                else:
                                    log_list['CS'] = ''
                                    log_list['Customer Name'] = ''
                                if 'Remark' in msg:
                                    log_list['Remark'] = msg['Remark']
                                else:
                                    log_list['Remark'] = ''
                                log_obj.log(log_list)
                                # outputFlie = msg['Invoice No'] + '-' + msg['Company Name'] + '.pdf'
                                pdfOperate.saveAs(r'%s' % fileUrl, '%s\\%s' % (configContent['Invoice_Files_Export_URL'], outputFlie))
                            self.textBrowser_3.append('%s' % outputFlie)
                            QApplication.processEvents()
                        except Exception as errorMsg:
                            # self.textBrowser_3.append("<font color='red'>第%s份文件：</font>" % i)
                            self.textBrowser_3.append("<font color='red'>出错信息：%s </font>" % errorMsg)
                            self.textBrowser_3.append("<font color='red'>出错的文件：%s </font>" % fileUrl)
                            try:
                                # 创建错误文件目录
                                error_folder = os.path.join(desktopUrl, "error_invoices")
                                error_folder = os.path.normpath(error_folder)  # 新增路径规范化
                                os.makedirs(error_folder, exist_ok=True)
                                # 保存出错文件副本
                                error_file = os.path.join(error_folder, os.path.basename(fileUrl))
                                error_file = os.path.normpath(error_file)  # 新增路径规范化

                                # 添加文件存在检查
                                if os.path.exists(fileUrl):
                                    has_error = True  # 设置错误标志
                                    shutil.copy(fileUrl, error_file)
                                    self.textBrowser_3.append(f"已将出错文件复制到：{error_file}")
                                else:
                                    self.textBrowser_3.append(f"<font color='red'>源文件不存在: {fileUrl}</font>")

                            except PermissionError as perm_err:
                                self.textBrowser_3.append(f"<font color='red'>权限拒绝错误: {perm_err}</font>")
                                self.textBrowser_3.append(
                                    "<font color='red'>请检查：1.是否具有网络路径写入权限 2.文件是否被其他程序占用</font>")
                        i += 1
                    log_obj.save_log_to_excel()
                    os.startfile(Log_file)
                    os.startfile(configContent['Invoice_Files_Export_URL'])
                    self.textBrowser_3.append('生成的数据文件：%s' % Log_file)
                    self.textBrowser_3.append('----------------------------------')
                    if has_error:
                        os.startfile(error_folder)
            except Exception as errorMsg:
                log_obj.save_log_to_excel()
                self.textBrowser_3.append("<font color='red'>出错信息：%s </font>" % errorMsg)
                QApplication.processEvents()
    def getEleInvoiceFiles(self):
        # 获取电子发票文件
        selectBatchFile = QFileDialog.getOpenFileNames(self, '选择文件',
                                                       '%s' % configContent['Ele_Invoice_Files_Import_URL'],
                                                       'files(*.pdf)')
        self.filesUrl = selectBatchFile[0]
        if self.filesUrl != []:
            self.textBrowser_3.append('选中文件:')
            self.textBrowser_3.append('\n'.join(self.filesUrl))
            self.textBrowser_3.append('----------------------------------')
        else:
            self.textBrowser_3.append('无选中文件')
            self.textBrowser_3.append('----------------------------------')
        QApplication.processEvents()
        return self.filesUrl
    def electronicInvoice(self):
        # 需要名称，金额，发票号，税号
        # Excel或csv数据保留对吧billing list的金额
        fileUrls = self.filesUrl
        flag = 'Y'
        if fileUrls == []:
            reply = QMessageBox.question(self, '信息', '没有选中文件，是否重新选择文件', QMessageBox.Yes | QMessageBox.No,
                                         QMessageBox.Yes)
            if reply == QMessageBox.Yes:
                fileUrls = self.__class__.getPdfFiles(self)
                if fileUrls == []:
                    flag = 'N'
            else:
                QMessageBox.information(self, "提示信息", "没有选中文件，请重新选择文件", QMessageBox.Yes)
                flag = 'N'
        if flag == 'Y':
            if self.checkBox_25.isChecked() and self.lineEdit_25.text() == '':
                self.textBrowser_3.append('未选择Billing List文件')
                self.textBrowser_3.append('----------------------------------')
                QApplication.processEvents()
            else:
                log_file_name = 'Electronic Invoice %s.xlsx' % time.strftime('%Y-%m-%d %H.%M.%S')
                log_file = os.path.join(configContent['Ele_Invoice_Files_Export_URL'], log_file_name)
                log_obj = Logger(log_file, [
                    'Update',
                    'id',
                    'Company Name',
                    'Invoice No',
                    'File Name',
                    'Order No',
                    'Revenue',
                    'fapiao',
                    'update',
                    'path',
                    'CS',
                    'ODM Revenue',
                    'ODM Customer Name',
                    'ODMRe - Re',
                    '判断客户名称是否正确'
                ])
                has_error = False
                try:
                    pdfOperate = PDF_Operate
                    adminGuiData = self.__class__.getAdminGuiData(self)
                    billing_dict = {}  # 初始化字典，避免未定义警告
                    self.textBrowser_3.append('导出文件夹：%s' % configContent['Ele_Invoice_Files_Export_URL'])
                    self.textBrowser_3.append('导出文件名称：')
                    if self.checkBox_25.isChecked():
                        billing_df = myWin.getBillingListData()
                        # 性能优化：预先创建 Billing List 字典映射，避免循环中重复扫描 DataFrame
                        # 一次性类型转换，避免每次查询都重新转换
                        billing_df['Final Invoice No.'] = billing_df['Final Invoice No.'].astype('int64')
                        # 创建字典映射：{invoice_no: {cs, odmRe, customerName}}
                        billing_dict = {}
                        for _, row in billing_df.iterrows():
                            invoice_no = row['Final Invoice No.']
                            billing_dict[invoice_no] = {
                                'CS': row['CS'],
                                'odmRe': row['求和项:Amount with VAT'],
                                'CustomerName': row['Customer Name']
                            }
                        self.textBrowser_3.append(f'✓ 已加载 Billing List 索引：{len(billing_dict)} 条记录')
                        self.textBrowser_3.append('----------------------------------')
                    i = 1

                    # 构造正则表达式
                    inv_pattern = r'%s\d{%s}' % (
                        adminGuiData['eleInvoiceStsrtNum'],
                        int(adminGuiData['eleInvoiceBits']) - len(
                            str(adminGuiData['eleInvoiceStsrtNum']))
                    )
                    inv_pattern = r'(?<!\d)' + inv_pattern + r'(?!\d)'

                    order_pattern = r'%s\d{%s}' % (
                        adminGuiData['eleOrderStsrtNum'],
                        int(adminGuiData['eleOrderBits']) - len(
                            str(adminGuiData['eleOrderStsrtNum']))
                    )

                    for fileUrl in fileUrls:
                        try:
                            self.textBrowser_3.append('第%s份文件：' % i)
                            msg = {}
                            with (open(fileUrl, 'rb') as pdfFile):
                                fileCon = pdfOperate.readPdf(pdfFile)

                                for num, each in enumerate(fileCon):
                                    # 使用统一的PDF字段解析函数 (优化后更简洁)
                                    parse_pdf_fields(msg, each, inv_pattern, order_pattern)

                                if 'Order No' not in msg:
                                    msg['Order No'] = ''
                                if 'Invoice No' not in msg:
                                    msg['Invoice No'] = re.findall(inv_pattern,fileUrl)[0]
                                if self.checkBox_25.isChecked():
                                    # 性能优化：使用字典查询替代 DataFrame 扫描，O(1) 复杂度
                                    try:
                                        invoice_no = int(msg['Invoice No'])
                                        invoice_data = billing_dict.get(invoice_no)
                                        if invoice_data is None:
                                            msg['CS'] = ''
                                            msg['odmRe'] = 0.00
                                            msg['Customer Name'] = ''
                                        else:
                                            msg['CS'] = invoice_data['CS']
                                            msg['odmRe'] = invoice_data['odmRe']
                                            msg['Customer Name'] = invoice_data['CustomerName']
                                    except (ValueError, KeyError):
                                        # Invoice No 格式错误或缺失
                                        msg['CS'] = ''
                                        msg['odmRe'] = 0.00
                                        msg['Customer Name'] = ''
                                else:
                                    msg['CS'] = ''
                                    msg['odmRe'] = 0.00
                                    msg['Customer Name'] = ''
                                # 文件命名规则
                                pdfNameRule = adminGuiData['fapiaoName'].split(' + ')
                                outputFlieName = ''
                                for eachName in pdfNameRule:
                                    if outputFlieName != '':
                                        outputFlieName += '-'
                                    if eachName == 'Invoice No':
                                        outputFlieName += str(msg['Invoice No'])
                                    elif eachName == 'Company Name':
                                        outputFlieName += msg['Company Name']
                                    elif eachName == 'Order No':
                                        outputFlieName += str(msg['Order No'])
                                    elif eachName == 'FaPiao No':
                                        outputFlieName += msg['fapiao']
                                    elif eachName == 'Revenue':
                                        outputFlieName += msg['Revenue']
                                    elif eachName == 'CS':
                                        outputFlieName += msg['CS']

                                # outputFlieName = msg['Company Name'] + '-' + msg['Revenue'] + '-' + msg['fapiao']
                                outputFlie = outputFlieName + '.pdf'
                                exportUrl = configContent['Ele_Invoice_Files_Export_URL']
                                self.createFolder(exportUrl)
                                pdfOperate.saveAs(fileUrl, '%s\\%s' % (exportUrl, outputFlie))
                                log_list = {
                                    'id': i,
                                    'Company Name': msg.get('Company Name', ''),
                                    'Invoice No': msg.get('Invoice No', ''),
                                    'File Name': outputFlie,
                                    'Order No': msg.get('Order No', ''),
                                    'Revenue': float(msg.get('Revenue', '0.0').replace('¥', '')),
                                    'fapiao': msg.get('fapiao', ''),
                                    'update': time.strftime('%Y-%m-%d %H.%M.%S'),
                                    'path': '%s\\%s' % (exportUrl, outputFlie),
                                    'CS': msg.get('CS', ''),
                                    'ODM Revenue': msg.get('odmRe', 0.00),
                                    'ODM Customer Name': msg.get('Customer Name', ''),
                                    'ODMRe - Re': msg.get('odmRe', 0.00) - float(
                                        msg.get('Revenue', '0.0').replace('¥', '')),
                                    '判断客户名称是否正确': msg.get('Company Name', '') == msg.get('Customer Name', '')
                                }
                                log_obj.log(log_list)
                            self.textBrowser_3.append('%s' % outputFlie)
                            QApplication.processEvents()

                        except Exception as errorMsg:

                            self.textBrowser_3.append("<font color='red'>出错信息：%s </font>" % errorMsg)
                            self.textBrowser_3.append("<font color='red'>出错的文件：%s </font>" % fileUrl)
                            QApplication.processEvents()
                            try:
                                # 创建错误文件目录
                                error_folder = os.path.join(desktopUrl, "error_invoices")
                                error_folder = os.path.normpath(error_folder)  # 新增路径规范化
                                os.makedirs(error_folder, exist_ok=True)
                                # 保存出错文件副本
                                error_file = os.path.join(error_folder, os.path.basename(fileUrl))
                                error_file = os.path.normpath(error_file)  # 新增路径规范化

                                # 添加文件存在检查
                                if os.path.exists(fileUrl):
                                    has_error = True  # 设置错误标志
                                    shutil.copy(fileUrl, error_file)
                                    self.textBrowser_3.append(f"已将出错文件复制到：{error_file}")
                                else:
                                    self.textBrowser_3.append(f"<font color='red'>源文件不存在: {fileUrl}</font>")

                            except PermissionError as perm_err:
                                self.textBrowser_3.append(f"<font color='red'>权限拒绝错误: {perm_err}</font>")
                                self.textBrowser_3.append(
                                    "<font color='red'>请检查：1.是否具有网络路径写入权限 2.文件是否被其他程序占用</font>")
                            QApplication.processEvents()
                        i += 1

                    log_obj.save_log_to_excel()
                    self.textBrowser_3.append('已生成数据文件：%s' % log_file)
                    self.textBrowser_3.append('----------------------------------')
                    os.startfile(configContent['Ele_Invoice_Files_Export_URL'])
                    os.startfile(log_file)
                    if has_error:
                        os.startfile(error_folder)
                except Exception as errorMsg:
                    log_obj.save_log_to_excel()
                    os.startfile(log_file)
                    self.textBrowser_3.append("<font color='red'>出错信息：%s </font>" % errorMsg)
                    QApplication.processEvents()
    def getBillingListFile(self):
        try:
            selectBatchFile = QFileDialog.getOpenFileName(self, '选择文件',
                                                          '%s' % configContent['Billing_List_URL'],
                                                          'files(*.xlsx)')
            fileUrl = selectBatchFile[0]
            if fileUrl:
                self.lineEdit_25.setText(fileUrl)
                QApplication.processEvents()
                self.textBrowser_3.append('选中Billing List文件：%s' % fileUrl)
                self.textBrowser_3.append('----------------------------------')
            else:
                self.textBrowser_3.append('无选中文件')
                self.textBrowser_3.append('----------------------------------')
            QApplication.processEvents()
            return fileUrl
        except Exception as errorMsg:
            self.textBrowser_3.append("<font color='red'>出错信息：%s </font>" % errorMsg)
            QApplication.processEvents()
            return
    def getBillingListData(self, sheet_name=[]):
        try:
            billing_list_url = self.lineEdit_25.text()
            if billing_list_url == '':
                self.textBrowser_3.append('无选中文件')
                self.textBrowser_3.append('----------------------------------')
                QApplication.processEvents()
                return None
            else:
                billing_list_obj = Get_Data()
                billing_list_data = billing_list_obj.getFileMoreSheetData(billing_list_url, sheet_name)
                pivotTableKey = ['Final Invoice No.', 'Customer Name', 'CS', 'Cur.']
                valusKey = ['求和项:Amount with VAT']
                billing_list_data = billing_list_obj.pivotTable(pivotTableKey, valusKey)
                billing_list_data = billing_list_data.reset_index()
                return billing_list_data
        except Exception as errorMsg:
            self.textBrowser_3.append("<font color='red'>出错信息：%s </font>" % errorMsg)
            QApplication.processEvents()
            return
    def viewBillingListData(self):
        fileUrl = self.lineEdit_25.text()
        if fileUrl:
            sheet_name = []
            df = myWin.getBillingListData(sheet_name)
            try:
                myTable.createTable(df)
                myTable.showMaximized()
            except Exception as errorMsg:
                self.textBrowser_3.append('数据有问题%s' % errorMsg)
                self.textBrowser_3.append('----------------------------------')
                QApplication.processEvents()
        else:
            self.textBrowser_3.append('无选中文件')
            self.textBrowser_3.append('----------------------------------')
            QApplication.processEvents()
    def open_file(self, path):
        os.startfile(path)

