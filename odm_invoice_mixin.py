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
from Sap_Function import *
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
    def getODMDataFileUrl(self):
        fileUrl = self.__class__.getFile(self, configContent['SAP_Date_URL'])
        if fileUrl:
            self.lineEdit_7.setText(fileUrl)
            QApplication.processEvents()
        else:
            self.textBrowser_2.append("请重新选择ODM文件")
            QMessageBox.information(self, "提示信息", "请重新选择ODM文件", QMessageBox.Yes)
    def getCombineFileUrl(self):
        fileUrl = self.__class__.getFile(self, configContent['SAP_Date_URL'])
        if fileUrl:
            self.lineEdit_8.setText(fileUrl)
            QApplication.processEvents()
        else:
            self.textBrowser_2.append("请重新选择ODM文件")
            QMessageBox.information(self, "提示信息", "请重新选择ODM文件", QMessageBox.Yes)
    def getLogFileUrl(self):
        fileUrl = self.__class__.getFile(self, configContent['SAP_Date_URL'])
        if fileUrl:
            self.lineEdit_9.setText(fileUrl)
            QApplication.processEvents()
        else:
            self.textBrowser_2.append("请重新选择ODM文件")
            QMessageBox.information(self, "提示信息", "请重新选择ODM文件", QMessageBox.Yes)
    def viewOdmData(self):
        fileUrl = self.lineEdit_6.text()
        odm_data_obj = Get_Data()
        df = odm_data_obj.getFileData(fileUrl)
        myTable.createTable(df)
        myTable.showMaximized()
    def odmDataToSap(self):
        try:
            fileUrl = self.lineEdit_6.text()
            (filepath, filename) = os.path.split(fileUrl)
            if fileUrl:
                # 下拉框默认选择0
                self.comboBox.setCurrentIndex(0)
                self.comboBox_2.setCurrentIndex(0)
                self.comboBox_3.setCurrentIndex(0)
                # log文件
                logFileUrl = '%s/log' % filepath
                self.__class__.createFolder(self, logFileUrl)
                excelFileType = 'xlsx'
                logFileName = 'log'
                logDataPath = self.__class__.getFileName(self, logFileUrl, logFileName, excelFileType)

                # 获取最终ODM数据
                newData = Get_Data()
                newData.getFileData(fileUrl)
                deleteList = {'Amount': 0}
                newData.deleteTheRows(deleteList)
                headList = newData.getHeaderData()
                # 去除Amount with VAT中数值为空的数据，因为数据sales为空
                newData.fileData = newData.fileData[newData.fileData['Amount with VAT'].notnull()]
                newData.fileData = newData.fileData.reset_index(drop=True)

                if ("PHY Material Code" in headList) and ("CHM Material Code" in headList):
                    fillNanColumnKey = {'Material Code': ["PHY Material Code", "CHM Material Code"]}
                    newData.fillNanColumn(fillNanColumnKey)
                getFileDataListKey = ['Project No.', 'CS', 'Sales', 'Currency', 'GPC Glo. Par. Code', 'Material Code',
                                      'SAP No.', 'Amount', 'Amount with VAT', 'Exchange Rate', 'Total Cost']

                combineKeyFieldsList = ['GPC Glo. Par. Code', 'SAP No.', 'Amount', 'Amount with VAT', 'Total Cost']

                if 'Text' in headList:
                    getFileDataListKey.append('Text')
                    combineKeyFieldsList.append('Text')
                if 'Long Text' in headList:
                    getFileDataListKey.append('Long Text')
                    combineKeyFieldsList.append('Long Text')
                # log文件
                combinekeyFields = self.lineEdit_15.text()
                combineKeyFieldsList += combinekeyFields.split(';')
                combineKeyFieldsList.append('Project No.')
                combineKeyFieldsList = excel_field_mapper.update_field_names(combineKeyFieldsList)
                logFile = newData.fileData[combineKeyFieldsList]
                logFile['Order No.'] = ''
                logFile['Remark'] = ''
                logFile['Proforma No.'] = ''
                logFile['sapAmountVat'] = ''
                logFile['Update Time'] = '未开Order'
                if 'Text' not in headList:
                    logFile['Text'] = ''
                if 'Long Text' not in headList:
                    logFile['Long Text'] = ''

                fileDataList = newData.getFileDataList(getFileDataListKey)
                headerData = newData.getHeaderData()
                n = 0
                # 实例化sap
                sap_obj = Sap()
                if sap_obj.res['flag']:
                    for n in range(len(fileDataList['Amount'])):
                        if fileDataList['Material Code'][n] == '':
                            QMessageBox.information(self, "提示信息", "无Material Code，请检查", QMessageBox.Yes)
                            break
                        else:
                            materialCode = fileDataList['Material Code'][n]
                        self.current_material_code = materialCode
                        self.lineEdit_2.setText(str(fileDataList['Project No.'][n]))
                        self.lineEdit_3.setText(str(int(fileDataList['GPC Glo. Par. Code'][n])))
                        self.textBrowser.append("No.:%s" % (n + 1))
                        # if pd.isnull(fileDataList['SAP No.'][n]):
                        # # if math.isnan(fileDataList['SAP No.'][n]):
                        # 	self.textBrowser.append("没有SAP No.")
                        # 	self.lineEdit.setText('')
                        # else:
                        # 	self.lineEdit.setText(str(int(fileDataList['SAP No.'][n])))
                        try:
                            self.lineEdit.setText(str(int(fileDataList['SAP No.'][n])))
                        except:
                            self.lineEdit.setText('')
                        else:
                            pass
                        # materialCodeList = ['', 'T75-441-A2', 'T75-405-A2', 'T20-441-00', 'T20-405-00', 'T75-441-00', 'T75-405-00', 'T75-441-D2', 'T75-405-D2', 'S11-441-10', 'S11-405-10']
                        QApplication.processEvents()
                        # TODO 需要将MC改为订单信息

                        if fileDataList['CS'][n] in configContent:
                            # self.comboBox_2.setCurrentIndex(username.index(fileDataList['CS'][n])+1)
                            self.comboBox_2.setItemText(int(0), fileDataList['CS'][n])
                        else:
                            # self.comboBox_2.setCurrentIndex(0)
                            self.comboBox_2.setItemText(int(0), '')
                        if fileDataList['Sales'][n] in configContent:
                            # self.comboBox_3.setCurrentIndex(username.index(fileDataList['Sales'][n]) + 1)
                            self.comboBox_3.setItemText(int(0), fileDataList['Sales'][n])
                        else:
                            # self.comboBox_3.setCurrentIndex(0)
                            self.comboBox_3.setItemText(int(0), '')
                        self.comboBox.setItemText(int(0), fileDataList['Currency'][n])
                        self.doubleSpinBox_2.setValue(fileDataList['Amount'][n])
                        self.doubleSpinBox_4.setValue(fileDataList['Amount with VAT'][n])
                        self.doubleSpinBox_3.setValue(fileDataList['Total Cost'][n])
                        self.doubleSpinBox.setValue(fileDataList['Exchange Rate'][n])
                        if 'Text' in headList:
                            try:
                                self.lineEdit_5.setText(fileDataList['Text'][n])
                            except:
                                self.lineEdit_5.setText('Testing Fee')
                            else:
                                pass
                        else:
                            self.lineEdit_5.setText('Testing Fee')

                        # TODO 需要将text改为订单信息
                        if 'Long Text' in headList:
                            try:
                                self.current_long_text = fileDataList['Long Text'][n]
                            except:
                                self.current_long_text = ''
                            else:
                                pass
                        else:
                            self.current_long_text = ''
                        QApplication.processEvents()
                        logMsg = self.__class__.sapOperate(self, sap_obj)
                        # 写log
                        logIndex = logFile[(logFile['Project No.'] == fileDataList['Project No.'][n])].index.tolist()[0]
                        logFile.loc[logIndex, 'Order No.'] = logMsg['orderNo']
                        logFile.loc[logIndex, 'Remark'] = logMsg['Remark']
                        logFile.loc[logIndex, 'Proforma No.'] = logMsg['Proforma No.']
                        nowDate = datetime.datetime.today()
                        logFile.loc[logIndex, 'Update Time'] = nowDate.strftime('%Y-%m-%d %H:%M:%S')
                        logDataFile = logFile.to_excel('%s' % logDataPath, merge_cells=False)
                        self.lineEdit_9.setText(logDataPath)
                        if n < len(fileDataList['Amount']) - 1:
                            if self.checkBox_5.isChecked():
                                reply = QMessageBox.question(self, '信息', '是否继续填写下一个Order',
                                                             QMessageBox.Yes | QMessageBox.No,
                                                             QMessageBox.Yes)
                                if reply == QMessageBox.Yes:
                                    continue
                                else:
                                    break
                        else:
                            os.startfile(logFileUrl)
                            os.startfile(logDataPath)
                            self.textBrowser.append("ODM数据已全部填写完成")
                            self.textBrowser.append("log数据:%s" % logDataPath)
                            self.textBrowser.append('----------------------------------')
                            QMessageBox.information(self, "提示信息", "ODM数据已全部填写完成", QMessageBox.Yes)
                    sap_obj.end_sap()
                else:
                    # sap实例化失败
                    self.textBrowser.append("SAP系统未启动")
            else:
                self.textBrowser.append("请重新选择ODM文件")
                QMessageBox.information(self, "提示信息", "请重新选择ODM文件", QMessageBox.Yes)
        except Exception as msg:
            fileData = self.lineEdit_6.text()
            self.textBrowser.append('这份%s的ODM获取数据有问题' % fileData)
            self.textBrowser.append('错误信息：%s' % msg)
            self.textBrowser.append('----------------------------------')
            QMessageBox.information(self, "提示信息", '这份%s的ODM获取数据有问题' % fileData, QMessageBox.Yes)
    def getFileName(self, fileUrl, fileName, fileType):
        nowTime = time.strftime('%Y-%m-%d %H.%M.%S')
        fileName = fileUrl + '/' + nowTime + ' - ' + fileName + '.' + fileType
        return fileName
    def createFolder(self, url):
        isExists = os.path.exists(url)
        if not isExists:
            os.makedirs(url)
    def lineEditChange(self, url):
        combineKey = self.lineEdit_15.text()
        self.lineEdit_16.setText(combineKey)
    def getInvoiceMsg(self):
        try:
            # 特殊开票文件
            global specialInvoiceMsg
            global invoiceName
            global orderMode
            global invoiceRemarks
            global invoiceField
            global invoiceGroup

            # invoiceFile = pd.read_csv(r'%s/%s' % (configFileUrl, configContent['Invoice_File_Name']), encoding="utf8")
            invoiceFile = pd.read_excel(
                '%s/%s' % (configContent['Invoice_File_URL'], configContent['Invoice_File_Name']))
            invoiceName = list(invoiceFile['Invoice name'])
            orderMode = list(invoiceFile['开order方式'])
            invoiceRemarks = list(invoiceFile['开票要求(特殊的外币开票请参考近期的invoice)'])
            invoiceField = list(invoiceFile['字段'])
            invoiceGroup = list(invoiceFile['组别'])
            orderModeList = list(set(orderMode))
            specialInvoiceMsg = {}
            for each in orderModeList:
                specialInvoiceMsg[each] = {}
                # 保留包含关键字的行
                seachInvoiceFile = pd.DataFrame(invoiceFile[invoiceFile['开order方式'] == each])
                # 嵌套字典
                # invoiceName = {}
                # orderMode = {}
                # invoiceRemarks = {}
                # invoiceField = {}
                # invoiceGroup = {}
                embeddedDict = {}
                embeddedDict['Invoice name'] = list(seachInvoiceFile['Invoice name'])
                embeddedDict['Order Mode'] = list(seachInvoiceFile['开order方式'])
                embeddedDict['Invoice Remarks'] = list(seachInvoiceFile['开票要求(特殊的外币开票请参考近期的invoice)'])
                embeddedDict['Invoice Field'] = list(seachInvoiceFile['字段'])
                embeddedDict['Invoice Group'] = list(seachInvoiceFile['组别'])
                specialInvoiceMsg[each] = embeddedDict

            self.textBrowser_2.append('特殊开票信息获取成功')
            self.textBrowser_2.append('特殊开票文件名称：%s/%s' % (configContent['Invoice_File_URL'], configContent['Invoice_File_Name']))
            self.textBrowser_2.append('----------------------------------')
        except Exception as msg:
            self.textBrowser_2.append('错误信息：%s' % msg)
            self.textBrowser_2.append('----------------------------------')
    def splitOdmData(self):
        try:
            fileUrl = self.lineEdit_7.text()
            self.textBrowser_2.append('区分特殊开票的原始文件:%s' % fileUrl)
            (filepath, filename) = os.path.split(fileUrl)
            if fileUrl:
                newData = Get_Data()
                newData.getFileData(fileUrl)
                generalData = newData.fileData[~newData.fileData["Invoices' name (Chinese)"].isin(invoiceName)]
                # 新文件地址
                newFolderUrl = '%s/%s' % (filepath, today)
                self.createFolder(newFolderUrl)
                ExcelFileType = 'xlsx'
                if generalData.empty:
                    pass
                else:
                    invoiceFileName = '1.正常合并'
                    invoiceFilePath = self.getFileName(newFolderUrl, invoiceFileName, ExcelFileType)
                    self.textBrowser_2.append('1:%s' % invoiceFilePath)
                    generalFile = generalData.to_excel('%s' % invoiceFilePath, merge_cells=False)
                specialData = newData.fileData[newData.fileData["Invoices' name (Chinese)"].isin(invoiceName)]
                fileNum = 2
                for each in specialInvoiceMsg:
                    eachSpecialData = specialData[
                        specialData["Invoices' name (Chinese)"].isin(specialInvoiceMsg[each]['Invoice name'])]
                    if eachSpecialData.empty:
                        pass
                    else:
                        invoiceFileName = str(fileNum) + '.' + each
                        invoiceFilePath = self.getFileName(newFolderUrl, invoiceFileName, ExcelFileType)
                        self.textBrowser_2.append('%s:%s' % (fileNum, invoiceFilePath))
                        specialFile = eachSpecialData.to_excel('%s' % invoiceFilePath, merge_cells=False)
                        fileNum += 1
                os.startfile(newFolderUrl)
                self.textBrowser_2.append('处理好特殊数据')
                self.textBrowser_2.append('----------------------------------')
            else:
                self.textBrowser_2.append('请重新选择ODM文件')
                self.textBrowser_2.append('----------------------------------')
        except Exception as msg:
            self.textBrowser_2.append('错误信息：%s' % msg)
            self.textBrowser_2.append('----------------------------------')
    def addColumnMsg(self, data):
        key = self.lineEdit_24.text().split(';')
        key = excel_field_mapper.update_field_names(key)
        df = data
        df['column_msg'] = df[key].apply(lambda row: '\t'.join(map(str, row)), axis=1)
        return df
    def addRowMsg(self, data):
        key = self.lineEdit_23.text().split(';')
        key = excel_field_mapper.update_field_names(key)
        df = data
        df['row_msg'] = df[key].apply(lambda row: '\n'.join(f"{col}:{val}" for col, val in zip(df[key], row)),
                                      axis=1)
        return df
    def combineMsg(self, key, data):
        newData = Get_Data()
        combineProject = data.groupby(key).apply(newData.concat_func).reset_index()
        df = pd.merge(data, combineProject, on=key, how='right')
        return df
    def addMsg(self):
        fileUrl = self.lineEdit_7.text()
        (filepath, filename) = os.path.split(fileUrl)
        if fileUrl:
            rowChecked = self.checkBox_17.isChecked()
            columnChecked = self.checkBox_18.isChecked()
            if rowChecked or columnChecked:
                # try:
                    column_name = self.comboBox_5.currentText()
                    # 转关键词
                    column_key = self.lineEdit_24.text().split(';')
                    # # 老的数据名称
                    # column_key_list = excel_field_mapper.update_field_names(column_key)
                    # column_key_str = ';'.join(column_key_list)
                    # 新的数据名称
                    column_key_str = ';'.join(column_key)
                    key = column_key_str.replace(';', '\t')
                    newData = Get_Data()
                    newData.getFileData(fileUrl)
                    if rowChecked:
                        newData.fileData = self.__class__.addRowMsg(self, newData.fileData)
                    if columnChecked:
                        newData.fileData = self.__class__.addColumnMsg(self, newData.fileData)
                        newData.fileData['column_msg'] = key + '\n' + newData.fileData['column_msg']
                    if rowChecked and columnChecked:
                        newData.fileData[column_name] = newData.fileData['row_msg'] + '\n' + newData.fileData[
                            'column_msg']
                    elif rowChecked:
                        newData.fileData[column_name] = newData.fileData['row_msg']
                    else:
                        newData.fileData[column_name] = newData.fileData['column_msg']
                    if self.checkBox_17.isChecked() or self.checkBox_18.isChecked():
                        newData.fileData[column_name] = newData.fileData[self.comboBox_5.currentText()].str.replace('nan', '')
                        newData.fileData[column_name] = newData.fileData[self.comboBox_5.currentText()].str.replace('XXXXXX', '')
                    ExcelFileType = 'xlsx'
                    odmFileName = '添加信息后的数据'
                    odmDataPath = self.__class__.getFileName(self, filepath, odmFileName, ExcelFileType)
                    # newData.fileData = newData.fileData.applymap(lambda x: x.strip('"') if isinstance(x, str) else x)
                    odmDataFile = newData.fileData.to_excel('%s' % (odmDataPath), merge_cells=False)
                    self.textBrowser_2.append('已完成该文件的信息添加')
                    self.textBrowser_2.append('文件保存在：%s' % odmDataPath)
                    self.textBrowser_2.append('----------------------------------')
                    QApplication.processEvents()
    def odmCombineData(self):
        try:
            fileUrl = self.lineEdit_7.text()
            self.textBrowser_2.append('合并前的原始数据：%s' % fileUrl)
            (filepath, filename) = os.path.split(fileUrl)
            if fileUrl:
                newData = Get_Data()
                newData.getFileData(fileUrl)
                # 删除Amount为0的数据
                deleteRowList = {'Amount': 0}
                newData.deleteTheRows(deleteRowList)
                newData.fileData.sort_values(
                    by=["Invoices' name (Chinese)", 'CS', 'Sales', 'Currency', 'Material Code', 'Buyer(GPC)', 'Month'],
                    axis=0, ascending=[True, True, True, True, True, True, True], inplace=True)
                # 只保留Order No为空的数据
                newData.fileData = newData.fileData[newData.fileData[['SAP Order No.']].isnull().T.any()]
                # material code将空值填上
                headList = newData.getHeaderData()
                if ("PHY Material Code" in headList) and ("CHM Material Code" in headList):
                    fillNanColumnKey = {'Material Code': ["PHY Material Code", "CHM Material Code"]}
                    newData.fillNanColumn(fillNanColumnKey)
                # 将联系人空值填上
                newData.fileData['Client Contact Name'] = newData.fileData['Client Contact Name'].fillna("XXXXXX")
                # 单个数据保留原始数据
                if self.checkBox_17.isChecked():
                    newData.fileData = myWin.addRowMsg(newData.fileData)
                if self.checkBox_18.isChecked():
                    newData.fileData = myWin.addColumnMsg(newData.fileData)
                # 保存原始数据
                fileUrl = '%s/%s' % (filepath, today)
                self.__class__.createFolder(self, fileUrl)
                ExcelFileType = 'xlsx'
                odmFileName = '1.ODM Raw Data'
                odmDataPath = self.__class__.getFileName(self, fileUrl, odmFileName, ExcelFileType)
                odmDataFile = newData.fileData.to_excel('%s' % (odmDataPath), merge_cells=False)
                # 数据透视并保存
                combinekeyFields = self.lineEdit_15.text()
                combineKeyFieldsList = combinekeyFields.split(';')
                # 老的数据名称
                combineKeyFieldsList = excel_field_mapper.update_field_names(combineKeyFieldsList)
                pivotTableKey = combineKeyFieldsList
                # pivotTableKey = ['CS', 'Sales', 'Currency', 'Material Code', "Invoices' name (Chinese)", 'Buyer(GPC)', 'Month', 'Exchange Rate']
                valusKey = ['Amount', 'Amount with VAT', 'Total Cost', 'Revenue\n(RMB)']
                pivotTable = newData.pivotTable(pivotTableKey, valusKey)
                combineFileName = '2.Combine'
                combineFileNamePath = self.__class__.getFileName(self, fileUrl, combineFileName, ExcelFileType)
                combineFile = pivotTable.to_excel('%s' % (combineFileNamePath), merge_cells=False)
                # 读取数据透视数据
                combineData = Get_Data()
                combineData = combineData.getFileData(combineFileNamePath)
                # 删除列
                deleteColumnList = ['Amount', 'Amount with VAT', 'Total Cost', 'Revenue\n(RMB)']
                newData = newData.deleteTheColumn(deleteColumnList)
                # 合并多行数据
                if self.checkBox_17.isChecked():
                    combineRrData = newData.groupby(combineKeyFieldsList).apply(Get_Data().row_concat_func).reset_index()
                    newData = pd.merge(newData, combineRrData, on=combineKeyFieldsList, how='right')
                if self.checkBox_18.isChecked():
                    combineRcData = newData.groupby(combineKeyFieldsList).apply(Get_Data().column_concat_func).reset_index()
                    newData = pd.merge(newData, combineRcData, on=combineKeyFieldsList, how='right')
                    # 转关键词
                    column_key = self.lineEdit_24.text().split(';')
                    # # 老的数据名称
                    # column_key_list = excel_field_mapper.update_field_names(column_key)
                    # column_key_str = ';'.join(column_key_list)
                    # 新字段
                    column_key_str = ';'.join(column_key)
                    key = column_key_str.replace(';', '\t')
                    # key = self.lineEdit_24.text().replace(';', '\t')
                    newData['combine_column_msg'] = key + '\n' + newData['combine_column_msg']
                # 合并两列数据
                if self.checkBox_17.isChecked() and self.checkBox_18.isChecked():
                    newData[self.comboBox_5.currentText()] = newData['combine_row_msg'] + '\n' + newData['combine_column_msg']
                elif self.checkBox_17.isChecked():
                    newData[self.comboBox_5.currentText()] = newData['combine_row_msg']
                elif self.checkBox_18.isChecked():
                    newData[self.comboBox_5.currentText()] = newData['combine_column_msg']
                #  替换无用字符
                if self.checkBox_17.isChecked() or self.checkBox_18.isChecked():
                    newData[self.comboBox_5.currentText()] = newData[self.comboBox_5.currentText()].str.replace('nan', '')
                    newData[self.comboBox_5.currentText()] = newData[self.comboBox_5.currentText()].str.replace('XXXXXX', '')
                # merge数据，combine和原始数据
                onData = combineKeyFieldsList
                # onData = ['CS', 'Sales', 'Currency', 'Material Code', "Invoices' name (Chinese)", 'Buyer(GPC)', 'Month', 'Exchange Rate']
                mergeData = pd.merge(combineData, newData, on=onData, how='right')
                mergeDataName = '3.Merge to Project'
                mergeFileNamePath = self.__class__.getFileName(self, fileUrl, mergeDataName, ExcelFileType)
                mergeFile = mergeData.to_excel('%s' % (mergeFileNamePath), merge_cells=False)
                self.lineEdit_8.setText(mergeFileNamePath)
                # merge数据去重得到最终数据
                mergeData.drop_duplicates(subset=pivotTableKey, keep='first', inplace=True)
                # mergeData['Project No.'] = mergeData['Project No.'].astype(str)
                # mergeData['Project No.'] = mergeData['Project No.'].apply(lambda x: '{:.8f}'.format(float(x)))
                finalDataName = '4.final'
                finalFileNamePath = self.__class__.getFileName(self, fileUrl, finalDataName, ExcelFileType)
                ascendingList = [True] * len(combineKeyFieldsList)
                mergeData.sort_values(by=combineKeyFieldsList, axis=0, ascending=ascendingList, inplace=True)
                # mergeData.sort_values(by=["Invoices' name (Chinese)", 'CS', 'Sales', 'Currency', 'Material Code', 'Buyer(GPC)', 'Month', 'Exchange Rate'], axis=0, ascending=[True, True, True, True, True, True, True, True], inplace=True)
                finalFile = mergeData.to_excel('%s' % (finalFileNamePath), merge_cells=False)
                self.textBrowser_2.append('ODM原始数据：%s' % odmDataPath)
                self.textBrowser_2.append('数据透视数据：%s' % combineFileNamePath)
                self.textBrowser_2.append('添加Project No.的数据：%s' % mergeFileNamePath)
                self.textBrowser_2.append('最终的SAP应用数据：%s' % finalFileNamePath)
                self.lineEdit_6.setText(finalFileNamePath)
                self.textBrowser_2.append('ODM数据已处理完成')
                self.textBrowser_2.append('----------------------------------')
                QApplication.processEvents()
                os.startfile(fileUrl)
                os.startfile(finalFileNamePath)
            else:
                self.textBrowser_2.append('请重新选择ODM文件')
                self.textBrowser_2.append('----------------------------------')
        except Exception as msg:
            fileData = self.lineEdit_7.text()
            self.textBrowser_2.append('这份%s的ODM获取数据有问题' % fileData)
            self.textBrowser_2.append('错误信息：%s' % msg)
            self.textBrowser_2.append('----------------------------------')
            QApplication.processEvents()
    def orderMergeProject(self):
        try:
            combineFileUrl = self.lineEdit_8.text()
            (combineFilepath, combineFilename) = os.path.split(combineFileUrl)
            logFileUrl = self.lineEdit_9.text()
            (logFilepath, logFilename) = os.path.split(logFileUrl)
            if combineFileUrl and logFileUrl:
                ExcelFileType = 'xlsx'
                fileUrl = combineFilepath
                combineFile = Get_Data()
                combineFile.getFileData(combineFileUrl)
                logFile = Get_Data()
                logFile.getFileData(logFileUrl)
                # merge数据，combine和原始数据
                mergekeyFields = self.lineEdit_16.text()
                mergekeyFieldsList = mergekeyFields.split(';')
                mergekeyFieldsList = excel_field_mapper.update_field_names(mergekeyFieldsList)
                # 原来根据多个字段meger
                # combineFile.fileData['SAP No.'] = combineFile.fileData['SAP No.'].apply(int)
                # logFile['SAP No.'] = logFile['SAP No.'].apply(int)
                delColums = ['Text', 'Long Text', 'Proforma No.']
                for col in delColums:
                    if col in combineFile.fileData.columns:
                        combineFile.fileData = combineFile.fileData.drop(columns=col)
                onData = mergekeyFieldsList
                mergeData = pd.merge(combineFile.fileData, logFile.fileData, on=onData, how='outer', indicator=True)
                mergeData.sort_values(by=['Order No.'], axis=0, ascending=[True], inplace=True)
                # 保留数据
                leaveDataList = ["_merge", 'Proforma No.', 'Project No._x', 'Order No.', 'Text', 'Long Text', 'Total Cost_x',
                                 'Revenue\n(RMB)', 'SAP No._x', 'Project No._y', 'Remark', 'Update Time']
                leaveDataList += mergekeyFieldsList
                mergeData = mergeData[leaveDataList]
                ascendingList = [True] * len(leaveDataList)
                mergeData.sort_values(by=leaveDataList, axis=0, ascending=ascendingList, inplace=True)

                mergeDataName = '5.Order Merge Project'
                mergeFileNamePath = self.__class__.getFileName(self, fileUrl, mergeDataName, ExcelFileType)
                mergeFile = mergeData.to_excel('%s' % (mergeFileNamePath), merge_cells=False)
                self.textBrowser_2.append('Order NO 与 Project No合并的数据：%s' % mergeFileNamePath)
                self.textBrowser_2.append(
                    'Order Merge Project 数据,根据Order No数据透视算Amount with VAT的平均数值与ODM导出数据算Amount with VAT总值比较大小，有差说明错误。')
                self.textBrowser_2.append('SAP数据已处理完成')
                self.textBrowser_2.append('----------------------------------')
                os.startfile(combineFileUrl)
                os.startfile(mergeFileNamePath)
                os.startfile(fileUrl)
            else:
                self.textBrowser_2.append('请重新选择文件')
                self.textBrowser_2.append('----------------------------------')
        except Exception as msg:
            self.textBrowser_2.append('Order No Merge Project No数据有问题')
            self.textBrowser_2.append('错误信息：%s' % msg)
            self.textBrowser_2.append('----------------------------------')
            QApplication.processEvents()
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

