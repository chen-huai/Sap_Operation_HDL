import sys, win32com.client, time, datetime, re

from Sap_Operate import *


# from PyQt5 import QtCore, QtGui, QtWidgets
# from PyQt5.QtWidgets import Qself.application, QMainWindow
# from PyQt5.QtWidgets import *
# from PyQt5.QtCore import *

class Sap():
    def __init__(self):
        self.res = {}
        self.res['flag'] = 1
        self.today = time.strftime('%Y.%m.%d')
        self.oneWeekday = (datetime.datetime.now() + datetime.timedelta(days=7)).strftime('%Y.%m.%d')
        res = {}
        res['flag'] = 1
        try:
            self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
            if not type(self.SapGuiAuto) == win32com.client.CDispatch:
                return

            self.application = self.SapGuiAuto.GetScriptingEngine
            if not type(self.application) == win32com.client.CDispatch:
                self.SapGuiAuto = None
                return

            self.connection = self.application.Children(0)
            if not type(self.connection) == win32com.client.CDispatch:
                self.application = None
                self.SapGuiAuto = None
                return

            self.session = self.connection.Children(0)
            if not type(self.session) == win32com.client.CDispatch:
                self.connection = None
                self.application = None
                self.SapGuiAuto = None
                return
        except Exception as msg:
            self.res['flag'] = 0
            self.res['msg'] = ''
            print('SAP未启动')


    # 创建order
    def va01_operate(self, guiData, revenueData):
        res = {}
        res['flag'] = 1
        res['msg'] = ''
        try:
            # 相当于VA01操作
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nva01"
            # 回车键功能
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/ctxtVBAK-AUART").text = guiData['orderType']
            self.session.findById("wnd[0]/usr/ctxtVBAK-VKORG").text = guiData['salesOrganization']
            self.session.findById("wnd[0]/usr/ctxtVBAK-VTWEG").text = guiData['distributionChannels']
            self.session.findById("wnd[0]/usr/ctxtVBAK-VKBUR").text = guiData['salesOffice']
            self.session.findById("wnd[0]/usr/ctxtVBAK-VKGRP").text = guiData['salesGroup']
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById(
                "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR").text = \
                guiData['sapNo']
            self.session.findById(
                "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR").caretPosition = 6
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").text = guiData[
                'projectNo']
            self.session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBKD-BSTDK").text = self.today
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/ctxtVBKD-FBUDA").text = self.today
            self.session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").setFocus()
            self.session.findById(
                "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").caretPosition = 17
            self.session.findById("wnd[0]").sendVKey(0)
            # 售达方按钮
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
            self.session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press()
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBAK-WAERK").text = \
                guiData['currencyType']

            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBAK-WAERK").setFocus()
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBAK-WAERK").caretPosition = 3
            self.session.findById("wnd[0]").sendVKey(0)
            try:
                self.session.findById("wnd[1]").sendVKey(0)
            except:
                pass
            else:
                pass
            if guiData['currencyType'] != "CNY":
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBKD-KURSK").text = \
                    guiData['exchangeRate']
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBKD-KURSK").setFocus()
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBKD-KURSK").caretPosition = 8
                self.session.findById("wnd[0]").sendVKey(0)
            # 会计
            self.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\06").select()
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\06/ssubSUBSCREEN_BODY:SAPMV45A:4311/txtVBAK-XBLNR").text = "*"
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\06/ssubSUBSCREEN_BODY:SAPMV45A:4311/txtVBAK-XBLNR").setFocus()
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\06/ssubSUBSCREEN_BODY:SAPMV45A:4311/txtVBAK-XBLNR").caretPosition = 1
            self.session.findById("wnd[0]").sendVKey(0)
            # 合作伙伴
            self.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09").select()

            # 获取文本名称
            fourName = self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,4]").text
            fiveName = self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,5]").text

            # # eNum负责雇员位置，gNum送达方位置
            if fourName == '负责雇员' or fourName == 'Employee respons.':
                eNum = 4
                gNum = 5
            else:
                eNum = 5
                gNum = 4
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,%s]" % gNum).key = "ZG"
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,%s]" % gNum).text = \
                guiData['globalPartnerCode']
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,%s]" % gNum).setFocus()
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,%s]" % gNum).caretPosition = 8
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,%s]" % eNum).text = \
                guiData['csCode']
            self.session.findById("wnd[0]").sendVKey(0)

            # 联系人
            if guiData['contactCheck']:
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,6]").key = "AP"
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,6]").setFocus()
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,6]").caretPosition = 0
                self.session.findById("wnd[0]").sendVKey(4)
                self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                self.session.findById("wnd[0]").sendVKey(0)

            # 销售
            if guiData['salesName'] != '':
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,7]").key = "VE"
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,7]").text = \
                    guiData['salesCode']
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,7]").setFocus()
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,7]").caretPosition = 4
                self.session.findById("wnd[0]").sendVKey(0)

            # 文本
            self.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10").select()
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text = \
                guiData['shortText']
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").setSelectionIndexes(
                11, 11)
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cmbLV70T-SPRAS").key = "EN"
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cmbLV70T-SPRAS").setFocus()
            self.session.findById("wnd[0]").sendVKey(0)

            # DATA A
            self.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13").select()
            if 'D2' in guiData['materialCode'] or 'D3' in guiData['materialCode']:
                if guiData['sapNo'] in guiData['dataAE1']:
                    self.session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR1").key = "E1"
                else:
                    self.session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR1").key = "Z0"
            elif guiData['sapNo'] in guiData['dataAZ2']:
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR1").key = "Z2"
            else:
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR1").key = "00"

            # DATA B
            self.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14").select()
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/ctxtVBAK-ZZAUART").text = "WO"
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/ctxtVBAK-ZZUNLIMITLIAB").text = "N"
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/ctxtZAUFTD-VORAUS_AUFENDE").text = self.oneWeekday
            if revenueData['revenueForCny'] >= 35000:
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/txtZAUFTD-AUFTRAGSWERT").text = format(
                    revenueData['revenueForCny'], '.2f')
        except Exception as msg:
            res['flag'] = 0
            res['msg'] = 'Order No未创建成功，%s' % msg
            # myWin.textBrowser.append("Order No未创建成功")
        finally:
            return res

    # 填写Data B
    def lab_cost(self, guiData, revenueData):
        res = {}
        res['flag'] = 1
        res['msg'] = ''
        try:
            # revenuedata包含revenue,planCost,revenueForCny,chmCost,phyCost,chmRe,phyRe,chmCsCostAccounting,chmLabCostAccounting,phyCsCostAccounting
            if 'A2' in guiData['materialCode'] or 'D2' in guiData['materialCode'] or 'D3' in guiData['materialCode']:
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AZULEISTENDE/ctxtTABL-KOSTL[0,0]").text = \
                    guiData['chmCostCenter']
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AZULEISTENDE/ctxtTABL-KOSTL[0,1]").text = \
                    guiData['phyCostCenter']
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AKOSTENSAETZE/ctxtTABD-KOSTL[0,0]").text = \
                    guiData['chmCostCenter']
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AKOSTENSAETZE/ctxtTABD-KOSTL[0,1]").text = \
                    guiData['phyCostCenter']
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AKOSTENSAETZE/txtTABD-FESTPREIS[5,0]").text = \
                    revenueData['chmCost']
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AKOSTENSAETZE/txtTABD-FESTPREIS[5,1]").text = \
                    revenueData['phyCost']
            elif 'T20' in guiData['materialCode'] or '430' in guiData['materialCode']:
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AZULEISTENDE/ctxtTABL-KOSTL[0,0]").text = \
                    guiData['phyCostCenter']
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AKOSTENSAETZE/ctxtTABD-KOSTL[0,0]").text = \
                    guiData['phyCostCenter']
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AKOSTENSAETZE/txtTABD-FESTPREIS[5,0]").text = \
                    revenueData['phyCost']
            else:
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AZULEISTENDE/ctxtTABL-KOSTL[0,0]").text = \
                    guiData['chmCostCenter']
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AKOSTENSAETZE/ctxtTABD-KOSTL[0,0]").text = \
                    guiData['chmCostCenter']
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AKOSTENSAETZE/txtTABD-FESTPREIS[5,0]").text = \
                    revenueData['chmCost']
        except Exception as msg:
            res['flag'] = 0
            res['msg'] = 'Data B未填写，%s' % msg
            # myWin.textBrowser.append("Data B未填写")
        finally:
            return res

    # 保存
    def save_sap(self, info):
        res = {}
        res['flag'] = 1
        res['msg'] = ''
        # 保存操作
        try:
            self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        except:
            try:
                self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
            except Exception as msg:
                pass
            else:
                pass
        else:
            try:
                self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
            except Exception as msg:
                pass
            else:
                pass

        finally:
            saveMessageText = self.session.findById("wnd[0]/sbar/pane[0]").text
            if '已保存' in saveMessageText or 'saved' in saveMessageText:
                pass
            else:
                res['flag'] = 0
                res['msg'] += '%s保存失败，%s' % (info, saveMessageText)
            return res
    # 添加item
    def va02_operate(self, guiData, revenueData):
        res = {}
        res['flag'] = 1
        res['orderNo'] = ''
        res['msg'] = ''
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/NVA02"
            self.session.findById("wnd[0]").sendVKey(0)
            orderNo = self.session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text
            res['orderNo'] = orderNo
            self.session.findById("wnd[0]").sendVKey(0)
            if 'A2' in guiData['materialCode']:
                if '405' in guiData['materialCode']:
                    self.session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").text = "T75-405-00"
                    self.session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,1]").text = "T20-405-00"
                elif '430' in guiData['materialCode']:
                    # 顺序上更换
                    self.session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").text = "T20-430-00"
                    self.session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,1]").text = "T75-430-00"
                else:
                    self.session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").text = "T75-441-00"
                    self.session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,1]").text = "T20-441-00"
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/txtVBAP-ZMENG[2,0]").text = "1"
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/txtVBAP-ZMENG[2,1]").text = "1"
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtVBAP-ZIEME[3,0]").text = "pu"
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtVBAP-ZIEME[3,1]").text = "pu"
                self.session.findById("wnd[0]").sendVKey(0)
                # Item的金额填写
                if '430' in guiData['materialCode']:
                    self.session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").setFocus()
                    self.session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").caretPosition = 10
                else:
                    self.session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/txtVBAP-ZMENG[2,1]").setFocus()
                    self.session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/txtVBAP-ZMENG[2,1]").caretPosition = 16
                self.session.findById("wnd[0]").sendVKey(2)
                self.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06").select()
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").text = \
                    revenueData['phyRe']
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").setFocus()
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").caretPosition = 16
                self.session.findById("wnd[0]").sendVKey(0)
                sapAmountVatStr = self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").text
                sapAmountVat = float(sapAmountVatStr.replace(',', ''))

                self.session.findById("wnd[0]/tbar[0]/btn[3]").press()

                if '430' in guiData['materialCode']:
                    self.session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/txtVBAP-ZMENG[2,1]").setFocus()
                    self.session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/txtVBAP-ZMENG[2,1]").caretPosition = 16
                else:
                    self.session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").setFocus()
                    self.session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").caretPosition = 10
                self.session.findById("wnd[0]").sendVKey(2)
                self.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06").select()
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").text = \
                    revenueData['chmRe']
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").setFocus()
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").caretPosition = 16
                self.session.findById("wnd[0]").sendVKey(0)
                sapAmountVatStr = self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").text

                sapAmountVat += float(sapAmountVatStr.replace(',', ''))
                sapAmountVat = format(sapAmountVat, '.2f')
                sapAmountVat = re.sub(r"(\d)(?=(\d\d\d)+(?!\d))", r"\1,", sapAmountVat)

                # 是否需要填写plan cost
                plan_cost_res = Sap.plan_cost(self, guiData, revenueData)
                if not plan_cost_res['flag']:
                    res['msg'] += plan_cost_res['msg']
            else:
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").text = \
                    guiData['materialCode']
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/txtVBAP-ZMENG[2,0]").text = "1"
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtVBAP-ZIEME[3,0]").text = "pu"
                self.session.findById("wnd[0]").sendVKey(0)
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").setFocus()
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").caretPosition = 10
                self.session.findById("wnd[0]").sendVKey(2)
                self.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06").select()
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").text = format(
                    revenueData['revenue'], '.2f')
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").setFocus()
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").caretPosition = 16
                self.session.findById("wnd[0]").sendVKey(0)
                sapAmountVat = self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").text

                # 是否需要填写plan cost
                plan_cost_res = Sap.plan_cost(self, guiData, revenueData)
                if not plan_cost_res['flag']:
                    res['msg'] += plan_cost_res['msg']
            if guiData['longText'] != '':
                # if myWin.checkBox_8.isChecked() or revenueData['revenueForCny'] >= 35000:
                if guiData['planCostCheck'] or revenueData['revenueForCny'] >= 35000:
                    self.session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").setFocus()
                    self.session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").caretPosition = 10
                    self.session.findById("wnd[0]").sendVKey(2)

                self.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09").select()
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text = \
                    guiData['longText']
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").setSelectionIndexes(
                    4, 4)
                try:
                    # 好像国内公司可以成功
                    self.session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cmbLV70T-SPRAS").key = "EN"
                    self.session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cmbLV70T-SPRAS").setFocus()
                    self.session.findById("wnd[0]").sendVKey(0)
                    self.session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").setSelectionIndexes(0, 0)
                except:
                    try:
                        self.session.findById(
                            "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cmbLV70T-SPRAS").key = "EN"
                        self.session.findById(
                            "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cmbLV70T-SPRAS").setFocus()
                        self.session.findById("wnd[0]").sendVKey(0)
                    except:
                        res['msg'] += 'Long Text 添加失败'
                # finally:
                #     # 返回键
                #     self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            if guiData['planCostCheck'] or revenueData['revenueForCny'] >= 35000 or guiData['longText'] == '':
                pass
            else:
                self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            res['sapAmountVat'] = sapAmountVat
        except Exception as msg:
            res['flag'] = 0
            res['msg'] += 'Order添加Item失败，%s' % msg
        finally:
            return res

    # 填写plan cost
    def plan_cost(self, guiData, revenueData):
        res = {}
        res['flag'] = 1
        res['msg'] = ''
        try:
            if guiData['planCostCheck'] or revenueData['revenueForCny'] >= 35000:
                # D2/D3特殊处理：CS两值相加，LAB分两行填写
                if 'D2' in guiData['materialCode'] or 'D3' in guiData['materialCode']:
                    self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                    if revenueData['revenueForCny'] >= 1000:
                        # 定位到material
                        self.session.findById(
                            "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").setFocus()
                        self.session.findById(
                            "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").caretPosition = 10
                        self.session.findById("wnd[0]/mbar/menu[3]/menu[7]").select()
                        self.session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press()
                        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()

                        # CS字段：chmCsCostAccounting + phyCsCostAccounting
                        if guiData['csCheck']:
                            cs_total = round(float(revenueData['chmCsCostAccounting']) + float(revenueData['phyCsCostAccounting']), 0)
                            if cs_total > 0:
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,0]").text = "E"
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,0]").text = guiData['csCostCenter']
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,0]").text = "T01AST"
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").text = cs_total
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").setFocus()
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").caretPosition = 20
                                self.session.findById("wnd[0]").sendVKey(0)

                        # LAB第1行：CHM
                        if guiData['chmCheck']:
                            chm_lab = round(float(revenueData['chmLabCostAccounting']), 0)
                            if chm_lab > 0:
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,1]").text = "E"
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,1]").text = guiData['chmCostCenter']
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,1]").text = "T01AST"
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").text = chm_lab
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").setFocus()
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").caretPosition = 20
                                self.session.findById("wnd[0]").sendVKey(0)

                        # LAB第2行：PHY
                        if guiData['phyCheck']:
                            phy_lab = round(float(revenueData['phyLabCostAccounting']), 0)
                            if phy_lab > 0:
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,2]").text = "E"
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,2]").text = guiData['phyCostCenter']
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,2]").text = "T01AST"
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,2]").text = phy_lab
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,2]").setFocus()
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,2]").caretPosition = 20
                                self.session.findById("wnd[0]").sendVKey(0)

                        # FREMDL外币成本
                        if guiData['cost'] > 0:
                            # 动态计算行号
                            n = 0
                            if guiData['csCheck']:
                                n += 1
                            if guiData['chmCheck']:
                                n += 1
                            if guiData['phyCheck']:
                                n += 1
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,%s]" % n).text = "E"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,%s]" % n).text = guiData['csCostCenter']
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,%s]" % n).text = "FREMDL"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).text = format(guiData['cost'], '.2f')
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).setFocus()
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).caretPosition = 20
                            self.session.findById("wnd[0]").sendVKey(0)

                        # 保存并返回
                        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                        self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                elif 'A2' in guiData['materialCode']:
                    self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                    if revenueData['revenueForCny'] >= 1000:
                        if '430' in guiData['materialCode']:
                            self.session.findById(
                                "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").setFocus()
                            self.session.findById(
                                "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").caretPosition = 10
                        else:
                            # 这个是Item2000的
                            self.session.findById(
                                "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,1]").setFocus()
                            self.session.findById(
                                "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,1]").caretPosition = 10
                        self.session.findById("wnd[0]/mbar/menu[3]/menu[7]").select()
                        self.session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press()
                        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                        # cs
                        if guiData['csCheck'] and round(float(revenueData['phyCsCostAccounting']), 0) > 0:
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,0]").text = "E"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,0]").text = guiData['csCostCenter']
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,0]").text = "T01AST"
                            # 录金额
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").text = round(
                                float(revenueData['phyCsCostAccounting']), 0)
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").setFocus()
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").caretPosition = 20
                            self.session.findById("wnd[0]").sendVKey(0)
                        # phy
                        if guiData['phyCheck'] and round(float(revenueData['phyLabCostAccounting']), 0) > 0:
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,1]").text = "E"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,1]").text = guiData['phyCostCenter']
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,1]").text = "T01AST"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").text = round(
                                float(revenueData['phyLabCostAccounting']), 0)
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").setFocus()
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").caretPosition = 20
                        self.session.findById("wnd[0]").sendVKey(0)

                        # self.session.findById("wnd[0]").sendVKey(0)
                        if '430' in guiData['materialCode']:
                            if guiData['cost'] > 0:
                                if guiData['chmCheck']:
                                    n = 2
                                else:
                                    n = 1
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,%s]" % n).text = "E"
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,%s]" % n).text = guiData[
                                    'csCostCenter']
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,%s]" % n).text = "FREMDL"
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).text = format(
                                    guiData['cost'], '.2f')
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).setFocus()
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).caretPosition = 20
                                self.session.findById("wnd[0]").sendVKey(0)

                        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()

                        # self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
                        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                        self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()

                        # Items1000的plan cost
                        if '430' in guiData['materialCode']:
                            self.session.findById(
                                "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,1]").setFocus()
                            self.session.findById(
                                "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,1]").caretPosition = 10
                        else:
                            self.session.findById(
                                "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").setFocus()
                            self.session.findById(
                            "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").caretPosition = 10
                        self.session.findById("wnd[0]/mbar/menu[3]/menu[7]").select()
                        self.session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press()
                        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                        # cs
                        if guiData['csCheck'] and round(float(revenueData['chmCsCostAccounting']), 0) > 0:
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,0]").text = "E"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,0]").text = guiData['csCostCenter']
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,0]").text = "T01AST"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").text = round(
                                float(revenueData['chmCsCostAccounting']), 0)
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").setFocus()
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").caretPosition = 19
                        # 	chm
                        if guiData['chmCheck'] and round(float(revenueData['chmLabCostAccounting']), 0) > 0:
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,1]").text = "E"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,1]").text = guiData['chmCostCenter']
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,1]").text = "T01AST"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").text = round(
                                float(revenueData['chmLabCostAccounting']), 0)
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").setFocus()
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").caretPosition = 20
                        self.session.findById("wnd[0]").sendVKey(0)
                        #
                        if '430' not in guiData['materialCode']:
                            if guiData['cost'] > 0:
                                if guiData['chmCheck']:
                                    n = 2
                                else:
                                    n = 1
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,%s]" % n).text = "E"
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,%s]" % n).text = guiData[
                                    'csCostCenter']
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,%s]" % n).text = "FREMDL"
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).text = format(
                                    guiData['cost'], '.2f')
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).setFocus()
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).caretPosition = 20
                                self.session.findById("wnd[0]").sendVKey(0)

                        # self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
                        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                        self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                    # self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                else:
                    self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                    if revenueData['revenueForCny'] >= 1000:
                        self.session.findById(
                            "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").setFocus()
                        self.session.findById(
                            "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").caretPosition = 10
                        self.session.findById("wnd[0]/mbar/menu[3]/menu[7]").select()
                        self.session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press()
                        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                        if guiData['csCheck'] and round(float(revenueData['csCostAccounting']), 0) > 0:
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,0]").text = "E"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,0]").text = guiData['csCostCenter']
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,0]").text = "T01AST"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").text = round(
                                float(revenueData['csCostAccounting']), 0)
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").setFocus()
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").caretPosition = 19

                        if (guiData['chmCheck'] or guiData['phyCheck']) and round(float(revenueData['labCostAccounting']), 0) > 0:
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,1]").text = "E"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,1]").text = "T01AST"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").text = round(
                                float(revenueData['labCostAccounting']), 0)
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").setFocus()
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").caretPosition = 20

                        if 'T75' in guiData['materialCode']:
                            if guiData['chmCheck'] and round(float(revenueData['labCostAccounting']), 0) > 0:
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,1]").text = guiData[
                                    'chmCostCenter']
                        else:
                            if guiData['phyCheck'] and round(float(revenueData['labCostAccounting']), 0) > 0:
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,1]").text = guiData[
                                    'phyCostCenter']

                        if guiData['cost'] > 0:
                            if guiData['chmCheck'] or guiData['phyCheck']:
                                n = 2
                            else:
                                n = 1
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,%s]" % n).text = "E"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,%s]" % n).text = guiData[
                                'csCostCenter']
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,%s]" % n).text = "FREMDL"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).text = format(
                                guiData['cost'], '.2f')
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).setFocus()
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).caretPosition = 20
                            self.session.findById("wnd[0]").sendVKey(0)
                        # 直接保存
                        # self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
                        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                        self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        except Exception as msg:
            res['flag'] = 0
            res['msg'] += 'plan cost未添加成功,%s' % msg
        finally:
            return res

    def vf01_operate(self):
        res = {}
        res['flag'] = 1
        res['msg'] = ''
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nvf01"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
        except Exception as msg:
            res['flag'] = 0
            res['msg'] += '形式发票添加失败，%s' % msg
        finally:
            return res

    def vf03_operate(self):
        res = {}
        res['flag'] = 1
        res['msg'] = ''
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nvf03"
            self.session.findById("wnd[0]").sendVKey(0)
            proformaNo = self.session.findById("wnd[0]/usr/ctxtVBRK-VBELN").text
            res['Proforma No.'] = proformaNo
            self.session.findById("wnd[0]/mbar/menu[0]/menu[11]").select()
            self.session.findById("wnd[1]/tbar[0]/btn[37]").press()
        except Exception as msg:
            res['flag'] = 0
            res['msg'] += '形式发票查看失败，%s' % msg
        finally:
            return res

    # 打开order
    def open_va02(self, orderNo):
        res = {}
        res['flag'] = 1
        res['msg'] = ''
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/NVA02"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = orderNo
            self.session.findById("wnd[0]").sendVKey(0)
        except Exception as msg:
            res['flag'] = 0
            res['msg'] = "该Order No %s 未开启，%s" % (orderNo, msg)
            # myWin.textBrowser.append("该Order No %s 未开启" % orderNo)
        finally:
            return res

    # 解锁order
    def unlock_or_lock_order(self, flag):
        res = {}
        res['flag'] = 1
        res['msg'] = ''
        try:
            # 锁order操作
            self.session.findById("wnd[1]").sendVKey(0)
            self.session.findById(
                "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/lblKUAGV-KUNNR").setFocus()
            self.session.findById(
                "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/lblKUAGV-KUNNR").caretPosition = 3
            self.session.findById("wnd[0]").sendVKey(2)
            self.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\12").select()
            self.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\12/ssubSUBSCREEN_BODY:SAPMV45A:4305/btnBT_KSTC").press()
            if flag == 'Unlock':
                self.session.findById(
                    "wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,1]").selected = False
                self.session.findById(
                    "wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,0]").selected = False
            else:
                # self.session.findById(
                #     "wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,1]").selected = True
                self.session.findById(
                    "wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,0]").selected = True
            self.session.findById(
                "wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,0]").setFocus()
            self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            self.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13").select()
            if flag == 'Unlock':
                self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR4").key = "100"
            else:
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR4").key = " "
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR4").setFocus()
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            res['msg'] = "%s 成功" % flag
        except Exception as msg:
            res['flag'] = 0
            res['msg'] = "%s 未成功，%s" % (flag, msg)
        finally:
            return res

    # 结束sap
    def end_sap(self):
        self.session = None
        self.connection = None
        self.application = None
        self.SapGuiAuto = None

    def login_hour_gui(self, hour_data):
        res = {}
        res['flag'] = 1
        res['msg'] = ''
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/NZRU1"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/ctxtZRUCKD-PERNR").text = hour_data['staff_id']
            self.session.findById("wnd[0]/usr/txtZRUCKD-KWEEK").text = hour_data['week']
            self.session.findById("wnd[0]/usr/txtZRUCKD-KWEEK").setFocus()
            self.session.findById("wnd[0]/usr/txtZRUCKD-KWEEK").caretPosition = 2
            self.session.findById("wnd[0]").sendVKey(0)

            time.sleep(1)  # Give SAP time to process
            status_text = self.session.findById("wnd[0]/sbar/pane[0]").text
            if "doesn't exist" in status_text or "does not exist" in status_text or "不存在" in status_text:
                res['flag'] = 0
                res['msg'] = f"登录工时系统失败，员工ID无效: {status_text}"
                # raise Exception(f"登录工时系统失败，员工ID无效: {status_text}")

        except Exception as msg:
            res['flag'] = 0
            res['msg'] = "Hour界面失败，%s" % msg
        return res

    def recording_hours(self, hour_data, row_num=0):
        res = {}
        res['flag'] = 1
        res['msg'] = ''
        try:
            row_num = row_num
            while self.session.findById(
                    f"wnd[0]/usr/tblZIIZRUECKMELD00DYNPRO200/txtZRUCKDS-DATUMK[2,{row_num}]").text != '':
                row_num += 1
            self.session.findById(f"wnd[0]/usr/tblZIIZRUECKMELD00DYNPRO200/txtZRUCKDS-DATUMK[2,{row_num}]").text = \
            hour_data['allocated_day']
            self.session.findById(f"wnd[0]/usr/tblZIIZRUECKMELD00DYNPRO200/ctxtZRUCKDS-BEARBAUFNR[3,{row_num}]").text = \
            hour_data['order_no']
            self.session.findById(f"wnd[0]/usr/tblZIIZRUECKMELD00DYNPRO200/txtZRUCKDS-UEPOS[4,{row_num}]").text = \
            hour_data['item']
            self.session.findById(f"wnd[0]/usr/tblZIIZRUECKMELD00DYNPRO200/ctxtZRUCKDS-ZZTAETIGNR[9,{row_num}]").text = \
            hour_data['material_code']
            self.session.findById(f"wnd[0]/usr/tblZIIZRUECKMELD00DYNPRO200/ctxtZRUCKDS-ZZTAETIGNR[9,{row_num}]").setFocus()
            self.session.findById(f"wnd[0]/usr/tblZIIZRUECKMELD00DYNPRO200/txtZRUCKDS-PZEIT[13,{row_num}]").text = \
            hour_data['allocated_hours']
            self.session.findById(f"wnd[0]/usr/tblZIIZRUECKMELD00DYNPRO200/txtZRUCKDS-BZEIT[15,{row_num}]").text = \
            hour_data['office_time']
            self.session.findById(f"wnd[0]/usr/tblZIIZRUECKMELD00DYNPRO200/txtZRUCKDS-BZEIT[15,{row_num}]").setFocus()
            self.session.findById(
                f"wnd[0]/usr/tblZIIZRUECKMELD00DYNPRO200/txtZRUCKDS-BZEIT[15,{row_num}]").caretPosition = 1
        except Exception as msg:
            res['flag'] = 0
            res['msg'] = "录Hour失败，%s" % msg
        return res

    def save_hours(self):
        res = {}
        res['flag'] = 1
        res['msg'] = ''
        try:
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()

        except Exception as msg:
            max_retries = 14  # 最大重试次数
            retry_count = 0
            last_retry_error = None
            while retry_count < max_retries:
                try:
                    # 回车
                    self.session.findById("wnd[0]").sendVKey(0)
                    # # 勾选弹窗，勾选弹窗最后一步会直接保存
                    # self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                    self.session.findById("wnd[0]").sendVKey(0)
                    saveMessageText = self.session.findById("wnd[0]/sbar/pane[0]").text
                    if 'Fixed price item is allready fully invoiced' in saveMessageText:
                        continue
                    elif 'Data was saved' in saveMessageText:
                        res['msg'] = '录Hour成功'
                        break
                    else:
                        self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
                        self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                        break
                except Exception as retry_error:
                    last_retry_error = retry_error
                    retry_count += 1
                    if retry_count >= max_retries:
                        res['flag'] = 0
                        res['msg'] = f"保存失败，已重试{max_retries}次。初始错误: {msg}. 最后一次重试错误: {last_retry_error}"
                        # raise Exception(f"保存失败，已重试{max_retries}次。初始错误: {msg}. 最后一次重试错误: {last_retry_error}")
                    continue
        return res

# if __name__ == "__main__":
#     revenue = 230
#     guiData = {}
#     guiData['sapNo'] = 5010920197
#     guiData['projectNo'] = '66.405.23.7556.02A'
#     guiData['materialCode'] = 'T75-405-A2'
#     guiData['currencyType'] = 'CNY'
#     guiData['exchangeRate'] = float(1)
#     guiData['globalPartnerCode'] = 1500155
#     guiData['csName'] = 'cai, barry'
#     guiData['salesName'] = ''
#     guiData['amount'] = float(200)
#     guiData['cost'] = float(0)
#     guiData['amountVat'] = float(212)
#     guiData['csHourlyRate'] = float(300)
#     guiData['chmHourlyRate'] = float(250)
#     guiData['phyHourlyRate'] = float(280)
#     guiData['longText'] = ''
#     guiData['shortText'] = 'TEST'
#     guiData['planCostRate'] = float(0.9)
#     guiData['significantDigits'] = int(0)
#     guiData['chmCostRate'] = float(0.3)
#     guiData['phyCostRate'] = float(0.3)
#     guiData['dataAE1'] = ''
#     guiData['dataAZ2'] = ''
#     guiData['orderType'] = 'DR'
#     guiData['salesOrganization'] = '0486'
#     guiData['distributionChannels'] = '01'
#     guiData['salesOffice'] = '>601'
#     guiData['salesGroup'] = '240'
#     guiData['csCostCenter'] = '48601240'
#     guiData['chmCostCenter'] = '48601293'
#     guiData['phyCostCenter'] = '48601294'
#     guiData['csCode'] = '6375108'
#     guiData['salesCode'] = ''
#     my_w = MyMainWindow()
#     revenueData = my_w.getRevenueData(guiData)
#     sap_obj = Sap()
#     if sap_obj.flag != 0:
#         sap_obj.va01_operate(guiData, revenueData)
