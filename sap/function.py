"""
SAP GUI 自动化操作

使用方法:
    from sap import Sap, SapConfig, OperationFlags, OrderData, RevenueData, HourData, SapResult

    config = SapConfig(order_type='ZOR', sales_organization='3002', ...)
    flags = OperationFlags(va01=True, va02=True, ...)
    sap_obj = Sap(config=config, flags=flags)

    order = OrderData(sap_no='123456', project_no='PRJ-001', ...)
    revenue = RevenueData(revenue=10000.0, revenue_cny=72500.0, ...)

    result = sap_obj.va01_operate(order, revenue)
    if result.success:
        print('VA01 创建成功')
"""

import sys
import time
import re

import win32com.client

from sap.models import SapConfig, OrderData, RevenueData, OperationFlags, HourData, SapResult


class Sap:
    def __init__(self, config: SapConfig, flags: OperationFlags):
        self.config = config
        self.flags = flags
        self.result = SapResult()
        self.today = time.strftime('%Y.%m.%d')

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
        except Exception as e:
            self.result = SapResult.fail(str(e))
            print('SAP未启动')

    def _get_a2_materials(self, material_code: str) -> tuple[str, str]:
        """根据物料代码查找 A2 物料映射, 返回 (Item1000物料号, Item2000物料号)"""
        for sub_code, materials in self.config.a2_material_mapping.items():
            if sub_code in material_code:
                return materials
        # 默认使用 441 映射
        return self.config.a2_material_mapping.get('441', ('T75-441-00', 'T20-441-00'))

    # 创建order
    def va01_operate(self, order: OrderData, revenue: RevenueData) -> SapResult:
        result = SapResult()
        try:
            # 相当于VA01操作
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nva01"
            # 回车键功能
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/ctxtVBAK-AUART").text = self.config.order_type
            self.session.findById("wnd[0]/usr/ctxtVBAK-VKORG").text = self.config.sales_organization
            self.session.findById("wnd[0]/usr/ctxtVBAK-VTWEG").text = self.config.distribution_channels
            self.session.findById("wnd[0]/usr/ctxtVBAK-VKBUR").text = self.config.sales_office
            self.session.findById("wnd[0]/usr/ctxtVBAK-VKGRP").text = self.config.cost_center
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById(
                "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR").text = \
                order.sap_no
            self.session.findById(
                "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR").caretPosition = 6
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").text = order.project_no
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
                order.currency_type

            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBAK-WAERK").setFocus()
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBAK-WAERK").caretPosition = 3
            self.session.findById("wnd[0]").sendVKey(0)
            try:
                self.session.findById("wnd[1]").sendVKey(0)
            except:
                pass
            if order.currency_type != "CNY":
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBKD-KURSK").text = \
                    order.exchange_rate
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

            # eNum负责雇员位置，gNum送达方位置
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
                order.global_partner_code
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,%s]" % gNum).setFocus()
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,%s]" % gNum).caretPosition = 8
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,%s]" % eNum).text = \
                self.config.cs_code
            self.session.findById("wnd[0]").sendVKey(0)

            # 联系人
            if self.flags.contact:
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
            if order.sales_name != '':
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,7]").key = "VE"
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,7]").text = \
                    self.config.sales_code
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,7]").setFocus()
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/ctxtGVS_TC_DATA-REC-PARTNER[1,7]").caretPosition = 4
                self.session.findById("wnd[0]").sendVKey(0)

            # 文本
            self.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10").select()
            self.session.findById(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text = \
                order.short_text
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
            if 'D2' in order.material_code or 'D3' in order.material_code:
                if order.sap_no in self.config.data_ae1:
                    self.session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR1").key = "E1"
                else:
                    self.session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR1").key = "Z0"
            elif order.sap_no in self.config.data_az2:
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
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/ctxtZAUFTD-VORAUS_AUFENDE").text = order.ecd
            if revenue.revenue_cny >= self.config.revenue_threshold:
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/txtZAUFTD-AUFTRAGSWERT").text = format(
                    revenue.revenue_cny, '.2f')
        except Exception as e:
            return SapResult.fail(f'Order No未创建成功，{e}')
        return result

    # 填写Data B
    def lab_cost(self, order: OrderData, revenue: RevenueData) -> SapResult:
        result = SapResult()
        try:
            if 'A2' in order.material_code or 'D2' in order.material_code or 'D3' in order.material_code:
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AZULEISTENDE/ctxtTABL-KOSTL[0,0]").text = \
                    self.config.sub_cost_center_chm
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AZULEISTENDE/ctxtTABL-KOSTL[0,1]").text = \
                    self.config.sub_cost_center_phy
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AKOSTENSAETZE/ctxtTABD-KOSTL[0,0]").text = \
                    self.config.sub_cost_center_chm
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AKOSTENSAETZE/ctxtTABD-KOSTL[0,1]").text = \
                    self.config.sub_cost_center_phy
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AKOSTENSAETZE/txtTABD-FESTPREIS[5,0]").text = \
                    revenue.chm_cost
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AKOSTENSAETZE/txtTABD-FESTPREIS[5,1]").text = \
                    revenue.phy_cost
            elif 'T20' in order.material_code or '430' in order.material_code:
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AZULEISTENDE/ctxtTABL-KOSTL[0,0]").text = \
                    self.config.sub_cost_center_phy
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AKOSTENSAETZE/ctxtTABD-KOSTL[0,0]").text = \
                    self.config.sub_cost_center_phy
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AKOSTENSAETZE/txtTABD-FESTPREIS[5,0]").text = \
                    revenue.phy_cost
            else:
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AZULEISTENDE/ctxtTABL-KOSTL[0,0]").text = \
                    self.config.sub_cost_center_chm
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AKOSTENSAETZE/ctxtTABD-KOSTL[0,0]").text = \
                    self.config.sub_cost_center_chm
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/tblSAPMV45AKOSTENSAETZE/txtTABD-FESTPREIS[5,0]").text = \
                    revenue.chm_cost
        except Exception as e:
            return SapResult.fail(f'Data B未填写，{e}')
        return result

    # 保存
    def save_sap(self, info: str) -> SapResult:
        result = SapResult()
        save_error: Exception | None = None

        # 保存操作 — 多次尝试点击确认按钮
        try:
            self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        except Exception as e1:
            save_error = e1
            try:
                self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                save_error = None  # 重试成功，清除错误
            except Exception as e2:
                save_error = e2
        else:
            try:
                self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
            except:
                pass

        # 检查状态栏确认保存结果
        try:
            save_msg = self.session.findById("wnd[0]/sbar/pane[0]").text
        except Exception as e_bar:
            # 状态栏也读不到，带上之前的保存异常
            err = f'{info}保存失败，无法读取状态栏: {e_bar}'
            if save_error:
                err += f'；保存操作异常: {save_error}'
            return SapResult.fail(err)

        if '已保存' not in save_msg and 'saved' not in save_msg:
            err = f'{info}保存失败，{save_msg}'
            if save_error:
                err += f'；保存操作异常: {save_error}'
            return SapResult.fail(err)

        return result

    # 添加item
    def va02_operate(self, order: OrderData, revenue: RevenueData) -> SapResult:
        result = SapResult()
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/NVA02"
            self.session.findById("wnd[0]").sendVKey(0)
            orderNo = self.session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text
            result.order_no = orderNo
            self.session.findById("wnd[0]").sendVKey(0)
            if 'A2' in order.material_code:
                item1000, item2000 = self._get_a2_materials(order.material_code)
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").text = item1000
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,1]").text = item2000
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
                if '430' in order.material_code:
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
                    revenue.phy_revenue
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").setFocus()
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").caretPosition = 16
                self.session.findById("wnd[0]").sendVKey(0)
                sapAmountVatStr = self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").text
                sapAmountVat = float(sapAmountVatStr.replace(',', ''))

                self.session.findById("wnd[0]/tbar[0]/btn[3]").press()

                if '430' in order.material_code:
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
                    revenue.chm_revenue
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
                plan_cost_res = self.plan_cost(order, revenue)
                if not plan_cost_res.success:
                    result.append_message(plan_cost_res.message)
            else:
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").text = \
                    order.material_code
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
                    revenue.revenue, '.2f')
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").setFocus()
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").caretPosition = 16
                self.session.findById("wnd[0]").sendVKey(0)
                sapAmountVat = self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]").text

                # 是否需要填写plan cost
                plan_cost_res = self.plan_cost(order, revenue)
                if not plan_cost_res.success:
                    result.append_message(plan_cost_res.message)
            if order.long_text != '':
                if self.flags.plan_cost or revenue.revenue_cny >= self.config.revenue_threshold:
                    self.session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").setFocus()
                    self.session.findById(
                        "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").caretPosition = 10
                    self.session.findById("wnd[0]").sendVKey(2)

                self.session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09").select()
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text = \
                    order.long_text
                self.session.findById(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").setSelectionIndexes(
                    4, 4)
                try:
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
                        result.append_message('Long Text 添加失败')
            if self.flags.plan_cost or revenue.revenue_cny >= self.config.revenue_threshold or order.long_text == '':
                pass
            else:
                self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            result.sap_amount_vat = sapAmountVat
        except Exception as e:
            result.success = False
            result.append_message(f'Order添加Item失败，{e}')
        return result

    # 填写plan cost
    def plan_cost(self, order: OrderData, revenue: RevenueData) -> SapResult:
        result = SapResult()
        try:
            if self.flags.plan_cost or revenue.revenue_cny >= self.config.revenue_threshold:
                # D2/D3特殊处理：CS两值相加，LAB分两行填写
                if 'D2' in order.material_code or 'D3' in order.material_code:
                    self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                    if revenue.revenue_cny >= self.config.plan_cost_min_threshold:
                        # 定位到material
                        self.session.findById(
                            "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").setFocus()
                        self.session.findById(
                            "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").caretPosition = 10
                        self.session.findById("wnd[0]/mbar/menu[3]/menu[7]").select()
                        self.session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press()
                        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()

                        # CS字段：chm_cs_cost + phy_cs_cost
                        if self.flags.cs:
                            cs_total = round(float(revenue.chm_cs_cost) + float(revenue.phy_cs_cost), 0)
                            if cs_total > 0:
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,0]").text = "E"
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,0]").text = self.config.sub_cost_center_cs
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
                        if self.flags.chm:
                            chm_lab = round(float(revenue.chm_lab_cost), 0)
                            if chm_lab > 0:
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,1]").text = "E"
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,1]").text = self.config.sub_cost_center_chm
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
                        if self.flags.phy:
                            phy_lab = round(float(revenue.phy_lab_cost), 0)
                            if phy_lab > 0:
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,2]").text = "E"
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,2]").text = self.config.sub_cost_center_phy
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
                        if order.cost > 0:
                            # 动态计算行号
                            n = 0
                            if self.flags.cs:
                                n += 1
                            if self.flags.chm:
                                n += 1
                            if self.flags.phy:
                                n += 1
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,%s]" % n).text = "E"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,%s]" % n).text = self.config.sub_cost_center_cs
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,%s]" % n).text = "FREMDL"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).text = format(order.cost, '.2f')
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).setFocus()
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).caretPosition = 20
                            self.session.findById("wnd[0]").sendVKey(0)

                        # 保存并返回
                        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                        self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                elif 'A2' in order.material_code:
                    self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                    if revenue.revenue_cny >= self.config.plan_cost_min_threshold:
                        if '430' in order.material_code:
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
                        if self.flags.cs and round(float(revenue.phy_cs_cost), 0) > 0:
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,0]").text = "E"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,0]").text = self.config.sub_cost_center_cs
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,0]").text = "T01AST"
                            # 录金额
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").text = round(
                                float(revenue.phy_cs_cost), 0)
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").setFocus()
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").caretPosition = 20
                            self.session.findById("wnd[0]").sendVKey(0)
                        # phy
                        if self.flags.phy and round(float(revenue.phy_lab_cost), 0) > 0:
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,1]").text = "E"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,1]").text = self.config.sub_cost_center_phy
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,1]").text = "T01AST"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").text = round(
                                float(revenue.phy_lab_cost), 0)
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").setFocus()
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").caretPosition = 20
                        self.session.findById("wnd[0]").sendVKey(0)

                        if '430' in order.material_code:
                            if order.cost > 0:
                                if self.flags.chm:
                                    n = 2
                                else:
                                    n = 1
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,%s]" % n).text = "E"
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,%s]" % n).text = self.config.sub_cost_center_cs
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,%s]" % n).text = "FREMDL"
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).text = format(
                                    order.cost, '.2f')
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).setFocus()
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).caretPosition = 20
                                self.session.findById("wnd[0]").sendVKey(0)

                        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()

                        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                        self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()

                        # Items1000的plan cost
                        if '430' in order.material_code:
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
                        if self.flags.cs and round(float(revenue.chm_cs_cost), 0) > 0:
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,0]").text = "E"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,0]").text = self.config.sub_cost_center_cs
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,0]").text = "T01AST"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").text = round(
                                float(revenue.chm_cs_cost), 0)
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").setFocus()
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").caretPosition = 19
                        # 	chm
                        if self.flags.chm and round(float(revenue.chm_lab_cost), 0) > 0:
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,1]").text = "E"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,1]").text = self.config.sub_cost_center_chm
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,1]").text = "T01AST"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").text = round(
                                float(revenue.chm_lab_cost), 0)
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").setFocus()
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").caretPosition = 20
                        self.session.findById("wnd[0]").sendVKey(0)
                        #
                        if '430' not in order.material_code:
                            if order.cost > 0:
                                if self.flags.chm:
                                    n = 2
                                else:
                                    n = 1
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,%s]" % n).text = "E"
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,%s]" % n).text = self.config.sub_cost_center_cs
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,%s]" % n).text = "FREMDL"
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).text = format(
                                    order.cost, '.2f')
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).setFocus()
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).caretPosition = 20
                                self.session.findById("wnd[0]").sendVKey(0)

                        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                        self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                else:
                    self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                    if revenue.revenue_cny >= self.config.plan_cost_min_threshold:
                        self.session.findById(
                            "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").setFocus()
                        self.session.findById(
                            "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,0]").caretPosition = 10
                        self.session.findById("wnd[0]/mbar/menu[3]/menu[7]").select()
                        self.session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press()
                        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                        if self.flags.cs and round(float(revenue.cs_cost), 0) > 0:
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,0]").text = "E"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,0]").text = self.config.sub_cost_center_cs
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,0]").text = "T01AST"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").text = round(
                                float(revenue.cs_cost), 0)
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").setFocus()
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,0]").caretPosition = 19

                        if (self.flags.chm or self.flags.phy) and round(float(revenue.lab_cost), 0) > 0:
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,1]").text = "E"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,1]").text = "T01AST"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").text = round(
                                float(revenue.lab_cost), 0)
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").setFocus()
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,1]").caretPosition = 20

                        if 'T75' in order.material_code:
                            if self.flags.chm and round(float(revenue.lab_cost), 0) > 0:
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,1]").text = self.config.sub_cost_center_chm
                        else:
                            if self.flags.phy and round(float(revenue.lab_cost), 0) > 0:
                                self.session.findById(
                                    "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,1]").text = self.config.sub_cost_center_phy

                        if order.cost > 0:
                            if self.flags.chm or self.flags.phy:
                                n = 2
                            else:
                                n = 1
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,%s]" % n).text = "E"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,%s]" % n).text = self.config.sub_cost_center_cs
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,%s]" % n).text = "FREMDL"
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).text = format(
                                order.cost, '.2f')
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).setFocus()
                            self.session.findById(
                                "wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,%s]" % n).caretPosition = 20
                            self.session.findById("wnd[0]").sendVKey(0)
                        # 直接保存
                        self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
                        self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        except Exception as e:
            result.success = False
            result.append_message(f'plan cost未添加成功,{e}')
        return result

    def vf01_operate(self) -> SapResult:
        result = SapResult()
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nvf01"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
        except Exception as e:
            return SapResult.fail(f'形式发票添加失败，{e}')
        return result

    def vf03_operate(self) -> SapResult:
        result = SapResult()
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nvf03"
            self.session.findById("wnd[0]").sendVKey(0)
            proformaNo = self.session.findById("wnd[0]/usr/ctxtVBRK-VBELN").text
            result.proforma_no = proformaNo
            self.session.findById("wnd[0]/mbar/menu[0]/menu[11]").select()
            self.session.findById("wnd[1]/tbar[0]/btn[37]").press()
        except Exception as e:
            return SapResult.fail(f'形式发票查看失败，{e}')
        return result

    # 打开order
    def open_va02(self, orderNo: str) -> SapResult:
        result = SapResult()
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/NVA02"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = orderNo
            self.session.findById("wnd[0]").sendVKey(0)
        except Exception as e:
            return SapResult.fail(f'该Order No {orderNo} 未开启，{e}')
        return result

    # 解锁order
    def unlock_or_lock_order(self, flag: str) -> SapResult:
        result = SapResult()
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
            result.message = f"{flag} 成功"
        except Exception as e:
            return SapResult.fail(f'{flag} 未成功，{e}')
        return result

    # 结束sap
    def end_sap(self):
        self.session = None
        self.connection = None
        self.application = None
        self.SapGuiAuto = None

    def login_hour_gui(self, hour: HourData) -> SapResult:
        result = SapResult()
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/NZRU1"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/ctxtZRUCKD-PERNR").text = hour.staff_id
            self.session.findById("wnd[0]/usr/txtZRUCKD-KWEEK").text = hour.week
            self.session.findById("wnd[0]/usr/txtZRUCKD-KWEEK").setFocus()
            self.session.findById("wnd[0]/usr/txtZRUCKD-KWEEK").caretPosition = 2
            self.session.findById("wnd[0]").sendVKey(0)

            time.sleep(1)  # Give SAP time to process
            status_text = self.session.findById("wnd[0]/sbar/pane[0]").text
            if "doesn't exist" in status_text or "does not exist" in status_text or "不存在" in status_text:
                return SapResult.fail(f'登录工时系统失败，员工ID无效: {status_text}')

        except Exception as e:
            return SapResult.fail(f'Hour界面失败，{e}')
        return result

    def recording_hours(self, hour: HourData, row_num: int = 0) -> SapResult:
        result = SapResult()
        try:
            while self.session.findById(
                    f"wnd[0]/usr/tblZIIZRUECKMELD00DYNPRO200/txtZRUCKDS-DATUMK[2,{row_num}]").text != '':
                row_num += 1
            self.session.findById(f"wnd[0]/usr/tblZIIZRUECKMELD00DYNPRO200/txtZRUCKDS-DATUMK[2,{row_num}]").text = \
            hour.allocated_day
            self.session.findById(f"wnd[0]/usr/tblZIIZRUECKMELD00DYNPRO200/ctxtZRUCKDS-BEARBAUFNR[3,{row_num}]").text = \
            hour.order_no
            self.session.findById(f"wnd[0]/usr/tblZIIZRUECKMELD00DYNPRO200/txtZRUCKDS-UEPOS[4,{row_num}]").text = \
            hour.item
            self.session.findById(f"wnd[0]/usr/tblZIIZRUECKMELD00DYNPRO200/ctxtZRUCKDS-ZZTAETIGNR[9,{row_num}]").text = \
            hour.material_code
            self.session.findById(f"wnd[0]/usr/tblZIIZRUECKMELD00DYNPRO200/ctxtZRUCKDS-ZZTAETIGNR[9,{row_num}]").setFocus()
            self.session.findById(f"wnd[0]/usr/tblZIIZRUECKMELD00DYNPRO200/txtZRUCKDS-PZEIT[13,{row_num}]").text = \
            hour.allocated_hours
            self.session.findById(f"wnd[0]/usr/tblZIIZRUECKMELD00DYNPRO200/txtZRUCKDS-BZEIT[15,{row_num}]").text = \
            hour.office_time
            self.session.findById(f"wnd[0]/usr/tblZIIZRUECKMELD00DYNPRO200/txtZRUCKDS-BZEIT[15,{row_num}]").setFocus()
            self.session.findById(
                f"wnd[0]/usr/tblZIIZRUECKMELD00DYNPRO200/txtZRUCKDS-BZEIT[15,{row_num}]").caretPosition = 1
        except Exception as e:
            return SapResult.fail(f'录Hour失败，{e}')
        return result

    def save_hours(self) -> SapResult:
        result = SapResult()
        try:
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()

        except Exception as e:
            max_retries = 14
            retry_count = 0
            last_retry_error = None
            while retry_count < max_retries:
                try:
                    # 回车
                    self.session.findById("wnd[0]").sendVKey(0)
                    self.session.findById("wnd[0]").sendVKey(0)
                    saveMessageText = self.session.findById("wnd[0]/sbar/pane[0]").text
                    if 'Fixed price item is allready fully invoiced' in saveMessageText:
                        continue
                    elif 'Data was saved' in saveMessageText:
                        result.message = '录Hour成功'
                        break
                    else:
                        self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
                        self.session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                        break
                except Exception as retry_error:
                    last_retry_error = retry_error
                    retry_count += 1
                    if retry_count >= max_retries:
                        return SapResult.fail(
                            f'保存失败，已重试{max_retries}次。初始错误: {e}. 最后一次重试错误: {last_retry_error}'
                        )
                    continue
        return result
