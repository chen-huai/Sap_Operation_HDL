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
from sap import CostOptions, OrderData, OrderItemData, OrderService, PartnerOptions, RevenueData, SapConfig, SapSession
from runtime_globals import configContent

class SapOrderMixin:
    def getRevenueData(self, guiData):
        # 计算金额
        # revenue,planCost,revenueForCny,chmCost,phyCost,chmRe,phyRe,chmCsCostAccounting,chmLabCostAccounting,phyCsCostAccounting
        revenueData = {}
        revenueData['revenue'] = guiData['amountVat'] / 1.06
        # plan cost
        # planCost = revenueData['revenue'] * guiData['exchangeRate'] * 0.9 - guiData['cost']
        revenueData['planCost'] = revenueData['revenue'] * guiData['exchangeRate']
        revenueData['revenueForCny'] = revenueData['revenue'] * guiData['exchangeRate']

        # 405-A2的计算公式
        if ('405' in guiData['materialCode']) and (
                ("A2" in guiData['materialCode']) or ("D2" in guiData['materialCode']) or (
                "D3" in guiData['materialCode'])):
            # DataB-CHM成本
            revenueData['chmCost'] = format(
                (revenueData['revenueForCny'] - guiData['cost']) * guiData['chmCostRate'] * 0.5, '.2f')
            # DataB-PHY成本
            revenueData['phyCost'] = format(
                (revenueData['revenueForCny'] - guiData['cost']) * guiData['phyCostRate'] * 0.5, '.2f')
            # Item1000 的revenue
            revenueData['chmRe'] = format(revenueData['revenue'] * 0.5, '.2f')
            # Item2000 的revenue
            revenueData['phyRe'] = format(revenueData['revenue'] * 0.5, '.2f')
            # plan cost总算法
            # revenueData['chmCsCostAccounting'] = format(revenueData['planCost'] * 0.5 * (1 - 0.3  - (1 - guiData['planCostRate'] )) / guiData['csHourlyRate'], '.%sf' % guiData['significantDigits'])
            # revenueData['chmLabCostAccounting'] = format(revenueData['planCost'] * 0.5 * 0.3 / guiData['chmHourlyRate'], '.%sf' % guiData['significantDigits'])
            # revenueData['phyCsCostAccounting'] = format(revenueData['planCost'] * 0.5 * (1 - 0.3  - (1 - guiData['planCostRate'] )) / guiData['csHourlyRate'], '.%sf' % guiData['significantDigits'])
            # revenueData['phyLabCostAccounting'] = format(revenueData['planCost'] * 0.5 * 0.3 / guiData['chmHourlyRate'], '.%sf' % guiData['significantDigits'])

            # plan cost，理论上（revenue-total cost）*0.9*0.5，实际上SFL省略了0.9的计算（金额不大）

            # CS的Item1000-Cost
            revenueData['chmCsCostAccounting'] = format(
                (revenueData['revenueForCny'] - guiData['cost']) * guiData['planCostRate'] * 0.5 * (
                        1 - guiData['chmCostRate']) / guiData['csHourlyRate'],
                '.%sf' % guiData['significantDigits'])
            # CHM的Item1000-Cost
            revenueData['chmLabCostAccounting'] = format(
                (revenueData['revenueForCny'] - guiData['cost']) * guiData['planCostRate'] * 0.5 * guiData[
                    'chmCostRate'] /
                guiData['chmHourlyRate'], '.%sf' % guiData['significantDigits'])
            # CS的Item2000-Cost
            revenueData['phyCsCostAccounting'] = format(
                (revenueData['revenueForCny'] - guiData['cost']) * guiData['planCostRate'] * 0.5 * (
                        1 - guiData['phyCostRate']) / guiData['csHourlyRate'],
                '.%sf' % guiData['significantDigits'])
            # PHY的Item2000-Cost
            revenueData['phyLabCostAccounting'] = format(
                (revenueData['revenueForCny'] - guiData['cost']) * guiData['planCostRate'] * 0.5 * guiData[
                    'phyCostRate'] /
                guiData['phyHourlyRate'], '.%sf' % guiData['significantDigits'])

        # 441-A2计算公式
        elif ('441' in guiData['materialCode']) and ((
                "A2" in guiData['materialCode'] or ("D2" in guiData['materialCode']) or (
                "D3" in guiData['materialCode']))):
            # DataB-CHM成本
            revenueData['chmCost'] = format(
                (revenueData['revenueForCny'] - guiData['cost']) * guiData['chmCostRate'] * 0.8, '.2f')
            # DataB-PHY成本
            revenueData['phyCost'] = format(
                (revenueData['revenueForCny'] - guiData['cost']) * guiData['phyCostRate'] * 0.2, '.2f')
            # Item1000 的revenue
            revenueData['chmRe'] = format(revenueData['revenue'] * 0.8, '.2f')
            # Item2000 的revenue
            revenueData['phyRe'] = format(revenueData['revenue'] * 0.2, '.2f')
            # plan cost总算法
            # revenueData['chmCsCostAccounting'] = format(revenueData['planCost'] * 0.8 * (1 - 0.3  - (1 - guiData['planCostRate'] )) / guiData['csHourlyRate'], '.%sf' % guiData['significantDigits'])
            # revenueData['chmLabCostAccounting'] = format(revenueData['planCost'] * 0.8 * 0.3 / guiData['chmHourlyRate'], '.%sf' % guiData['significantDigits'])
            # revenueData['phyCsCostAccounting'] = format(revenueData['planCost'] * 0.2 * (1 - 0.3  - (1 - guiData['planCostRate'] )) / guiData['csHourlyRate'], '.%sf' % guiData['significantDigits'])
            # revenueData['phyLabCostAccounting'] = format(revenueData['planCost'] * 0.2 * 0.3 / guiData['chmHourlyRate'], '.%sf' % guiData['significantDigits'])

            # CS的Item1000-Cost
            revenueData['chmCsCostAccounting'] = format(
                (revenueData['revenueForCny'] - guiData['cost']) * guiData['planCostRate'] * 0.8 * (
                        1 - guiData['chmCostRate']) / guiData['csHourlyRate'],
                '.%sf' % guiData['significantDigits'])
            # CHM的Item1000-Cost
            revenueData['chmLabCostAccounting'] = format(
                (revenueData['revenueForCny'] - guiData['cost']) * guiData['planCostRate'] * 0.8 * guiData[
                    'chmCostRate'] /
                guiData['chmHourlyRate'], '.%sf' % guiData['significantDigits'])
            # CS的Item2000-Cost
            revenueData['phyCsCostAccounting'] = format(
                (revenueData['revenueForCny'] - guiData['cost']) * guiData['planCostRate'] * 0.2 * (
                        1 - guiData['phyCostRate']) / guiData['csHourlyRate'],
                '.%sf' % guiData['significantDigits'])
            # PHY的Item2000-Cost
            revenueData['phyLabCostAccounting'] = format(
                (revenueData['revenueForCny'] - guiData['cost']) * guiData['planCostRate'] * 0.2 * guiData[
                    'phyCostRate'] /
                guiData['phyHourlyRate'], '.%sf' % guiData['significantDigits'])

        # 430-A2计算公式
        elif ('430' in guiData['materialCode']) and (
                "A2" in guiData['materialCode']):
            # DataB-CHM成本
            revenueData['chmCost'] = format(
                (revenueData['revenueForCny'] - guiData['cost']) * guiData['chmCostRate'] * 0.2, '.2f')
            # DataB-PHY成本
            revenueData['phyCost'] = format(
                (revenueData['revenueForCny'] - guiData['cost']) * guiData['phyCostRate'] * 0.8, '.2f')
            # Item1000 的revenue
            revenueData['chmRe'] = format(revenueData['revenue'] * 0.2, '.2f')
            # Item2000 的revenue
            revenueData['phyRe'] = format(revenueData['revenue'] * 0.8, '.2f')
            # plan cost总算法
            # revenueData['chmCsCostAccounting'] = format(revenueData['planCost'] * 0.8 * (1 - 0.3  - (1 - guiData['planCostRate'] )) / guiData['csHourlyRate'], '.%sf' % guiData['significantDigits'])
            # revenueData['chmLabCostAccounting'] = format(revenueData['planCost'] * 0.8 * 0.3 / guiData['chmHourlyRate'], '.%sf' % guiData['significantDigits'])
            # revenueData['phyCsCostAccounting'] = format(revenueData['planCost'] * 0.2 * (1 - 0.3  - (1 - guiData['planCostRate'] )) / guiData['csHourlyRate'], '.%sf' % guiData['significantDigits'])
            # revenueData['phyLabCostAccounting'] = format(revenueData['planCost'] * 0.2 * 0.3 / guiData['chmHourlyRate'], '.%sf' % guiData['significantDigits'])

            # CS的Item1000-Cost
            revenueData['chmCsCostAccounting'] = format(
                (revenueData['revenueForCny'] - guiData['cost']) * guiData['planCostRate'] * 0.2 * (
                        1 - guiData['chmCostRate']) / guiData['csHourlyRate'],
                '.%sf' % guiData['significantDigits'])
            # CHM的Item1000-Cost
            revenueData['chmLabCostAccounting'] = format(
                (revenueData['revenueForCny'] - guiData['cost']) * guiData['planCostRate'] * 0.2 * guiData[
                    'chmCostRate'] /
                guiData['chmHourlyRate'], '.%sf' % guiData['significantDigits'])
            # CS的Item2000-Cost
            revenueData['phyCsCostAccounting'] = format(
                (revenueData['revenueForCny'] - guiData['cost']) * guiData['planCostRate'] * 0.8 * (
                        1 - guiData['phyCostRate']) / guiData['csHourlyRate'],
                '.%sf' % guiData['significantDigits'])
            # PHY的Item2000-Cost
            revenueData['phyLabCostAccounting'] = format(
                (revenueData['revenueForCny']  - guiData['cost']) * guiData['planCostRate'] * 0.8 * guiData[
                    'phyCostRate'] /
                guiData['phyHourlyRate'], '.%sf' % guiData['significantDigits'])
        else:
            revenueData['chmCost'] = format((revenueData['revenueForCny'] - guiData['cost']) * guiData['chmCostRate'],
                                            '.2f')
            revenueData['phyCost'] = format((revenueData['revenueForCny'] - guiData['cost']) * guiData['phyCostRate'],
                                            '.2f')
            revenueData['chmRe'] = format(revenueData['revenue'], '.2f')
            revenueData['phyRe'] = format(revenueData['revenue'], '.2f')
            # plan cost总算法
            # csCostAccounting = format(planCost * (1 - 0.3  - (1 - guiData['planCostRate'] )) / guiData['csHourlyRate'], '.%sf' % guiData['significantDigits'])
            # labCostAccounting = format(planCost * 0.3 / guiData['chmHourlyRate'], '.%sf' % guiData['significantDigits'])
            if 'T75' in guiData['materialCode']:
                revenueData['labCostRate'] = guiData['chmCostRate']
                revenueData['labHourlyRate'] = guiData['chmHourlyRate']
            else:
                revenueData['labCostRate'] = guiData['phyCostRate']
                revenueData['labHourlyRate'] = guiData['phyHourlyRate']

            revenueData['csCostAccounting'] = format(
                (revenueData['revenueForCny'] - guiData['cost']) * guiData['planCostRate'] * (
                        1 - revenueData['labCostRate']) / guiData[
                    'csHourlyRate'], '.%sf' % guiData['significantDigits'])
            revenueData['labCostAccounting'] = format(
                (revenueData['revenueForCny'] - guiData['cost']) * guiData['planCostRate'] * revenueData[
                    'labCostRate'] / revenueData['labHourlyRate'],
                '.%sf' % guiData['significantDigits'])
        return revenueData
    def getRevenueDataUnified(self, guiData, configContent):
        """
        统一的收入分配计算方法，使用 Revenue_Operate.py 的逻辑
        
        Args:
            guiData: GUI界面数据
            configContent: 配置内容（从 Revenue_Operate.py 传入）
            
        Returns:
            与原 getRevenueData 方法相同格式的数据
        """
        try:
            # 导入 Revenue_Operate 模块
            from Revenue_Operate import RevenueAllocator
            
            # 创建分配器实例
            allocator = RevenueAllocator()
            # 将未税的cost，更新为含税的cost，与直接获取order的Excel保持一致
            if self.checkBox_27.isChecked():
                cost = guiData.get('cost', 0) * 1.06
            else:
                cost = guiData.get('cost', 0)
            # 转换 GUI 数据为 Revenue_Operate 格式
            revenueDataForAllocation = {
                'Order Number': guiData.get('orderNo', ''),
                'Material Code': guiData.get('materialCode', ''),
                'Total Subcon Cost': cost,
                'Primary CS': guiData.get('csName', ''),
                'Tax-inclusive amount': guiData.get('amountVat', ''),
                'Rate': guiData.get('exchangeRate', ''),
            }
            
            # 调用新的分配方法，获取 result 字典格式的数据
            result = allocator.allocate_department_hours(revenueDataForAllocation, configContent, return_format='traditional')
            
            # 将 result 数据转换为原 getRevenueData 格式
            return self._convertResultToRevenueData(result, guiData, configContent)
            
        except Exception as e:
            # 如果新方法失败，回退到原方法
            print(f"统一计算方法失败，使用原方法: {e}")
            return self.getRevenueData(guiData)
    def _convertResultToRevenueData(self, result, guiData, configContent):
        """
        将 allocate_department_hours 的 result 字典直接映射为 getRevenueData 格式
        直接引用已计算好的数据，不重新计算
        """
        revenueData = {}
        
        # 基础计算 - 与原 getRevenueData 保持一致
        revenueData['revenue'] = guiData['amountVat'] / 1.06
        revenueData['planCost'] = revenueData['revenue'] * guiData['exchangeRate']
        revenueData['revenueForCny'] = revenueData['revenue'] * guiData['exchangeRate']
        
        # 获取有效位数
        significant_digits = int(configContent.get('Significant_Digits', 2))
        
        # 直接引用 result 中的计算结果
        # 根据实验室分配情况映射数据

        # 情况1: 配置不存在
        if guiData['materialCode'] not in configContent:
            revenueData['csCostAccounting'] = format(result.get('business_dept_1000_hours', 0),
                                                        f'.{significant_digits}f')
            revenueData['labCostAccounting'] = format(result.get('lab_1000_hours', 0), f'.{significant_digits}f')
            if 'T75' in guiData['materialCode']:
                revenueData['labCostRate'] = guiData['chmCostRate']
                revenueData['labHourlyRate'] = guiData['chmHourlyRate']
                revenueData['chmCost'] = format(result.get('lab_1000_act_revenue', 0), '.2f')
                revenueData['chmRe'] = revenueData['phyRe'] =  format(result.get('item_1000_amount', 0), '.2f')
                revenueData['phyCost'] = format((revenueData['revenueForCny'] - guiData['cost']) * guiData['phyCostRate'] , '.2f')

            else:
                revenueData['labCostRate'] = guiData['phyCostRate']
                revenueData['labHourlyRate'] = guiData['phyHourlyRate']
                revenueData['chmCost'] = format((revenueData['revenueForCny'] - guiData['cost']) * guiData['chmCostRate'] , '.2f')
                revenueData['phyRe'] = revenueData['chmRe'] = format(result.get('item_1000_amount', 0), '.2f')
                revenueData['phyCost'] = format(result.get('lab_1000_act_revenue', 0), '.2f')
        else:
            # CHM 相关数据
            if result.get('lab_1000') == 'CHM':
                # CHM 在 Item1000 中
                revenueData['chmCost'] = format(result.get('lab_1000_act_revenue', 0), '.2f')
                revenueData['chmRe'] = format(result.get('item_1000_amount', 0), '.2f')
                revenueData['chmCsCostAccounting'] = format(result.get('business_dept_1000_hours', 0), f'.{significant_digits}f')
                revenueData['chmLabCostAccounting'] = format(result.get('lab_1000_hours', 0), f'.{significant_digits}f')
            elif result.get('lab_2000') == 'CHM':
                # CHM 在 Item2000 中
                revenueData['chmCost'] = format(result.get('lab_2000_act_revenue', 0) , '.2f')
                revenueData['chmRe'] = format(result.get('item_2000_amount', 0), '.2f')
                revenueData['chmCsCostAccounting'] = format(result.get('business_dept_2000_hours', 0), f'.{significant_digits}f')
                revenueData['chmLabCostAccounting'] = format(result.get('lab_2000_hours', 0), f'.{significant_digits}f')
            else:
                # CHM 没有分配
                revenueData['chmCost'] = '0.00'
                revenueData['chmRe'] = '0.00'
                revenueData['chmCsCostAccounting'] = '0'
                revenueData['chmLabCostAccounting'] = '0'

            # PHY 相关数据
            if result.get('lab_1000') == 'PHY':
                # PHY 在 Item1000 中
                revenueData['phyCost'] = format(result.get('lab_1000_act_revenue', 0), '.2f')
                revenueData['phyRe'] = format(result.get('item_1000_amount', 0), '.2f')
                revenueData['phyCsCostAccounting'] = format(result.get('business_dept_1000_hours', 0), f'.{significant_digits}f')
                revenueData['phyLabCostAccounting'] = format(result.get('lab_1000_hours', 0), f'.{significant_digits}f')
            elif result.get('lab_2000') == 'PHY':
                # PHY 在 Item2000 中
                revenueData['phyCost'] = format(result.get('lab_2000_act_revenue', 0), '.2f')
                revenueData['phyRe'] = format(result.get('item_2000_amount', 0), '.2f')
                revenueData['phyCsCostAccounting'] = format(result.get('business_dept_2000_hours', 0), f'.{significant_digits}f')
                revenueData['phyLabCostAccounting'] = format(result.get('lab_2000_hours', 0), f'.{significant_digits}f')
            else:
                # PHY 没有分配
                revenueData['phyCost'] = '0.00'
                revenueData['phyRe'] = '0.00'
                revenueData['phyCsCostAccounting'] = '0'
                revenueData['phyLabCostAccounting'] = '0'
        
        return revenueData
    def sap_operate(self, guiData=None, revenueData=None, sap_obj=None, include_followup=True):
        direct_gui_call = not isinstance(guiData, dict) and not isinstance(revenueData, dict) and sap_obj is None
        # 主按钮使用的统一 SAP 订单流程。
        # include_followup=False 用于兼容旧的 sapOperate()；
        # 旧流程仍会单独执行 Data B、VA02、发票等后续步骤。
        def _to_float(value, default=0.0):
            try:
                if value in ("", None):
                    return default
                return float(value)
            except (TypeError, ValueError):
                return default

        def _to_str(value):
            if value is None:
                return ''
            return str(value).strip()

        def _build_order_items(data):
            raw_items = data.get('items') or []
            items = []
            if isinstance(raw_items, list):
                for raw_item in raw_items:
                    if not isinstance(raw_item, dict):
                        continue
                    material_code = _to_str(raw_item.get('material_code', raw_item.get('materialCode')))
                    if not material_code:
                        continue
                    items.append(OrderItemData(
                        item=_to_str(raw_item.get('item')),
                        material_code=material_code,
                        revenue=_to_float(raw_item.get('revenue', raw_item.get('amount'))),
                        quantity=_to_str(raw_item.get('quantity')) or '1',
                        unit=_to_str(raw_item.get('unit')) or 'pu',
                    ))
            return items

        # 存放本次 SAP 操作的结构化步骤日志。
        operation_steps = []

        def _add_step(step, success=True, msg='', order_no='', sap_amount_vat=''):
            # 记录单个操作步骤，供界面显示、Excel 日志和问题排查共用。
            operation_steps.append({
                'step': step,
                'status': 'success' if success else 'failed',
                'msg': msg or '',
                'orderNo': order_no or '',
                'sapAmountVat': sap_amount_vat or '',
                'time': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            })

        def _format_steps_remark():
            # 将结构化步骤日志汇总为旧 Remark 字段，兼容现有导出格式。
            remark_items = []
            for item in operation_steps:
                status_text = '成功' if item['status'] == 'success' else '失败'
                item_msg = f":{item['msg']}" if item['msg'] else ''
                remark_items.append(f"{item['step']}[{status_text}]{item_msg}")
            return ';'.join(remark_items)

        def _attach_log(data):
            # 在原返回结构上附加新日志字段，同时保留旧字段。
            data['steps'] = list(operation_steps)
            data['Remark'] = _format_steps_remark()
            return data

        def _html_escape(value):
            return str(value or '').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

        def _finish(data):
            data = _attach_log(data)
            if direct_gui_call and hasattr(self, 'textBrowser'):
                if data.get('flag') == 0:
                    message = data.get('msg') or data.get('Remark') or 'SAP operation failed.'
                    self.textBrowser.append("<font color='red'>SAP Error: %s</font>" % _html_escape(message))
                    for step in data.get('steps', []):
                        if step.get('status') == 'failed':
                            detail = "%s: %s" % (step.get('step', ''), step.get('msg', ''))
                            self.textBrowser.append("<font color='red'>%s</font>" % _html_escape(detail))
                    self.textBrowser.append('----------------------------------')
                    QApplication.processEvents()
                    QMessageBox.warning(self, "SAP Error", str(message), QMessageBox.Yes)
                else:
                    order_no = data.get('orderNo') or ''
                    if order_no:
                        self.textBrowser.append("Order No.:%s" % order_no)
                    self.textBrowser.append('SAP operation completed.')
                    self.textBrowser.append('----------------------------------')
                    QApplication.processEvents()
            return data

        def _legacy_result(result, **extra):
            # 保持旧 GUI 日志代码需要的返回结构。
            data = {
                'flag': 1 if result.success else 0,
                'msg': result.message,
                'orderNo': getattr(result, 'order_no', '') or '',
                'Proforma No.': getattr(result, 'proforma_no', '') or '',
                'sapAmountVat': getattr(result, 'sap_amount_vat', '') or '',
            }
            data.update(extra)
            return data

        def _extract_order_no(session):
            # 优先读取订单号字段；保存后再从状态栏兜底提取订单号。
            try:
                order_no = str(session.read_text("wnd[0]/usr/ctxtVBAK-VBELN")).strip()
                if order_no:
                    return order_no
            except Exception:
                pass

            try:
                status_text = session.read_status()
            except Exception:
                return ''

            match = re.search(r"(\d{6,})", status_text)
            return match.group(1) if match else ''

        if not isinstance(guiData, dict):
            # PyQt clicked 信号可能会把 checked 布尔值传入槽函数。
            guiData = None
        if not isinstance(revenueData, dict):
            revenueData = None

        guiData = guiData or self.__class__.getGuiData(self)
        revenueData = revenueData or self.__class__.getRevenueDataUnified(self, guiData, configContent)

        sap_session = None
        own_session = False
        if sap_obj and getattr(sap_obj, "_sap_session", None):
            sap_session = sap_obj._sap_session

        try:
            if sap_session is None:
                # 直接点击按钮时自行创建会话；批量模式可复用 sap_obj 的会话。
                sap_session = SapSession.connect()
                own_session = True

            # 将 GUI 字典转换为 sap 服务层使用的类型化入参。
            config = SapConfig(
                order_type=str(guiData.get('orderType', '')).strip(),
                sales_organization=str(guiData.get('salesOrganization', '')).strip(),
                distribution_channels=str(guiData.get('distributionChannels', '')).strip(),
                sales_office=str(guiData.get('salesOffice', '')).strip(),
                # 240
                cost_center=str(guiData.get('salesGroup', '')).strip(),
                sub_cost_center_cs=str(guiData.get('csCostCenter', '')).strip(),
                sub_cost_center_chm=str(guiData.get('chmCostCenter', '')).strip(),
                sub_cost_center_phy=str(guiData.get('phyCostCenter', '')).strip(),
                cs_code=str(guiData.get('csCode', '')).strip(),
                sales_code=str(guiData.get('salesCode', '')).strip(),
                data_ae1=guiData.get('dataAE1', []),
                data_az2=guiData.get('dataAZ2', []),
            )

            order = OrderData(
                sap_no=str(guiData.get('sapNo', '')).strip(),
                project_no=str(guiData.get('projectNo', '')).strip(),
                material_code=str(guiData.get('materialCode', '')).strip(),
                currency_type=str(guiData.get('currencyType', '')).strip(),
                exchange_rate=_to_float(guiData.get('exchangeRate'), 1.0),
                cost=_to_float(guiData.get('cost')),
                short_text=str(guiData.get('shortText', '')).strip(),
                amount_vat=_to_float(guiData.get('amountVat')),
                long_text=str(guiData.get('longText', '')).strip(),
                global_partner_code=str(guiData.get('globalPartnerCode', '')).strip(),
                sales_name=str(guiData.get('salesName', '')).strip(),
                ecd=time.strftime("%Y.%m.%d"),
                items=_build_order_items(guiData),
            )

            revenue = RevenueData(
                revenue=_to_float(revenueData.get('revenue'), _to_float(guiData.get('amount'))),
                revenue_cny=_to_float(
                    revenueData.get('revenueForCny'),
                    _to_float(guiData.get('amount')) * _to_float(guiData.get('exchangeRate'), 1.0),
                ),
                chm_cost=_to_float(revenueData.get('chmCost')),
                phy_cost=_to_float(revenueData.get('phyCost')),
                chm_revenue=_to_float(revenueData.get('chmRe')),
                phy_revenue=_to_float(revenueData.get('phyRe')),
                chm_cs_cost=_to_float(revenueData.get('chmCsCostAccounting')),
                chm_lab_cost=_to_float(revenueData.get('chmLabCostAccounting')),
                phy_cs_cost=_to_float(revenueData.get('phyCsCostAccounting')),
                phy_lab_cost=_to_float(revenueData.get('phyLabCostAccounting')),
                cs_cost=_to_float(revenueData.get('csCostAccounting')),
                lab_cost=_to_float(revenueData.get('labCostAccounting')),
            )

            service = OrderService(sap_session, config)
            final_res = {
                'flag': 1,
                'msg': '',
                'orderNo': '',
                'Proforma No.': '',
                'sapAmountVat': '',
                'steps': [],
                'Remark': '',
            }

            if guiData.get('va01Check', True):
                # 步骤 1：VA01 创建订单抬头，并写入伙伴/文本等数据。
                result = service.create_order(
                    order,
                    revenue,
                    partner_options=PartnerOptions(
                        add_contact=bool(guiData.get('contactCheck', True)),
                        add_sales_partner=bool(order.sales_name),
                    ),
                )
                _add_step('VA01', result.success, result.message, result.order_no, result.sap_amount_vat)
                final_res = _legacy_result(result)
                if not result.success or not include_followup:
                    return _finish(final_res)

                if guiData.get('labCostCheck'):
                    # 步骤 2：Data B 在订单抬头写入 lab cost 行。
                    data_b_result = service.fill_lab_cost(order, revenue)
                    _add_step('Data B', data_b_result.success, data_b_result.message, final_res['orderNo'])
                    if not data_b_result.success:
                        return _finish(_legacy_result(data_b_result, orderNo=final_res['orderNo']))
                    if data_b_result.message:
                        final_res['msg'] = data_b_result.message

                if guiData.get('va02Check') or guiData.get('saveCheck'):
                    # VA02 需要已保存的订单号，所以先保存 VA01。
                    save_result = service.save('VA01')
                    order_no = _extract_order_no(sap_session)
                    _add_step('Save VA01', save_result.success, save_result.message, order_no or final_res['orderNo'])
                    if sap_obj is not None and hasattr(sap_obj, 'current_order_no'):
                        sap_obj.current_order_no = order_no or getattr(sap_obj, 'current_order_no', '')
                    if not save_result.success:
                        return _finish(_legacy_result(save_result, orderNo=order_no or final_res['orderNo']))
                    final_res['orderNo'] = order_no or final_res['orderNo']

            if include_followup and guiData.get('va02Check'):
                # 步骤 3：VA02 重新打开订单，并追加 item 行和金额。
                order_no = (
                    final_res.get('orderNo')
                    or getattr(sap_obj, 'current_order_no', '')
                    or _extract_order_no(sap_session)
                )
                if not order_no:
                    _add_step('VA02', False, '未找到可用于VA02的Order No.')
                    return _finish({
                        'flag': 0,
                        'msg': '未找到可用于VA02的Order No.',
                        'orderNo': '',
                        'Proforma No.': '',
                        'sapAmountVat': final_res.get('sapAmountVat', ''),
                        'steps': list(operation_steps),
                        'Remark': _format_steps_remark(),
                    })

                open_result = service.open_order(order_no)
                _add_step('Open VA02', open_result.success, open_result.message, order_no, final_res.get('sapAmountVat', ''))
                if not open_result.success:
                    return _finish(_legacy_result(open_result, orderNo=order_no, sapAmountVat=final_res.get('sapAmountVat', '')))

                item_result = service.add_items(order, revenue)
                _add_step('VA02', item_result.success, item_result.message, item_result.order_no or order_no, item_result.sap_amount_vat)
                final_res = _legacy_result(item_result, orderNo=item_result.order_no or order_no)
                if sap_obj is not None and hasattr(sap_obj, 'current_order_no'):
                    sap_obj.current_order_no = final_res['orderNo']
                if not item_result.success:
                    return _finish(final_res)

                if guiData.get('planCostCheck'):
                    # 步骤 4：Plan Cost 为可选操作，并按 CS/CHM/PHY 勾选项执行。
                    cost_result = service.apply_plan_cost(
                        order,
                        revenue,
                        cost_options=CostOptions(
                            include_cs=bool(guiData.get('csCheck', True)),
                            include_chm=bool(guiData.get('chmCheck', True)),
                            include_phy=bool(guiData.get('phyCheck', True)),
                        ),
                    )
                    _add_step('Plan Cost', cost_result.success, cost_result.message, final_res['orderNo'], final_res.get('sapAmountVat', ''))
                    if not cost_result.success:
                        return _finish(_legacy_result(
                            cost_result,
                            orderNo=final_res['orderNo'],
                            sapAmountVat=final_res.get('sapAmountVat', ''),
                        ))
                    if cost_result.message:
                        final_res['msg'] = cost_result.message

                if guiData.get('vf01Check') or guiData.get('saveCheck'):
                    # 创建发票或结束流程前，先保存 VA02。
                    save_result = service.save('VA02')
                    _add_step('Save VA02', save_result.success, save_result.message, final_res['orderNo'], final_res.get('sapAmountVat', ''))
                    if not save_result.success:
                        return _finish(_legacy_result(
                            save_result,
                            orderNo=final_res['orderNo'],
                            sapAmountVat=final_res.get('sapAmountVat', ''),
                        ))

            return _finish(final_res)
        except Exception as exc:
            _add_step('SAP操作', False, str(exc))
            return _finish({
                'flag': 0,
                'msg': f'VA01调用失败：{exc}',
                'orderNo': '',
                'Proforma No.': '',
                'sapAmountVat': '',
                'steps': list(operation_steps),
                'Remark': _format_steps_remark(),
            })
        finally:
            if own_session and sap_session is not None:
                sap_session.close()
    def sapOperate(self, sap_obj):
        logMsg = {}
        logMsg['Remark'] = ''
        logMsg['orderNo'] = ''
        logMsg['Proforma No.'] = ''
        logMsg['sapAmountVat'] = ''
        try:
            flag = 1
            # 获取数据
            guiData = self.__class__.getGuiData(self)
            orderNo = ''
            proformaNo = ''
            if guiData['everyCheck'] or not hasattr(sap_obj, 'va01_operate'):
                sap_obj = Sap()
            if guiData['sapNo'] == '' or guiData['projectNo'] == '' or guiData['materialCode'] == '' or guiData[
                'currencyType'] == '' or guiData['exchangeRate'] == '' or guiData['globalPartnerCode'] == '' or guiData[
                'csName'] == '' or guiData['amount'] == 0.00 or guiData['amountVat'] == 0.00:
                self.textBrowser.append("<font color='red'>有关键信息未填</font>")
                logMsg['Remark'] = '有关键信息未填'
                self.textBrowser.append(
                    "'Project No.', 'CS', 'Sales', 'Currency', 'GPC Glo. Par. Code', 'Material Code','SAP No.', 'Amount', 'Amount with VAT', 'Exchange Rate'都是必须填写的")
                self.textBrowser.append('----------------------------------')
                QApplication.processEvents()
                if guiData['everyCheck']:
                    QMessageBox.information(self, "提示信息", "有关键信息未填", QMessageBox.Yes)
            else:
                # 使用新的统一计算方法
                revenueData = self.__class__.getRevenueDataUnified(self, guiData, configContent)
                # revenueData = self.__class__.getRevenueData(self, guiData)

                messageFlag = 1
                if self.checkBox_5.isChecked():
                    if guiData['salesName'] == '':
                        reply = QMessageBox.question(self, '信息', 'Sales未填，是否继续', QMessageBox.Yes | QMessageBox.No,
                                                     QMessageBox.Yes)
                        if reply == QMessageBox.Yes:
                            messageFlag = 1
                        else:
                            messageFlag = 2
                if guiData['salesName'] != '' or messageFlag == 1:
                    self.textBrowser.append("Sap No.:%s" % guiData['sapNo'])
                    self.textBrowser.append("Project No.:%s" % guiData['projectNo'])
                    self.textBrowser.append("Material Code:%s" % guiData['materialCode'])
                    self.textBrowser.append("Global Partner Code:%s" % guiData['globalPartnerCode'])
                    self.textBrowser.append("CS Name:%s" % guiData['csName'])
                    self.textBrowser.append("Sales Name:%s" % guiData['salesName'])
                    self.textBrowser.append("Amount:%s" % guiData['amount'])
                    self.textBrowser.append("Cost:%s" % guiData['cost'])
                    self.textBrowser.append("Currency Type:%s" % guiData['currencyType'])
                    self.textBrowser.append("CHM Cost:%s" % revenueData['chmCost'])
                    self.textBrowser.append("PHY Cost:%s" % revenueData['phyCost'])
                    self.textBrowser.append("CHM Amount:%s" % revenueData['chmRe'])
                    self.textBrowser.append("PHY Amount:%s" % revenueData['phyRe'])
                    QApplication.processEvents()

                    flag = 1
                    # VA01
                    if guiData['va01Check']:
                        va01_res = self.sap_operate(guiData, revenueData, sap_obj, include_followup=False)
                        if va01_res['flag'] == 1:
                            # 是否要添加lab cost
                            if guiData['labCostCheck'] and va01_res['flag'] == 1:
                                data_b_res = sap_obj.lab_cost(guiData, revenueData)
                                if data_b_res['flag'] == 0:
                                    logMsg['Remark'] += data_b_res['msg']
                                    self.textBrowser.append("<font color='red'>出错信息：%s </font>" % data_b_res['msg'])
                                    QApplication.processEvents()
                                    if guiData['everyCheck']:
                                        QMessageBox.information(self, "错误提示", "出错信息：%s" % data_b_res['msg'],
                                                                QMessageBox.Yes)
                            if guiData['va02Check'] or guiData['saveCheck']:
                                save_res = sap_obj.save_sap('VA01')
                                if save_res['flag'] == 0:
                                    flag = 0
                                    logMsg['Remark'] += ';' + save_res['msg']
                                    self.textBrowser.append("<font color='red'>出错信息：%s </font>" % save_res['msg'])
                                    QApplication.processEvents()
                                    if guiData['everyCheck']:
                                        QMessageBox.information(self, "错误提示", "出错信息：%s" % save_res['msg'],
                                                                QMessageBox.Yes)
                        else:
                            flag = 0
                            logMsg['Remark'] += va01_res['msg']
                            self.textBrowser.append("<font color='red'>出错信息：VA01出错；%s </font>" % va01_res['msg'])
                            QApplication.processEvents()
                            if guiData['everyCheck']:
                                QMessageBox.information(self, "错误提示", "出错信息：%s" % va01_res['msg'],
                                                        QMessageBox.Yes)
                    # VA02
                    if guiData['va02Check'] and flag == 1:
                        va02_res = sap_obj.va02_operate(guiData, revenueData)
                        logMsg['orderNo'] = va02_res['orderNo']
                        self.textBrowser.append("Order No.:%s" % logMsg['orderNo'])
                        QApplication.processEvents()
                        if va02_res['flag'] == 1:
                            amountVatStr = re.sub(r"(\d)(?=(\d\d\d)+(?!\d))", r"\1,",
                                                  format(guiData['amountVat'], '.2f'))
                            sapAmountVat = va02_res['sapAmountVat']
                            self.textBrowser.append("Sap Amount Vat:%s" % sapAmountVat)
                            self.textBrowser.append("Amount Vat:%s" % amountVatStr)
                            QApplication.processEvents()
                            # sapAmountVat在A2是数字，其它为字符串
                            if sapAmountVat.strip() != amountVatStr:
                                self.textBrowser.append("<font color='blue'>提示信息：SAP数据与ODM不一致，请确认并修改后再继续！！！ </font>")
                                QApplication.processEvents()
                                logMsg['Remark'] += ';' + 'SAP数据与ODM不一致，请确认并修改后再继续！！！'
                                if guiData['everyCheck']:
                                    reply = QMessageBox.question(self, '信息', 'SAP数据与ODM不一致，请确认并修改后再继续！！！',
                                                                 QMessageBox.Yes | QMessageBox.No,
                                                                 QMessageBox.Yes)
                                    if reply == QMessageBox.No:
                                        flag = 0
                            if (guiData['vf01Check'] or guiData['saveCheck']) and flag == 1:
                                sava_res = sap_obj.save_sap('VA02')
                                if sava_res['flag'] == 0:
                                    flag = 0
                                    logMsg['Remark'] += ';' + sava_res['msg']
                                    self.textBrowser.append("<font color='red'>出错信息：%s </font>" % sava_res['msg'])
                                    QApplication.processEvents()
                                    if guiData['everyCheck']:
                                        QMessageBox.information(self, "错误提示", "出错信息：%s" % sava_res['msg'],
                                                                QMessageBox.Yes)
                        else:
                            flag = 0
                            logMsg['Remark'] += ';' + va02_res['msg']
                            self.textBrowser.append("<font color='red'>出错信息：VA02出错；%s </font>" % va02_res['msg'])
                            QApplication.processEvents()
                            if guiData['everyCheck']:
                                QMessageBox.information(self, "错误提示", "出错信息：%s" % va02_res['msg'],
                                                        QMessageBox.Yes)

                    # VF01
                    if guiData['vf01Check'] and flag == 1:

                        # save_res = sap_obj.save_sap('VF01准备前')
                        # if save_res['flag'] == 0:
                        #     logMsg['Remark'] += ';' + save_res['msg']
                        #     self.textBrowser.append("<font color='red'>出错信息：%s </font>" % save_res['msg'])
                        #     QApplication.processEvents()
                        #     if guiData['everyCheck']:
                        #         QMessageBox.information(self, "错误提示", "出错信息：%s" % save_res['msg'],
                        #                                 QMessageBox.Yes)
                        vf01_res = sap_obj.vf01_operate()
                        if vf01_res['flag'] == 0:
                            flag = 0
                            logMsg['Remark'] += ';' + vf01_res['msg']
                            self.textBrowser.append("<font color='red'>出错信息：VF01出错；%s </font>" % vf01_res['msg'])
                            QApplication.processEvents()
                            if guiData['everyCheck']:
                                QMessageBox.information(self, "错误提示", "出错信息：%s" % vf01_res['msg'],
                                                        QMessageBox.Yes)
                    # VF03
                    if guiData['vf03Check'] and flag == 1:
                        vf03_res = sap_obj.vf03_operate()
                        if vf03_res['flag'] == 0:
                            logMsg['Remark'] += ';' + vf03_res['msg']
                            self.textBrowser.append("<font color='red'>出错信息：VF03出错;%s </font>" % vf03_res['msg'])
                            QApplication.processEvents()
                            if guiData['everyCheck']:
                                QMessageBox.information(self, "错误提示", "出错信息：%s" % vf03_res['msg'],
                                                        QMessageBox.Yes)
                        proformaNo = vf03_res['Proforma No.']
                        logMsg['Proforma No.'] = proformaNo
                        self.textBrowser.append("Proforma No.:%s" % proformaNo)
                        QApplication.processEvents()
                    self.textBrowser.append('SAP操作已完成')
                    self.textBrowser.append('----------------------------------')
                    QApplication.processEvents()
                    if self.checkBox_5.isChecked():
                        QMessageBox.information(self, "提示信息", "SAP操作已完成", QMessageBox.Yes)

            return logMsg

        except Exception as msg:
            guiData = self.__class__.getGuiData(self)
            self.textBrowser.append('这单%s的数据或者SAP有问题' % guiData['projectNo'])
            self.textBrowser.append('错误信息：%s' % msg)
            logMsg['Remark'] += '错误信息：%s' % msg
            self.textBrowser.append('----------------------------------')
            QMessageBox.information(self, "提示信息", '这单%s的数据或者SAP有问题' % guiData['projectNo'], QMessageBox.Yes)
            return logMsg
    def orderUnlockOrLock(self, flag):
        fileUrl = self.lineEdit_6.text()
        (filepath, filename) = os.path.split(fileUrl)
        if fileUrl:
            log_file_name = 'log %s.xlsx' % time.strftime('%Y-%m-%d %H.%M.%S')
            Log_file = '%s\\%s' % (filepath, log_file_name)
            log_obj = Logger(Log_file, ['Update', 'Order No', 'Type', 'Remark'])
            newData = Get_Data()
            file_data = newData.getFileData(fileUrl)
            order_list = list(file_data['Order No'])
            if not self.checkBox_16.isChecked():
                sap_obj = Sap()
            i = 1
            for orderNo in order_list:
                try:
                    log_list = {}
                    log_list['Order No'] = orderNo
                    log_list['Type'] = flag

                    if self.checkBox_16.isChecked():
                        sap_obj = Sap()
                    sap_obj.open_va02(orderNo)
                    lock_res = sap_obj.unlock_or_lock_order(flag)
                    self.textBrowser.append('%s.Order No: %s' % (i, orderNo))
                    self.textBrowser.append('%s' % lock_res['msg'])
                    QApplication.processEvents()
                    if not sap_obj.res['flag']:
                        log_list['Remark'] = lock_res['msg']
                    else:
                        log_list['Remark'] = ''
                    log_obj.log(log_list)
                    i += 1
                except:
                    self.textBrowser.append("<font color='red'>该Order: %s 有问题</font>" % orderNo)
                    QApplication.processEvents()
            log_obj.save_log_to_excel()
            self.textBrowser.append('%s' % Log_file)
            QApplication.processEvents()
            os.startfile(Log_file)
        else:
            self.textBrowser.append('没有文件请添加')
            QApplication.processEvents()

