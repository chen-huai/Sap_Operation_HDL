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
from sap import OrderData, OrderItemData, OrderService, PartnerOptions, RevenueData, SapConfig, SapSession
from runtime_globals import configContent

class SapOrderMixin:
    @staticmethod
    def _excel_value(value, default=''):
        """读取 Excel 单元格原始值；空值统一转换为默认值。"""
        if pd.isna(value):
            return default
        return value

    @staticmethod
    def _excel_str(value, default=''):
        """读取 Excel 单元格文本值；数字编号会去掉无意义的小数位。"""
        if pd.isna(value):
            return default
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value).strip()

    @staticmethod
    def _excel_float(value, default=0.0):
        """读取 Excel 单元格数值；无法转换时返回默认值。"""
        if pd.isna(value) or value == '':
            return default
        try:
            return float(value)
        except (TypeError, ValueError):
            return default

    def _filter_related_rows(self, dataframe, order_row):
        """按 Combine Id 严格筛选当前订单对应的明细行；调用前需确保 Combine Id 存在。"""
        if dataframe.empty:
            return dataframe
        combine_id = self._excel_value(order_row.get('Combine Id'))
        return dataframe[dataframe['Combine Id'] == combine_id]

    def _build_sap_config_from_order_row(self, order_row):
        """按当前订单行和系统配置构建 SAP 固定参数。"""
        cs_name = self._excel_str(order_row.get('Primary CS'))
        sales_name = self._excel_str(order_row.get('Sales'))
        return SapConfig(
            # 登录界面
            order_type=self.lineEdit_10.text(),
            sales_organization=self.lineEdit_11.text(),
            distribution_channels=self.lineEdit_12.text(),
            sales_office=self.lineEdit_13.text(),
            sales_group=self.lineEdit_14.text(),
            # TODO 业务和成本中心，可能可以删除
            sub_cost_center_cs=self.lineEdit_18.text(),
            sub_cost_center_chm=self.lineEdit_19.text(),
            sub_cost_center_phy=self.lineEdit_20.text(),
            # cs和sales
            cs_code=configContent.get(cs_name, ''),
            sales_code=configContent.get(sales_name, ''),
            # 客户选择是否为海外订单
            data_ae1=self.lineEdit_21.text().split(';'),
            data_az2=self.lineEdit_22.text().split(';'),
        )

    def _build_order_from_dataframes(self, order_row, item_df):
        """从订单头和 item 表构建 SAP 订单对象。"""
        order_items_df = self._filter_related_rows(item_df, order_row)
        items = []
        for _, item_row in order_items_df.iterrows():
            # item 表已经包含 SAP item 所需的物料和金额，不再从 GUI 或规则计算。
            items.append(OrderItemData(
                item=self._excel_str(item_row.get('item')),
                material_code=self._excel_str(item_row.get('Item Material Code')),
                long_text=self._excel_str(item_row.get('Item Group Description')),
                revenue=self._excel_float(item_row.get('Item price')),
                quantity='1',
                unit='pu',
            ))

        return OrderData(
            sap_no=self._excel_str(order_row.get('SAP Customer Code')),
            project_no=self._excel_str(order_row.get('Request Number')),
            amount_vat=self._excel_str(order_row.get('Tax-inclusive amount')),
            currency_type=self._excel_str(order_row.get('Currency')),
            exchange_rate=self._excel_float(order_row.get('Rate'), 1.0),
            short_text=self._excel_str(order_row.get('Additional Information')),
            global_partner_code=self._excel_str(order_row.get('GPC Code')),
            sales_name=self._excel_str(order_row.get('Sales')),
            sales_group=self._excel_str(order_row.get('Cost Center'))[-3:],
            ecd=self._excel_str(order_row.get('Ecd')),
            order_cost_center=self._excel_str(order_row.get('Order Center')),
            items=items,
        )

    def _build_revenue_from_order_row(self, order_row):
        """从订单表读取 Revenue；只做对象适配，不再重新分配或计算业务金额。"""
        revenue = self._excel_float(order_row.get('Revenue'), self._excel_float(order_row.get('Untaxed amount')))
        rate = self._excel_float(order_row.get('Rate'), 1.0)
        return RevenueData(
            revenue=revenue,
            revenue_cny=revenue * rate,
        )

    def _build_sub_entries_from_dataframe(self, order_row, sub_df):
        """从 sub 表构建 Data B 和 Plan Cost 的直接写入明细。"""
        order_sub_df = self._filter_related_rows(sub_df, order_row)
        data_b_entries = []
        plan_cost_entries_by_item = {}
        order_cost_center = self._excel_str(order_row.get('Cost Center'))

        for _, sub_row in order_sub_df.iterrows():
            item_no = self._excel_str(sub_row.get('item'))
            sub_cost_center = self._excel_str(sub_row.get('Sub Site Cost Center'))
            sub_cost = self._excel_float(sub_row.get('Sub-Cost RMB'))
            plan_hour = self._excel_float(sub_row.get('Sub Site Plan Hour'))

            if sub_cost_center and sub_cost:
                # Data B 只写 Sub Site Cost Center 和 Sub-Cost RMB 都有值的行。
                # 携带 item 号供 SAP POSNR 字段定位；多 item（如 "1000;3000"）由下游裁剪。
                data_b_entries.append({
                    'performer_cost_center': sub_cost_center,
                    'rate_cost_center': sub_cost_center,
                    'amount': sub_cost,
                    'item': item_no,
                })

                # Plan Cost 中 FREMDL 按对应 item 汇总 Sub-Cost RMB。
                plan_cost_entries_by_item.setdefault(item_no, []).append({
                    'cost_center': sub_cost_center,
                    'category': 'FREMDL',
                    'amount': sub_cost,
                })

            if order_cost_center and plan_hour:
                # Plan Cost 中 T01AST 使用订单信息 Cost Center + Sub Site Plan Hour。
                plan_cost_entries_by_item.setdefault(item_no, []).append({
                    'cost_center': order_cost_center,
                    'category': 'T01AST',
                    'amount': plan_hour,
                })

        summarized_plan_cost_entries = {}
        for item_no, entries in plan_cost_entries_by_item.items():
            summary = {}
            for entry in entries:
                key = (entry['cost_center'], entry['category'])
                summary[key] = summary.get(key, 0) + entry['amount']
            summarized_plan_cost_entries[item_no] = [
                {
                    'cost_center': cost_center,
                    'category': category,
                    'amount': amount,
                }
                for (cost_center, category), amount in summary.items()
                if amount
            ]

        return data_b_entries, summarized_plan_cost_entries

    @staticmethod
    def _find_item_row(order, item_no):
        """根据 item 编号找到 SAP item 表格中的行号。"""
        for row, item in enumerate(order.items):
            if item.item == item_no:
                return row
        return 0

    @staticmethod
    def _extract_order_no(session):
        """优先读取 VBELN，其次从状态栏中提取已保存的订单号。"""
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

    def odmDataToSap(self):
        """
        从 SAP_data 多 sheet 文件批量创建 SAP Order。

        业务数据来源：
            订单信息：创建订单头、伙伴、币种、Revenue 等订单级数据。
            item：创建 SAP item 行和 item 金额。
            sub：创建 Data B 和 Plan Cost 明细。
        """
        fileUrl = self.lineEdit_6.text()
        if not fileUrl:
            QMessageBox.information(self, "提示信息", "请选择订单数据文件", QMessageBox.Yes)
            return

        sap_session = None
        try:
            # 订单业务字段来自 Excel；事务流开关仍使用 GUI 复选框控制。
            flow_options = self.__class__.getGuiData(self)
            sheets = Get_Data().getExcelSheetsData(fileUrl)
            order_df = sheets['订单信息']
            item_df = sheets['item']
            sub_df = sheets.get('sub', pd.DataFrame())

            # 订单和 item 表 Combine Id 必须完整；sub 表 Combine Id 不要求对应。
            missing_order_rows = order_df.index[order_df['Combine Id'].isna()].tolist()
            missing_item_rows = item_df.index[item_df['Combine Id'].isna()].tolist()
            if missing_order_rows:
                self.textBrowser.append(
                    "<font color='red'>订单信息表 Combine Id 为空的行号: %s</font>" % missing_order_rows
                )
            if missing_item_rows:
                self.textBrowser.append(
                    "<font color='red'>item 表 Combine Id 为空的行号: %s</font>" % missing_item_rows
                )
            if missing_order_rows or missing_item_rows:
                QApplication.processEvents()

            # 每次批量处理生成一份 log，保留原订单表字段并追加 SAP 执行结果。
            filepath, _ = os.path.split(fileUrl)
            log_file_url = os.path.join(filepath, 'log')
            self.__class__.createFolder(self, log_file_url)
            log_data_path = self.__class__.getFileName(self, log_file_url, 'log', 'xlsx')
            log_file = order_df.copy()
            log_file['Order No.'] = ''
            log_file['Remark'] = ''
            log_file['Proforma No.'] = ''
            log_file['sapAmountVat'] = ''
            log_file['Update Time'] = '未开Order'

            sap_session = SapSession.connect()

            for index, order_row in order_df.iterrows():
                # Combine Id 是关联 item / sub 的唯一键，缺失直接跳过当前订单。
                if pd.isna(order_row.get('Combine Id')):
                    log_file.loc[index, 'Remark'] = '缺失 Combine Id，无法关联 item/sub'
                    log_file.to_excel(log_data_path, merge_cells=False, index=False)
                    continue

                # 三张 DataFrame 已包含完整业务数据，这里只做对象适配和 SAP 写入。
                order = self._build_order_from_dataframes(order_row, item_df)
                revenue = self._build_revenue_from_order_row(order_row)
                config = self._build_sap_config_from_order_row(order_row)
                data_b_entries, plan_cost_entries_by_item = self._build_sub_entries_from_dataframe(order_row, sub_df)
                service = OrderService(sap_session, config)

                if not order.sap_no or not order.project_no or not order.items:
                    log_file.loc[index, 'Remark'] = '关键订单信息缺失'
                    log_file.to_excel(log_data_path, merge_cells=False, index=False)
                    self.textBrowser.append(
                        "<font color='red'>No.%s 关键订单信息缺失（SAP No./Project No./items 任一为空）</font>"
                        % (index + 1)
                    )
                    QApplication.processEvents()
                    continue

                # textBrowser 抬头：基础信息 + Excel 含税金额。
                combine_id = self._excel_str(order_row.get('Combine Id'))
                primary_cs = self._excel_str(order_row.get('Primary CS'))
                sales_name = self._excel_str(order_row.get('Sales'))
                excel_amount_vat = self._excel_str(order_row.get('Tax-inclusive amount'))
                items_revenue_total = sum(item.revenue for item in order.items)

                self.textBrowser.append('==================== No.%s ====================' % (index + 1))
                self.textBrowser.append("Combine Id: %s" % combine_id)
                self.textBrowser.append("Request Number: %s" % order.project_no)
                self.textBrowser.append("Primary CS: %s" % primary_cs)
                self.textBrowser.append("Sales: %s" % sales_name)
                self.textBrowser.append("含税金额(Excel): %s" % excel_amount_vat)
                self.textBrowser.append("Items 加和金额: %s" % format(items_revenue_total, ',.2f'))
                QApplication.processEvents()

                remarks = []
                # 业务流程：VA01(可选) → Save VA01 → 打开 VA02(可选) → Add Item → Data B(可选) → Plan Cost(可选) → Save VA02
                # 所有步骤独立可选；跳过 VA01 时 order_no 取自 Excel 的 Order Number 列。
                order_no = self._excel_str(order_row.get('Order Number'))
                sap_amount_vat = ''

                def _report_step(step_name, step_result):
                    """步骤结果落到 textBrowser；失败时红字附带错误原因。"""
                    if step_result.success:
                        suffix = ': %s' % step_result.message if step_result.message else ''
                        self.textBrowser.append('%s 成功%s' % (step_name, suffix))
                    else:
                        message = step_result.message or '未知错误'
                        self.textBrowser.append(
                            "<font color='red'>%s 失败: %s</font>" % (step_name, message)
                        )
                    QApplication.processEvents()

                # Step 1: VA01 创建订单头
                va01_done = False
                if flow_options.get('va01Check'):
                    create_result = service.create_order(order, revenue)
                    remarks.append(f"VA01:{create_result.message}" if create_result.message else "VA01")
                    order_no = create_result.order_no or order_no
                    sap_amount_vat = create_result.sap_amount_vat or sap_amount_vat
                    va01_done = create_result.success
                    _report_step('VA01', create_result)

                # Step 2: Save VA01 —— VA01 成功后，若有后续步骤或显式 saveCheck 都需要落盘
                need_save_va01 = va01_done and (
                    flow_options.get('saveCheck')
                    or flow_options.get('va02Check')
                )
                if need_save_va01:
                    save_va01_result = service.save('VA01')
                    remarks.append(
                        f"Save VA01:{save_va01_result.message}" if save_va01_result.message else "Save VA01"
                    )
                    if save_va01_result.success:
                        saved_order_no = self._extract_order_no(sap_session)
                        if saved_order_no:
                            order_no = saved_order_no
                    _report_step('Save VA01', save_va01_result)

                # Step 3-6: VA02 段。只要勾选了 va02Check / labCostCheck / planCostCheck 任意一项，就需要进入 VA02。
                need_va02 = (flow_options.get('va02Check'))

                if need_va02:
                    open_result = service.open_order(order_no)
                    # 直接从 VA02 开始时，Excel 'Order Number' 可能为空；优先取 open_result，再兜底从 SAP 提取。
                    order_no = open_result.order_no or order_no or self._extract_order_no(sap_session)
                    remarks.append(f"VA02:{open_result.message}" if open_result.message else "VA02")
                    _report_step('Open VA02', open_result)
                    if order_no:
                        # 立即在显示框反馈识别到的订单号，便于直接开始 VA02 场景的用户确认。
                        self.textBrowser.append("识别到 Order No.: %s" % order_no)
                        QApplication.processEvents()

                    if open_result.success:
                        # add item 仅在 va02Check 时进行；纯 Data B / Plan Cost 场景不重复加 item。
                        if flow_options.get('va02Check'):
                            item_result = service.add_items(order, revenue)
                            order_no = item_result.order_no or order_no
                            remarks.append(f"Item:{item_result.message}" if item_result.message else "Item")
                            sap_amount_vat = item_result.sap_amount_vat or sap_amount_vat
                            _report_step('Add Item', item_result)

                        if flow_options.get('labCostCheck') and data_b_entries:
                            # items 加和换算 CNY，传给 service 用于判断是否回填订单价值字段。
                            items_revenue_total_cny = items_revenue_total * (order.exchange_rate or 1.0)
                            data_b_result = service.fill_lab_cost_entries(
                                data_b_entries,
                                auftragswert_cny=items_revenue_total_cny,
                            )
                            remarks.append(
                                f"Data B:{data_b_result.message}" if data_b_result.message else "Data B"
                            )
                            _report_step('Data B', data_b_result)

                        if flow_options.get('planCostCheck'):
                            for item_no, plan_cost_entries in plan_cost_entries_by_item.items():
                                focus_row = self._find_item_row(order, item_no)
                                plan_result = service.apply_plan_cost_entries(
                                    plan_cost_entries, focus_row=focus_row
                                )
                                remarks.append(
                                    f"Plan Cost {item_no}:{plan_result.message}"
                                    if plan_result.message
                                    else f"Plan Cost {item_no}"
                                )
                                _report_step('Plan Cost %s' % item_no, plan_result)

                        if flow_options.get('saveCheck'):
                            save_va02_result = service.save('VA02')
                            remarks.append(
                                f"Save VA02:{save_va02_result.message}"
                                if save_va02_result.message
                                else "Save VA02"
                            )
                            _report_step('Save VA02', save_va02_result)

                    # VA02 段结束后再做一次最终兜底，覆盖中间步骤未回传 order_no 的边界情况。
                    if not order_no:
                        order_no = self._extract_order_no(sap_session)

                # SAP 加和金额本身已经含税，理论上应等于 Excel "Tax-inclusive amount"。
                # 容差 0.01 容忍浮点误差；只有在 Excel 含税金额可用时才比较，避免空值误判。
                try:
                    sap_amount_value = float(str(sap_amount_vat).replace(',', '')) if sap_amount_vat else 0.0
                except (TypeError, ValueError):
                    sap_amount_value = 0.0
                try:
                    excel_amount_value = float(str(excel_amount_vat).replace(',', '')) if excel_amount_vat else 0.0
                except (TypeError, ValueError):
                    excel_amount_value = 0.0

                amount_diff = round(sap_amount_value - excel_amount_value, 2)
                amount_mismatch = excel_amount_value > 0 and abs(amount_diff) >= 0.01
                if amount_mismatch:
                    diff_msg = (
                        f"含税金额不一致: Excel={format(excel_amount_value, ',.2f')} "
                        f"SAP={format(sap_amount_value, ',.2f')} "
                        f"差额={format(amount_diff, ',.2f')}"
                    )
                    remarks.append(diff_msg)

                log_file.loc[index, 'Order No.'] = order_no
                log_file.loc[index, 'Remark'] = ';'.join([item for item in remarks if item])
                log_file.loc[index, 'Proforma No.'] = ''
                log_file.loc[index, 'sapAmountVat'] = sap_amount_vat
                log_file.loc[index, 'Update Time'] = datetime.datetime.today().strftime('%Y-%m-%d %H:%M:%S')
                log_file.to_excel(log_data_path, merge_cells=False, index=False)

                # 订单结束摘要：order no + SAP 含税金额；与 Excel 含税金额一致性提示。
                self.textBrowser.append("Order No.: %s" % order_no)
                self.textBrowser.append("SAP 金额(加和,含税): %s" % (sap_amount_vat or '--'))
                if amount_mismatch:
                    self.textBrowser.append("<font color='red'>%s</font>" % diff_msg)
                elif excel_amount_value > 0:
                    self.textBrowser.append("含税金额一致(Excel == SAP)")
                self.textBrowser.append('----------------------------------')
                QApplication.processEvents()

                if index < len(order_df) - 1 and self.checkBox_5.isChecked():
                    reply = QMessageBox.question(
                        self,
                        '信息',
                        '是否继续填写下一个Order',
                        QMessageBox.Yes | QMessageBox.No,
                        QMessageBox.Yes,
                    )
                    if reply != QMessageBox.Yes:
                        break

            self.lineEdit_9.setText(log_data_path)
            self.textBrowser.append("订单数据已处理完成")
            self.textBrowser.append("log数据:%s" % log_data_path)
            self.textBrowser.append('----------------------------------')
            QMessageBox.information(self, "提示信息", "订单数据已处理完成", QMessageBox.Yes)
        except Exception as msg:
            self.textBrowser.append('订单数据处理失败:%s' % msg)
            self.textBrowser.append('----------------------------------')
            QMessageBox.information(self, "提示信息", '订单数据处理失败:%s' % msg, QMessageBox.Yes)
        finally:
            if sap_session is not None:
                sap_session.close()

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
                        long_text=_to_str(raw_item.get('long_text', raw_item.get('Item Group Description'))),
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

        # 使用类级别的 _extract_order_no 共享实现。
        _extract_order_no = SapOrderMixin._extract_order_no

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
                sales_group=str(guiData.get('salesGroup', '')).strip(),
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
                currency_type=str(guiData.get('currencyType', '')).strip(),
                exchange_rate=_to_float(guiData.get('exchangeRate'), 1.0),
                short_text=str(guiData.get('shortText', '')).strip(),
                global_partner_code=str(guiData.get('globalPartnerCode', '')).strip(),
                sales_name=str(guiData.get('salesName', '')).strip(),
                ecd=time.strftime("%Y.%m.%d"),
                order_center=str(guiData.get('orderCenter', '')).strip(),
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
        """兼容旧调用入口；当前业务已改为通过 odmDataToSap 从多 sheet Excel 创建订单。"""
        return {
            'Remark': '当前业务已改为从 SAP_data 多 sheet 文件创建订单，请使用 odmDataToSap。',
            'orderNo': '',
            'Proforma No.': '',
            'sapAmountVat': '',
        }
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
            self.textBrowser.append('没有文件，请添加')
            QApplication.processEvents()


