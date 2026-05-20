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
from PyQt5.QtCore import QDate, QSignalBlocker
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
from sap import (
    DataBEntry,
    OrderData,
    OrderItemData,
    OrderService,
    PartnerOptions,
    PlanCostEntry,
    RevenueData,
    SapConfig,
    SapResult,
    SapSession,
)
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

    @staticmethod
    def _excel_date_dot(value, default=''):
        """读取 Excel 单元格日期值，统一格式化为 SAP 接受的 'YYYY.MM.DD'。

        - pd.Timestamp / datetime / 各种字符串日期格式 → '2026.05.15'
        - 空值 / NaT / 无法解析的字符串 → default（默认空串），避免污染 SAP
        """
        if pd.isna(value) or value == '':
            return default
        ts = pd.to_datetime(value, errors='coerce')
        if pd.isna(ts):
            return default
        return ts.strftime('%Y.%m.%d')

    def _filter_related_rows(self, dataframe, order_row):
        """按 Combine Id 严格筛选当前订单对应的明细行；调用前需确保 Combine Id 存在。"""
        if dataframe.empty:
            return dataframe
        combine_id = self._excel_value(order_row.get('Combine Id'))
        return dataframe[dataframe['Combine Id'] == combine_id]

    def _apply_order_row_to_gui(self, order_row):
        """将订单行 Excel 数据回填到主界面控件，便于用户实时跟踪当前订单。

        取值口径与 _build_order_from_dataframes / _build_sap_config_from_order_row 保持一致：
          - 未税金额优先 'Revenue'，兜底 'Untaxed amount'（与 _build_revenue_from_order_row 对齐）
          - 汇率默认 1.0
        comboBox 用 QSignalBlocker 包裹，避免触发任何已绑定（或未来误绑）的信号槽。
        Excel 中的 CS/Sales 若不在 configContent 中，setCurrentText 会静默保留原值，给出黄字提示但不阻断流程。
        """
        sap_no = self._excel_str(order_row.get('SAP Customer Code'))
        project_no = self._excel_str(order_row.get('Request Number'))
        currency_type = self._excel_str(order_row.get('Currency'))
        exchange_rate = self._excel_float(order_row.get('Rate'), 1.0)
        global_partner_code = self._excel_str(order_row.get('GPC Code'))
        cs_name = self._excel_str(order_row.get('Primary CS'))
        sales_name = self._excel_str(order_row.get('Sales'))
        amount = self._excel_float(
            order_row.get('Revenue'),
            self._excel_float(order_row.get('Untaxed amount')),
        )

        self.lineEdit.setText(sap_no)
        self.lineEdit_2.setText(project_no)
        self.lineEdit_3.setText(global_partner_code)
        self.doubleSpinBox.setValue(exchange_rate)
        self.doubleSpinBox_2.setValue(amount)

        with QSignalBlocker(self.comboBox):
            self.comboBox.setCurrentText(currency_type)
        with QSignalBlocker(self.comboBox_2):
            self.comboBox_2.setCurrentText(cs_name)
        with QSignalBlocker(self.comboBox_3):
            self.comboBox_3.setCurrentText(sales_name)

        if cs_name and cs_name not in configContent:
            self.textBrowser.append(
                "<font color='orange'>CS [%s] 不在配置文件中，csCode 将为空</font>" % cs_name
            )
        if sales_name and sales_name not in configContent:
            self.textBrowser.append(
                "<font color='orange'>Sales [%s] 不在配置文件中，salesCode 将为空</font>" % sales_name
            )

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
        """从订单头和 item 表构建 SAP 订单对象。

        items 列表会按 item 号数字升序做稳定排序，让 list 索引语义对齐 SAP VA02
        item 概览页回车后的物理 row 顺序（SAP 按 POSNR 升序自动重排）。
        空 / 非数字 item 保持 Excel 相对顺序排到末尾，对应 SAP 自动分配新号的行。
        排序后的不变量"order.items 索引 = SAP 物理 row"贯穿 _write_item_rows、
        _find_item_row 和 plan cost 循环，消除上游 Excel 顺序与 SAP 行号错位的隐患。
        """
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

        items = self._sort_items_for_sap(items)

        return OrderData(
            sap_no=self._excel_str(order_row.get('SAP Customer Code')),
            project_no=self._excel_str(order_row.get('Request Number')),
            amount_vat=self._excel_str(order_row.get('Tax-inclusive amount')),
            currency_type=self._excel_str(order_row.get('Currency')),
            exchange_rate=self._excel_float(order_row.get('Rate'), 1.0),
            short_text=self._excel_str(order_row.get('售达方的文本')),
            global_partner_code=self._excel_str(order_row.get('GPC Code')),
            sales_name=self._excel_str(order_row.get('Sales')),
            sales_group=self._excel_str(order_row.get('Cost Center'))[-3:],
            ecd=self._excel_date_dot(order_row.get('Ecd')),
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
        """从 sub 表构建 Data B 和 Plan Cost 的直接写入明细。

        Returns:
            tuple[list[DataBEntry], dict[str, list[PlanCostEntry]]]:
              - Data B: 按 sub 表行级保留（每条 sub 行一条 DataBEntry）。
              - Plan Cost: 按 item 分组，每 item 一组 PlanCostEntry 列表。

        Plan Cost 聚合规则：
          - FREMDL：按 (item, cost_center) 聚合 Σ Sub-Cost RMB；
            sub 行 Sub Site Cost Center 为空时使用订单 Cost Center 兜底。
          - T01AST：按 item 聚合 Σ Sub Site Plan Hour；cost_center = 订单 Cost Center。
        """
        order_sub_df = self._filter_related_rows(sub_df, order_row)
        order_cost_center = self._excel_str(order_row.get('Cost Center'))

        data_b_entries: list[DataBEntry] = []
        # FREMDL 维度：(item_no, cost_center) → Σ Sub-Cost RMB
        fremdl_summary: dict[tuple[str, str], float] = {}
        # T01AST 维度：item_no → Σ Sub Site Plan Hour
        t01ast_summary: dict[str, float] = {}

        for _, sub_row in order_sub_df.iterrows():
            item_no = self._excel_str(sub_row.get('item'))
            raw_sub_cc = self._excel_str(sub_row.get('Sub Site Cost Center'))
            sub_cost = self._excel_float(sub_row.get('Sub-Cost RMB'))
            plan_hour = self._excel_float(sub_row.get('Sub Site Plan Hour'))

            # Data B 行级写入：保留每条 sub 行的明细。
            if raw_sub_cc and sub_cost:
                data_b_entries.append(DataBEntry(
                    performer_cost_center=raw_sub_cc,
                    rate_cost_center=raw_sub_cc,
                    amount=sub_cost,
                    item=item_no,
                ))

            # Plan Cost FREMDL：cost_center 缺失时兜底为订单 Cost Center。
            if sub_cost:
                fremdl_cc = raw_sub_cc or order_cost_center
                if fremdl_cc:
                    key = (item_no, fremdl_cc)
                    fremdl_summary[key] = fremdl_summary.get(key, 0.0) + sub_cost

            # Plan Cost T01AST：按 item 累计工时，cost_center 必须用订单 Cost Center。
            if plan_hour and order_cost_center:
                t01ast_summary[item_no] = t01ast_summary.get(item_no, 0.0) + plan_hour

        plan_cost_entries_by_item: dict[str, list[PlanCostEntry]] = {}
        for (item_no, cost_center), amount in fremdl_summary.items():
            if amount:
                plan_cost_entries_by_item.setdefault(item_no, []).append(PlanCostEntry(
                    cost_center=cost_center,
                    category='FREMDL',
                    amount=amount,
                ))
        for item_no, amount in t01ast_summary.items():
            if amount:
                plan_cost_entries_by_item.setdefault(item_no, []).append(PlanCostEntry(
                    cost_center=order_cost_center,
                    category='T01AST',
                    amount=amount,
                ))

        return data_b_entries, plan_cost_entries_by_item

    @staticmethod
    def _sort_items_for_sap(items):
        """按 SAP VA02 item 概览页 POSNR 升序的物理 row 顺序对 items 排序。

        SAP 在写完 item 号并按回车后会自动按 POSNR 升序重排；本方法在适配层
        提前完成等效排序，让 order.items 列表顺序 == SAP 写入后的物理 row 顺序。

        排序规则：
            - 有效数字 item → 按数字升序（key=(0, int)）
            - 空 / 非数字 item → 保持原相对顺序，整体排到末尾（key=(1, idx)）

        排序键采用 (bucket, secondary) 元组而非 float('inf')，保证 sorted 的稳定
        性同时对非数字 item 维持 Excel 原相对顺序。
        """
        def _key(indexed):
            idx, item = indexed
            raw = (item.item or '').strip()
            if raw.isdigit():
                return (0, int(raw), idx)
            return (1, 0, idx)

        return [item for _, item in sorted(enumerate(items), key=_key)]

    @staticmethod
    def _find_item_row(order, item_no):
        """根据 item 编号定位 SAP item 表格中的物理行号。

        order.items 已由 _sort_items_for_sap 在适配层排好序，list 索引即 SAP 物理 row；
        本方法直接 enumerate 查找即可，无需再次排序。找不到时返回 0 作为兜底。
        """
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

            # 列顺序优化：核心追踪字段（业务键、订单号、备注）置顶，便于人工查阅 log。
            # 优先列中缺失的列自动跳过，避免 order_df 字段命名变更时直接报错。
            priority_cols = ['Combine Id', 'Request Number', 'Order No.', 'Remark']
            existing_priority = [col for col in priority_cols if col in log_file.columns]
            other_cols = [col for col in log_file.columns if col not in existing_priority]
            log_file = log_file[existing_priority + other_cols]

            sap_session = SapSession.connect()

            for index, order_row in order_df.iterrows():
                # Combine Id 是关联 item / sub 的唯一键，缺失直接跳过当前订单。
                if pd.isna(order_row.get('Combine Id')):
                    log_file.loc[index, 'Remark'] = '缺失 Combine Id，无法关联 item/sub'
                    log_file.to_excel(log_data_path, merge_cells=False, index=False)
                    continue

                # 将当前订单关键字段回填到 GUI 控件，便于用户实时跟踪正在处理的订单。
                self._apply_order_row_to_gui(order_row)

                # 三张 DataFrame 已包含完整业务数据，这里只做对象适配和 SAP 写入。
                order = self._build_order_from_dataframes(order_row, item_df)
                revenue = self._build_revenue_from_order_row(order_row)
                config = self._build_sap_config_from_order_row(order_row)
                data_b_entries, plan_cost_entries_by_item = self._build_sub_entries_from_dataframe(order_row, sub_df)
                service = OrderService(sap_session, config)

                # 按本次勾选的步骤分级校验：缺啥提示啥，支持只跑 Data B / Plan Cost 的分批验证场景。
                need_va01_check = flow_options.get('va01Check')
                need_va02_items_check = flow_options.get('va02Check')
                need_data_b_check = flow_options.get('labCostCheck')
                need_plan_cost_check = flow_options.get('planCostCheck')

                # 订单号：优先 Excel；仅首行 + 单跑 Data B/Plan Cost 场景允许从 SAP 当前会话兜底读取
                # （用户已手动打开 VA02 页面的情况）。后续行漏填则直接缺失报错，避免误写同一订单。
                order_no = self._excel_str(order_row.get('Order Number'))
                if (
                    not order_no
                    and index == 0
                    and (need_data_b_check or need_plan_cost_check)
                    and not need_va01_check
                ):
                    try:
                        order_no = self._extract_order_no(sap_session)
                    except Exception:
                        order_no = ''

                missing_fields = []
                if need_va01_check and (not order.sap_no or not order.project_no):
                    missing_fields.append('SAP No./Project No.')
                if (need_va01_check or need_va02_items_check) and not order.items:
                    missing_fields.append('items')
                if (need_data_b_check or need_plan_cost_check) and not need_va01_check and not order_no:
                    missing_fields.append('Order Number')

                if missing_fields:
                    missing_msg = '关键订单信息缺失（%s）' % '/'.join(missing_fields)
                    log_file.loc[index, 'Remark'] = missing_msg
                    log_file.to_excel(log_data_path, merge_cells=False, index=False)
                    self.textBrowser.append(
                        "<font color='red'>No.%s %s</font>" % (index + 1, missing_msg)
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
                # 所有步骤独立可选；跳过 VA01 时 order_no 取自 Excel 的 Order Number 列（已在校验前读取）。
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
                    if save_va01_result.success:
                        saved_order_no = self._extract_order_no(sap_session)
                        if saved_order_no:
                            order_no = saved_order_no
                        else:
                            # SAP 保存命令未抛错但读不到订单号（业务级静默失败）→ 视为 VA01 段失败，
                            # 否则下游 VA02/Data B/Plan Cost 会用空/残留订单号继续执行。
                            va01_done = False
                            save_va01_result = SapResult.fail(
                                "Save VA01 后未能读取到 Order No.", step="save"
                            )
                    else:
                        # Save VA01 显式失败 → 视为 VA01 段失败，由 va01_blocked 守卫拦截后续步骤。
                        va01_done = False
                    remarks.append(
                        f"Save VA01:{save_va01_result.message}" if save_va01_result.message else "Save VA01"
                    )
                    _report_step('Save VA01', save_va01_result)

                # Step 3-6: VA02 段。只要勾选了 va02Check / labCostCheck / planCostCheck 任意一项，就需要进入 VA02。
                # 当 VA01 被勾选但失败（va01_blocked=True）时短路 VA02 段，
                # 避免 SAP VA02 窗体残留上一个订单号导致 Add Item / Data B / Plan Cost 误写入上一单。
                has_va02_step = need_va02_items_check or need_data_b_check or need_plan_cost_check
                va01_blocked = bool(flow_options.get('va01Check')) and not va01_done
                need_va02 = not va01_blocked and has_va02_step

                # VA01 失败导致 VA02 段被跳过时给出红字提示，避免用户以为流程在静默运行。
                if va01_blocked and has_va02_step:
                    self.textBrowser.append(
                        "<font color='red'>VA01 失败，跳过当前订单的 VA02/Data B/Plan Cost 步骤</font>"
                    )
                    QApplication.processEvents()

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
                                order,
                                auftragswert_cny=items_revenue_total_cny,
                            )
                            remarks.append(
                                f"Data B:{data_b_result.message}" if data_b_result.message else "Data B"
                            )
                            _report_step('Data B', data_b_result)

                        if flow_options.get('planCostCheck'):
                            # 按 order.items 顺序（即 SAP 物理 row 顺序）调度 plan cost，
                            # 避免字典插入顺序（来自 sub 表出现顺序）与 SAP 行号顺序不一致；
                            # sub 表未提供 plan cost 数据的 item 直接跳过。
                            for row, item in enumerate(order.items):
                                plan_cost_entries = plan_cost_entries_by_item.get(item.item)
                                if not plan_cost_entries:
                                    continue
                                plan_result = service.apply_plan_cost_entries(
                                    plan_cost_entries, focus_row=row
                                )
                                remarks.append(
                                    f"Plan Cost {item.item}:{plan_result.message}"
                                    if plan_result.message
                                    else f"Plan Cost {item.item}"
                                )
                                _report_step('Plan Cost %s' % item.item, plan_result)

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

    def orderUnlockOrLock(self, flag):
        """批量锁定/解锁订单。

        从 Excel 文件读取 Order No 列，通过新 SAP 服务层（OrderService）
        逐单执行 lock / unlock；全程复用单个 SapSession，循环结束统一关闭。
        """
        fileUrl = self.lineEdit_6.text()
        if not fileUrl:
            self.textBrowser.append('没有文件，请添加')
            QApplication.processEvents()
            return

        filepath, _ = os.path.split(fileUrl)
        log_file_name = 'log %s.xlsx' % time.strftime('%Y-%m-%d %H.%M.%S')
        Log_file = '%s\\%s' % (filepath, log_file_name)
        log_obj = Logger(Log_file, ['Update', 'Order No', 'Type', 'Remark'])
        order_list = list(Get_Data().getFileData(fileUrl)['Order No'])

        # 锁/解锁不依赖任何 SapConfig 字段，仅为满足 OrderService 构造签名提供空实例。
        empty_config = SapConfig(
            order_type='',
            sales_organization='',
            distribution_channels='',
            sales_office='',
            sales_group='',
            sub_cost_center_cs='',
            sub_cost_center_chm='',
            sub_cost_center_phy='',
            cs_code='',
            sales_code='',
        )

        sap_session = None
        try:
            sap_session = SapSession.connect()
            service = OrderService(sap_session, empty_config)

            for i, orderNo in enumerate(order_list, start=1):
                log_list = {'Order No': orderNo, 'Type': flag, 'Remark': ''}
                try:
                    result = service.unlock(orderNo) if flag == 'Unlock' else service.lock(orderNo)
                    self.textBrowser.append('%s.Order No: %s' % (i, orderNo))
                    self.textBrowser.append('%s' % (result.message or ''))
                    QApplication.processEvents()
                    if not result.success:
                        log_list['Remark'] = result.message or ''
                except Exception as exc:
                    self.textBrowser.append(
                        "<font color='red'>该Order: %s 有问题: %s</font>" % (orderNo, exc)
                    )
                    log_list['Remark'] = str(exc)
                    QApplication.processEvents()
                log_obj.log(log_list)

            log_obj.save_log_to_excel()
            self.textBrowser.append('%s' % Log_file)
            QApplication.processEvents()
            os.startfile(Log_file)
        finally:
            if sap_session is not None:
                sap_session.close()


