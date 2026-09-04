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
    ItemAddInfo,
    OrderData,
    OrderEditService,
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

    @staticmethod
    def _is_date_before_today(date_text: str) -> bool:
        """判断 YYYY.MM.DD 日期是否早于今天；空值或非法日期不在这里拦截。"""
        if not date_text:
            return False
        ts = pd.to_datetime(date_text, errors='coerce')
        if pd.isna(ts):
            return False
        return ts.date() < pd.Timestamp.today().date()

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
        Sales 若不在 configContent 中，setCurrentText 会静默保留原值，给出黄字提示但不阻断流程
        （Sales 非必填）；CS 取不到编号的订单在调用本方法之前就已被拦截跳过，故此处不再提示。
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

        if sales_name and sales_name not in configContent:
            self.textBrowser.append(
                "<font color='orange'>Sales [%s] 不在配置文件中，salesCode 将为空</font>" % sales_name
            )

    def _resolve_cs_code(self, order_row):
        """解析 Primary CS 对应的 config 人员编号；CS 为空或未录入配置时返回空串。

        SAP 伙伴页的"负责雇员"写的是该编号，取不到即无法完成订单，故也是订单必填校验的口径。
        """
        cs_name = self._excel_str(order_row.get('Primary CS'))
        if not cs_name:
            return ''
        return self._excel_str(configContent.get(cs_name, ''))

    def _build_sap_config_from_order_row(self, order_row):
        """按当前订单行和系统配置构建 SAP 固定参数。"""
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
            cs_code=self._resolve_cs_code(order_row),
            sales_code=configContent.get(sales_name, ''),
            # 客户选择是否为海外订单
            data_ae1=self.lineEdit_21.text().split(';'),
            data_az2=self.lineEdit_22.text().split(';'),
            # TUV IC 订单清单：无 GUI 控件，直接读 config（命中则 VA01 写 IC_TRANSAKTION=O1）。
            data_b_tuv=str(configContent.get('Data_B_TUV', '')).split(';'),
        )

    def _resolve_sales_group(self, order_row) -> str:
        """解析 VA01 销售组，避免把 Cost Center 的非数字尾缀误写入 VKGRP。"""
        configured_sales_group = self.lineEdit_14.text().strip()
        cost_center = self._excel_str(order_row.get('Cost Center'))
        derived_sales_group = cost_center[-3:] if len(cost_center) >= 3 else ''

        if derived_sales_group.isdigit():
            return derived_sales_group

        if derived_sales_group:
            self.textBrowser.append(
                "<font color='orange'>Cost Center [%s] 不能解析出销售组，使用配置 Sales Group [%s]</font>"
                % (cost_center, configured_sales_group)
            )
            QApplication.processEvents()
        return configured_sales_group

    def _build_order_from_dataframes(self, order_row, item_df):
        """从订单头和 item 表构建 SAP 订单对象。

        items 列表会按 item 号数字升序做稳定排序，让写入顺序贴近 SAP VA02 item 概览页
        回车后的物理 row 顺序（SAP 按 POSNR 升序自动重排），减少重排幅度。
        空 / 非数字 item 保持 Excel 相对顺序排到末尾，对应 SAP 自动分配新号的行。

        注意：这里**只是对齐，不构成正确性依赖**。SAP 侧的行号一律由 `read_item_rows` /
        `find_item_row` 实时重读定位（见 order.py 同名方法），空 item 号由 SAP 自动分配、
        与 Excel 顺序无关，故"列表索引 == 物理 row"从来不是可以依赖的不变量。
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
            product_sub_category=self._excel_str(order_row.get('Product Sub-Category')),
            global_partner_code=self._excel_str(order_row.get('GPC Code')),
            sales_name=self._excel_str(order_row.get('Sales')),
            sales_group=self._resolve_sales_group(order_row),
            ecd=self._excel_date_dot(order_row.get('Ecd')),
            order_cost_center=self._excel_str(order_row.get('Order Center')),
            items=items,
        )

    def _build_revenue_from_order_row(self, order_row):
        """从订单表读取 Revenue；只做对象适配，不再重新分配或计算业务金额。

        'Revenue' 列本身已是 CNY 换算值（= Untaxed amount × Rate），故 revenue_cny 直接取该值，
        切勿再 × Rate（历史双重汇率 bug 根源）。订单价值改由 SAP item 净值加和 × 汇率单独计算。
        """
        revenue = self._excel_float(order_row.get('Revenue'), self._excel_float(order_row.get('Untaxed amount')))
        return RevenueData(
            revenue=revenue,
            revenue_cny=revenue,
        )

    @staticmethod
    def _forced_data_b_cost_centers():
        """解析 config `Data_B_Cost_Center`（`;` 分隔）为强制录入的成本中心列表。

        返回值保序去重，已剔除空白项；配置缺失或全空时返回空列表（等价于关闭该功能）。
        """
        raw = str(configContent.get('Data_B_Cost_Center', '') or '')
        result = []
        for cc in raw.split(';'):
            cc = cc.strip()
            if cc and cc not in result:
                result.append(cc)
        return result

    def _append_forced_data_b_entries(self, data_b_entries):
        """在表格行之后追加 config 强制成本中心行（原地修改 data_b_entries）。

        去重口径：表格行已含相同执行部门成本中心则跳过，避免 SAP 出现重复行。
        强制行只写执行部门列，故 rate/amount/item 留空，由 kostl_only 标记下游写入范围。
        """
        existing = {
            (entry.performer_cost_center or '').strip()
            for entry in data_b_entries
        }
        for cost_center in self._forced_data_b_cost_centers():
            if cost_center in existing:
                continue
            existing.add(cost_center)
            data_b_entries.append(DataBEntry(
                performer_cost_center=cost_center,
                rate_cost_center='',
                amount=0.0,
                item='',
                kostl_only=True,
            ))

    def _build_sub_entries_from_dataframe(self, order_row, sub_df):
        """从 sub 表构建 Data B 和 Plan Cost 的直接写入明细。

        Returns:
            tuple[list[DataBEntry], dict[str, list[PlanCostEntry]]]:
              - Data B: 按 sub 表行级保留（每条 sub 行一条 DataBEntry），
                末尾追加 config `Data_B_Cost_Center` 的强制成本中心行。
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

        # 强制成本中心行恒排在表格行之后，保证「entries 索引 = SAP 物理 row」不变量。
        self._append_forced_data_b_entries(data_b_entries)

        return data_b_entries, plan_cost_entries_by_item

    def _append_step_result(self, step_name, step_result):
        """把单个 SAP 步骤结果渲染到 textBrowser，按严重度区分颜色（创建/编辑两流程共用）。

        优先级（高→低）：
            - success=False → 红色「失败」；
            - 消息含"读取失败"（控件读不到，疑似控件 ID bug）→ 红色「异常」；
            - warning 标记 或 消息含"已跳过"（如 SAP 无对应 item、SAP 有/Excel 无）→ 橙色「警告」；
            - 其余 → 默认色「成功」。
        """
        message = step_result.message or ''
        if not step_result.success:
            self.textBrowser.append(
                "<font color='red'>%s 失败: %s</font>" % (step_name, message or '未知错误')
            )
        elif '读取失败' in message:
            self.textBrowser.append(
                "<font color='red'>%s 异常: %s</font>" % (step_name, message)
            )
        elif step_result.warning or '已跳过' in message:
            self.textBrowser.append(
                "<font color='orange'>%s 警告: %s</font>" % (step_name, message)
            )
        else:
            suffix = ': %s' % message if message else ''
            self.textBrowser.append('%s 成功%s' % (step_name, suffix))
        QApplication.processEvents()

    def _edit_order_row(
        self,
        index,
        order_row,
        order,
        config,
        data_b_entries,
        plan_cost_entries_by_item,
        order_no,
        flow_options,
        sap_session,
        log_file,
        log_data_path,
    ):
        """编辑分支：进 VA02 对比 Excel 与 SAP 现值，仅更新差异。

        编辑范围由流程控制按钮独立门控，四步互不影响：va01Check→编辑 header；
        va02Check→编辑 item；planCostCheck→编辑 Plan Cost；labCostCheck→编辑 Data B。
        未勾选的步骤直接跳过。全程收集差异写入 log Remark，操作类型标记 Edit。
        """
        _report_step = self._append_step_result

        combine_id = self._excel_str(order_row.get('Combine Id'))
        primary_cs = self._excel_str(order_row.get('Primary CS'))
        sales_name = self._excel_str(order_row.get('Sales'))
        excel_amount_untaxed = self._excel_str(order_row.get('Untaxed amount'))

        self.textBrowser.append('========== No.%s [编辑] ==========' % (index + 1))
        self.textBrowser.append("Combine Id: %s" % combine_id)
        self.textBrowser.append("Order No.: %s" % order_no)
        self.textBrowser.append("Request Number: %s" % order.project_no)
        self.textBrowser.append("Primary CS: %s / Sales: %s" % (primary_cs, sales_name))
        QApplication.processEvents()

        remarks = []
        sap_amount_vat = ''
        service = OrderEditService(sap_session, config)

        # 打开 VA02。
        open_result = service.open_order(order_no)
        order_no = open_result.order_no or order_no
        _report_step('Open VA02', open_result)
        if not open_result.success:
            remarks.append(f"Open VA02:{open_result.message}")
            self._write_edit_log(
                log_file, log_data_path, index, order_no, remarks, sap_amount_vat
            )
            return

        # Header 编辑（va01Check）：未勾选 VA01 则跳过，与其余三步同款独立门控。
        if flow_options.get('va01Check'):
            header_diffs: list[str] = []
            header_result = service.edit_header(order, header_diffs)
            self._append_remark(remarks, "Header", header_result, header_diffs)
            _report_step('Header 编辑', header_result)

        # Item 编辑（va02Check）：收集新增明细(added)，供落库后建立 ODM→SAP 号映射。
        item_no_map: dict[str, str] = {}
        added: list[ItemAddInfo] = []
        item_ok = True
        if flow_options.get('va02Check'):
            item_diffs: list[str] = []
            item_result = service.edit_items(order, item_diffs, added_out=added)
            item_ok = item_result.success
            sap_amount_vat = item_result.sap_amount_vat or sap_amount_vat
            self._append_remark(remarks, "Item", item_result, item_diffs)
            _report_step('Item 编辑', item_result)

        # 有 item 新增 → 先重读概览建立 ODM→SAP 号映射，供 Plan Cost 定位（不 save）。
        if item_ok and added:
            item_no_map = service.build_item_no_mapping(added)

        # Plan Cost 编辑（planCostCheck）：按（映射后）item 号在概览页实时定位物理行——
        # 编号可能与 ODM 不同、SAP 也可能重排，故每次重读匹配，绝不用写入时的行号。
        if flow_options.get('planCostCheck'):
            for item in order.items:
                plan_cost_entries = plan_cost_entries_by_item.get(item.item)
                if not plan_cost_entries:
                    continue
                pc_diffs: list[str] = []
                target_item = item_no_map.get(item.item, item.item)
                plan_result = service.edit_plan_cost(plan_cost_entries, pc_diffs, target_item=target_item)
                self._append_remark(remarks, f"Plan Cost {item.item}", plan_result, pc_diffs)
                _report_step('Plan Cost %s' % item.item, plan_result)

        # Data B 同步门控：勾选 labCostCheck 即执行（labCostCheck→编辑 Data B），与 Excel 是否
        # 有 Data B 行无关——Excel 整理后无某条须删对应 SAP 行，Data B 全空须删 SAP 全部行。
        # 旧版误用 `and data_b_entries` 短路，导致 Excel 清空时同步整段不被调用、删不掉原有行。
        data_b_enabled = bool(flow_options.get('labCostCheck'))

        # Data B 第一段（clear）：与 Excel 比对，一致则整段跳过（changed=False），有差异才两阶段删空。
        # 删空与重建之间必须隔一次 save + open_order：成本表行是对执行部门表成本中心的引用，
        # 同屏内 SAP 不认未落库的新增行，删后直接写必报 ZR520（见 clear_data_b 文档）。
        data_b_changed = False
        if data_b_enabled:
            clear_diffs: list[str] = []
            clear_result = service.clear_data_b(
                data_b_entries, order, clear_diffs, item_no_map=item_no_map,
            )
            data_b_changed = clear_result.changed
            self._append_remark(remarks, "Data B 清空", clear_result, clear_diffs)
            _report_step('Data B 清空', clear_result)
            if not clear_result.success:
                self._write_edit_log(
                    log_file, log_data_path, index, order_no, remarks, sap_amount_vat
                )
                return

        # Data B 前保存：需要重建且确有行要写时——删除须先落盘（否则成本表引用报错），
        # 且改号 item 的真实号仅落盘后才由 SAP 分配，保存 + 重开后重建映射拿到真实号供 POSNR 使用。
        # 纯删除（Excel Data B 为空）不写成本表，无需前保存；比对一致时更是零额外开销。
        if data_b_enabled and data_b_entries and (data_b_changed or (item_ok and added)):
            save_before_db = service.save('VA02 Edit - Before Data B')
            _report_step('Data B 前保存', save_before_db)
            if not save_before_db.success:
                remarks.append(f"Data B 前保存:{save_before_db.message}")
                self._write_edit_log(
                    log_file, log_data_path, index, order_no, remarks, sap_amount_vat
                )
                return
            reopen = service.open_order(order_no)
            order_no = reopen.order_no or order_no
            _report_step('重开订单(号映射)', reopen)
            if not reopen.success:
                remarks.append(f"重开订单:{reopen.message}")
                self._write_edit_log(
                    log_file, log_data_path, index, order_no, remarks, sap_amount_vat
                )
                return
            item_no_map = service.build_item_no_mapping(added)

        # Data B 第二段（write）：从空表重建全部行。POSNR 用前保存后重建的真实号
        # （映射为空时回退 ODM 号，绝不写空）。Excel Data B 全空时只删不写，天然跳过本段。
        # data_b_incomplete 跟踪"已清空但未成功重建"的中间态：删除已随前保存落库，
        # 若重建或最终保存失败，SAP 侧会停在 Data B 为空，必须显式告警而非只报步骤失败。
        data_b_incomplete = False
        if data_b_enabled and data_b_changed and data_b_entries:
            db_diffs: list[str] = []
            data_b_result = service.write_data_b(
                data_b_entries, order, db_diffs, item_no_map=item_no_map,
            )
            self._append_remark(remarks, "Data B", data_b_result, db_diffs)
            _report_step('Data B 重建', data_b_result)
            data_b_incomplete = not data_b_result.success

        # 订单价值(AUFTRAGSWERT)：Σ SAP item 未税净值 × 汇率，有 item 即幂等重算回填，
        # 并自愈历史双重汇率脏值。
        if order.items:
            ov_diffs: list[str] = []
            order_value_result = service.edit_order_value(order, ov_diffs)
            self._append_remark(remarks, "订单价值", order_value_result, ov_diffs)
            _report_step('订单价值编辑', order_value_result)

        # 最终保存（saveCheck）。
        if flow_options.get('saveCheck'):
            save_result = service.save('VA02 Edit')
            remarks.append(f"Save:{save_result.message}" if save_result.message else "Save")
            _report_step('Save VA02', save_result)
            if data_b_changed and data_b_entries and not save_result.success:
                data_b_incomplete = True

        # 中间态告警：Data B 的清空已落库、重建未成功 → 该订单当前 Data B 为空。
        # 全删重建是幂等的，重跑同一订单即可自动恢复，故提示重跑优先于人工补录。
        if data_b_incomplete:
            incomplete_msg = 'Data B 已清空未重建，需重跑该订单或人工补录'
            remarks.append(incomplete_msg)
            self.textBrowser.append("<font color='red'>%s</font>" % incomplete_msg)

        # 未税金额一致性校验（与创建分支同口径）：SAP 加和为未税净值，对 Excel Untaxed amount。
        try:
            sap_amount_value = float(str(sap_amount_vat).replace(',', '')) if sap_amount_vat else 0.0
        except (TypeError, ValueError):
            sap_amount_value = 0.0
        try:
            excel_amount_value = float(str(excel_amount_untaxed).replace(',', '')) if excel_amount_untaxed else 0.0
        except (TypeError, ValueError):
            excel_amount_value = 0.0
        amount_diff = round(sap_amount_value - excel_amount_value, 2)
        if excel_amount_value > 0 and abs(amount_diff) >= 0.01:
            diff_msg = (
                f"未税金额不一致: Excel={format(excel_amount_value, ',.2f')} "
                f"SAP={format(sap_amount_value, ',.2f')} 差额={format(amount_diff, ',.2f')}"
            )
            remarks.append(diff_msg)
            self.textBrowser.append("<font color='red'>%s</font>" % diff_msg)

        self._write_edit_log(log_file, log_data_path, index, order_no, remarks, sap_amount_vat)
        self.textBrowser.append("Order No.: %s 编辑完成" % order_no)
        self.textBrowser.append('----------------------------------')
        QApplication.processEvents()

    def _append_remark(self, remarks, label, result, diffs):
        """仅在"有话要说"时把该段结果写入 log Remark，无差异段不占篇幅。

        判定依据用 diffs（本段是否收集到任何差异/提示）而非 SapResult.changed：
        后者语义是"是否实际改动 SAP"、被 Data B 中途保存决策依赖，不能污染；
        而这里要的是"是否有内容值得记录"——如"SAP 有、Excel 无，已跳过"没改 SAP 但须留痕。
        失败或 warning（如 SAP 无对应 item 已跳过）也必须留痕。
        步骤执行情况仍由 _append_step_result 完整呈现在 UI 步骤面板，信息不丢失。
        """
        if not (diffs or not result.success or result.warning):
            return
        remarks.append(f"{label}:{result.message}" if result.message else label)

    def _write_edit_log(self, log_file, log_data_path, index, order_no, remarks, sap_amount_vat):
        """写编辑结果到 log，操作类型标记 Edit。"""
        log_file.loc[index, '操作类型'] = 'Edit'
        log_file.loc[index, 'Order No.'] = order_no
        log_file.loc[index, 'Remark'] = ';'.join([item for item in remarks if item])
        log_file.loc[index, 'sapAmountVat'] = sap_amount_vat
        log_file.loc[index, 'Update Time'] = datetime.datetime.today().strftime('%Y-%m-%d %H:%M:%S')
        log_file.to_excel(log_data_path, merge_cells=False, index=False)

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
            log_file['操作类型'] = ''
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

                # Invoice Number 有值 = 该单已开票，无论是否有 Order Number 都不新建/编辑，直接跳过。
                # 放在 Combine Id 之后、GUI 回填/确认弹窗/对象构建之前，避免对已开票单做任何无谓动作。
                invoice_number = self._excel_str(order_row.get('Invoice Number'))
                if invoice_number:
                    skip_msg = 'Invoice Number 已有值（%s），跳过不新建/编辑' % invoice_number
                    log_file.loc[index, 'Remark'] = skip_msg
                    log_file.to_excel(log_data_path, merge_cells=False, index=False)
                    self.textBrowser.append(
                        "<font color='orange'>No.%s %s</font>" % (index + 1, skip_msg)
                    )
                    QApplication.processEvents()
                    continue

                # Primary CS 必填：为空、或 CS 名未录入 config 人员名单（解析不出 CS 编号）时，
                # SAP 伙伴页的"负责雇员"无从写入，创建/编辑都没有意义，直接跳过当前订单。
                # 与 Invoice Number 同层拦截，故创建与编辑两条分支同时覆盖。
                cs_name = self._excel_str(order_row.get('Primary CS'))
                if not self._resolve_cs_code(order_row):
                    cs_msg = (
                        'Primary CS 为空（必填），跳过不新建/编辑' if not cs_name
                        else 'Primary CS [%s] 不在配置文件中（取不到 CS 编号），跳过不新建/编辑' % cs_name
                    )
                    log_file.loc[index, 'Remark'] = cs_msg
                    log_file.to_excel(log_data_path, merge_cells=False, index=False)
                    self.textBrowser.append(
                        "<font color='red'>No.%s %s</font>" % (index + 1, cs_msg)
                    )
                    QApplication.processEvents()
                    continue

                # 将当前订单关键字段回填到 GUI 控件，便于用户实时跟踪正在处理的订单。
                self._apply_order_row_to_gui(order_row)

                # everyCheck（checkBox_16）：每单开始前弹窗让用户确认是否处理；
                # 选 No 跳过当前订单（log 标记），继续下一单。放在 GUI 回填后、对象构建前，
                # 避免对跳过单做无谓的 DataFrame 转换和 SAP 校验。
                if flow_options.get('everyCheck'):
                    combine_id_preview = self._excel_str(order_row.get('Combine Id'))
                    project_no_preview = self._excel_str(order_row.get('Request Number'))
                    cs_preview = self._excel_str(order_row.get('Primary CS'))
                    sales_preview = self._excel_str(order_row.get('Sales'))
                    confirm_msg = (
                        '是否处理 No.%s 订单？\n\n'
                        'Combine Id: %s\n'
                        'Request Number: %s\n'
                        'Primary CS: %s\n'
                        'Sales: %s\n\n'
                        '点击 Yes 继续，No 跳过当前订单'
                    ) % (index + 1, combine_id_preview, project_no_preview, cs_preview, sales_preview)
                    reply = QMessageBox.question(
                        self,
                        '订单确认',
                        confirm_msg,
                        QMessageBox.Yes | QMessageBox.No,
                        QMessageBox.Yes,
                    )
                    if reply != QMessageBox.Yes:
                        log_file.loc[index, 'Remark'] = '用户选择跳过'
                        log_file.to_excel(log_data_path, merge_cells=False, index=False)
                        self.textBrowser.append(
                            "<font color='orange'>No.%s 用户选择跳过</font>" % (index + 1)
                        )
                        QApplication.processEvents()
                        continue

                # 三张 DataFrame 已包含完整业务数据，这里只做对象适配和 SAP 写入。
                order = self._build_order_from_dataframes(order_row, item_df)
                revenue = self._build_revenue_from_order_row(order_row)
                config = self._build_sap_config_from_order_row(order_row)
                data_b_entries, plan_cost_entries_by_item = self._build_sub_entries_from_dataframe(order_row, sub_df)
                service = OrderService(sap_session, config)

                # ===== 编辑分流：Order Number 有值 = 订单已存在，跳过 VA01 创建，进 VA02 对比更新 =====
                # Excel 会混排"有号=编辑 / 无号=创建"两类行；本分支只接管有号行，无号行落回下方原创建逻辑（行为不变）。
                excel_order_no = self._excel_str(order_row.get('Order Number'))
                if excel_order_no:
                    self._edit_order_row(
                        index, order_row, order, config,
                        data_b_entries, plan_cost_entries_by_item,
                        excel_order_no, flow_options, sap_session, log_file, log_data_path,
                    )
                    continue

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

                if need_va01_check and self._is_date_before_today(order.ecd):
                    ecd_msg = (
                        "VA01创建失败：ECD %s 早于今天，不能早于订单创建日期"
                        % order.ecd
                    )
                    log_file.loc[index, 'Remark'] = ecd_msg
                    log_file.to_excel(log_data_path, merge_cells=False, index=False)
                    self.textBrowser.append(
                        "<font color='red'>No.%s %s</font>" % (index + 1, ecd_msg)
                    )
                    QApplication.processEvents()
                    continue

                # textBrowser 抬头：基础信息 + Excel 未税金额。
                combine_id = self._excel_str(order_row.get('Combine Id'))
                primary_cs = self._excel_str(order_row.get('Primary CS'))
                sales_name = self._excel_str(order_row.get('Sales'))
                excel_amount_untaxed = self._excel_str(order_row.get('Untaxed amount'))
                items_revenue_total = sum(item.revenue for item in order.items)

                self.textBrowser.append('==================== No.%s ====================' % (index + 1))
                self.textBrowser.append("Combine Id: %s" % combine_id)
                self.textBrowser.append("Request Number: %s" % order.project_no)
                self.textBrowser.append("Primary CS: %s" % primary_cs)
                self.textBrowser.append("Sales: %s" % sales_name)
                self.textBrowser.append("未税金额(Excel): %s" % excel_amount_untaxed)
                self.textBrowser.append("Items 加和金额: %s" % format(items_revenue_total, ',.2f'))
                QApplication.processEvents()

                remarks = []
                # 业务流程：VA01(可选) -> Save VA01 -> 打开 VA02 -> Add Item
                # -> Plan Cost(可选) -> Save VA02 -> 打开 VA02 -> Data B(可选) -> Save VA02。
                # Save 复选框控制常规保存；但本次新增 item 后再做 Data B 时，会先强制保存 item，
                # 因为 Data B 依赖已落盘的 SAP item 号（1000/2000 等）。
                sap_amount_vat = ''

                _report_step = self._append_step_result

                # Step 1: VA01 创建订单头
                va01_done = False
                if flow_options.get('va01Check'):
                    # contactCheck（checkBox_19）：未勾选时 add_contact=False，
                    # _fill_partners 内部跳过联系人写入；add_sales_partner 保持默认。
                    partner_options = PartnerOptions(
                        add_contact=bool(flow_options.get('contactCheck')),
                    )
                    create_result = service.create_order(
                        order, revenue, partner_options=partner_options
                    )
                    remarks.append(f"VA01:{create_result.message}" if create_result.message else "VA01")
                    order_no = create_result.order_no or order_no
                    sap_amount_vat = create_result.sap_amount_vat or sap_amount_vat
                    va01_done = create_result.success
                    _report_step('VA01', create_result)

                # Step 2: Save VA01 —— VA01 成功后，若有后续步骤或显式 saveCheck 都需要落盘
                need_save_va01 = va01_done and (
                    flow_options.get('saveCheck')
                    or flow_options.get('va02Check')
                    or flow_options.get('labCostCheck')
                    or flow_options.get('planCostCheck')
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
                        first_va02_changed = False
                        item_added = False
                        item_failed = False
                        pre_data_b_save_ok = True
                        # add item 仅在 va02Check 时进行；纯 Data B / Plan Cost 场景不重复加 item。
                        if flow_options.get('va02Check'):
                            item_result = service.add_items(order, revenue)
                            order_no = item_result.order_no or order_no
                            remarks.append(f"Item:{item_result.message}" if item_result.message else "Item")
                            sap_amount_vat = item_result.sap_amount_vat or sap_amount_vat
                            item_added = item_result.success
                            item_failed = not item_result.success
                            first_va02_changed = first_va02_changed or item_result.success
                            _report_step('Add Item', item_result)

                        if flow_options.get('planCostCheck') and not item_failed:
                            # 按 item 号让 SAP 侧实时定位物理行（target_item）——SAP 写完 item
                            # 回车后按 POSNR 重排，列表索引不等于物理行；索引仅作 Excel 未给
                            # item 号（由 SAP 自动分配）时的兜底。
                            # sub 表未提供 plan cost 数据的 item 直接跳过。
                            for row, item in enumerate(order.items):
                                plan_cost_entries = plan_cost_entries_by_item.get(item.item)
                                if not plan_cost_entries:
                                    continue
                                plan_result = service.apply_plan_cost_entries(
                                    plan_cost_entries, focus_row=row, target_item=item.item
                                )
                                first_va02_changed = first_va02_changed or plan_result.success
                                remarks.append(
                                    f"Plan Cost {item.item}:{plan_result.message}"
                                    if plan_result.message
                                    else f"Plan Cost {item.item}"
                                )
                                _report_step('Plan Cost %s' % item.item, plan_result)

                        # 订单价值(AUFTRAGSWERT) 独立步骤：item 全部录入后，读 SAP 概览净值加和 × 汇率
                        # 回填抬头字段。与 Data B(labCostCheck) 解耦，仅在本次成功新增 item 时执行；
                        # 写入抬头后由下方 Save VA02 统一落盘（first_va02_changed 已因加 item 置真）。
                        if flow_options.get('va02Check') and item_added:
                            order_value_result = service.fill_order_value(order)
                            first_va02_changed = first_va02_changed or order_value_result.success
                            remarks.append(
                                f"订单价值:{order_value_result.message}"
                                if order_value_result.message else "订单价值"
                            )
                            _report_step('订单价值', order_value_result)

                        need_data_b = flow_options.get('labCostCheck') and data_b_entries
                        # Data B 的 item 依赖已保存的 SAP item 号（1000/2000 等）。
                        # 如果本次新增了 item，即使未勾选 Save，也要先保存再重新打开 VA02 写 Data B。
                        need_save_before_data_b = bool(need_data_b and item_added)
                        need_first_va02_save = first_va02_changed and (
                            flow_options.get('saveCheck') or need_save_before_data_b
                        )
                        if need_first_va02_save:
                            save_step_name = (
                                'Save VA02 Before Data B'
                                if need_save_before_data_b
                                else 'Save VA02'
                            )
                            save_va02_result = service.save('VA02')
                            pre_data_b_save_ok = save_va02_result.success
                            remarks.append(
                                f"{save_step_name}:{save_va02_result.message}"
                                if save_va02_result.message
                                else save_step_name
                            )
                            _report_step(save_step_name, save_va02_result)

                        if need_data_b and item_failed:
                            self.textBrowser.append(
                                "<font color='red'>Add Item 失败，跳过当前订单的 Data B 步骤</font>"
                            )
                            QApplication.processEvents()

                        if need_data_b and not item_failed and pre_data_b_save_ok:
                            reopen_result = SapResult()
                            if first_va02_changed and need_first_va02_save:
                                reopen_result = service.open_order(order_no)
                                order_no = reopen_result.order_no or order_no or self._extract_order_no(sap_session)
                                remarks.append(
                                    f"VA02 Data B:{reopen_result.message}"
                                    if reopen_result.message
                                    else "VA02 Data B"
                                )
                                _report_step('Open VA02 Data B', reopen_result)

                            if reopen_result.success:
                                data_b_result = service.fill_lab_cost_entries(
                                    data_b_entries,
                                    order,
                                )
                                remarks.append(
                                    f"Data B:{data_b_result.message}" if data_b_result.message else "Data B"
                                )
                                _report_step('Data B', data_b_result)

                                if flow_options.get('saveCheck'):
                                    save_data_b_result = service.save('VA02')
                                    remarks.append(
                                        f"Save VA02 Data B:{save_data_b_result.message}"
                                        if save_data_b_result.message
                                        else "Save VA02 Data B"
                                    )
                                    _report_step('Save VA02 Data B', save_data_b_result)

                    # VA02 段结束后再做一次最终兜底，覆盖中间步骤未回传 order_no 的边界情况。
                    if not order_no:
                        order_no = self._extract_order_no(sap_session)

                # SAP 加和金额为未税净值，理论上应等于 Excel "Untaxed amount"。
                # 容差 0.01 容忍浮点误差；只有在 Excel 未税金额可用时才比较，避免空值误判。
                try:
                    sap_amount_value = float(str(sap_amount_vat).replace(',', '')) if sap_amount_vat else 0.0
                except (TypeError, ValueError):
                    sap_amount_value = 0.0
                try:
                    excel_amount_value = float(str(excel_amount_untaxed).replace(',', '')) if excel_amount_untaxed else 0.0
                except (TypeError, ValueError):
                    excel_amount_value = 0.0

                amount_diff = round(sap_amount_value - excel_amount_value, 2)
                amount_mismatch = excel_amount_value > 0 and abs(amount_diff) >= 0.01
                if amount_mismatch:
                    diff_msg = (
                        f"未税金额不一致: Excel={format(excel_amount_value, ',.2f')} "
                        f"SAP={format(sap_amount_value, ',.2f')} "
                        f"差额={format(amount_diff, ',.2f')}"
                    )
                    remarks.append(diff_msg)

                log_file.loc[index, '操作类型'] = 'Create'
                log_file.loc[index, 'Order No.'] = order_no
                log_file.loc[index, 'Remark'] = ';'.join([item for item in remarks if item])
                log_file.loc[index, 'Proforma No.'] = ''
                log_file.loc[index, 'sapAmountVat'] = sap_amount_vat
                log_file.loc[index, 'Update Time'] = datetime.datetime.today().strftime('%Y-%m-%d %H:%M:%S')
                log_file.to_excel(log_data_path, merge_cells=False, index=False)

                # 订单结束摘要：order no + SAP 未税金额；与 Excel 未税金额一致性提示。
                self.textBrowser.append("Order No.: %s" % order_no)
                self.textBrowser.append("SAP 金额(加和,未税): %s" % (sap_amount_vat or '--'))
                if amount_mismatch:
                    self.textBrowser.append("<font color='red'>%s</font>" % diff_msg)
                elif excel_amount_value > 0:
                    self.textBrowser.append("未税金额一致(Excel == SAP)")
                self.textBrowser.append('----------------------------------')
                QApplication.processEvents()

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


