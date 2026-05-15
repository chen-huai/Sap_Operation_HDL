"""Order transactions."""

from __future__ import annotations

import re
import time

from sap.models import (
    DataBEntry,
    OrderData,
    OrderItemData,
    PartnerOptions,
    PlanCostEntry,
    RevenueData,
    SapConfig,
    SapResult,
)
from sap.rules import resolve_data_a_key, should_fill_auftragswert
from sap.session import SapSession


class OrderTransaction:
    """Encapsulate VA01/VA02 order operations."""

    def __init__(self, session: SapSession, config: SapConfig):
        """Initialize with shared SAP session and config."""
        self.session = session
        self.config = config

    @property
    def today(self) -> str:
        """Return today in SAP date format."""
        return time.strftime("%Y.%m.%d")

    def _safe_read_order_no(self, fallback: str = "") -> str:
        """多重兜底读取当前订单号: 抬头 VBELN → 状态栏数字 → fallback。"""
        try:
            order_no = (self.session.read_text("wnd[0]/usr/ctxtVBAK-VBELN") or "").strip()
            if order_no:
                return order_no
        except Exception:
            pass
        try:
            status_text = self.session.read_status() or ""
            match = re.search(r"(\d{6,})", status_text)
            if match:
                return match.group(1)
        except Exception:
            pass
        return fallback

    def create(self, order: OrderData, revenue: RevenueData, options: PartnerOptions) -> SapResult:
        """Create order header in VA01."""
        result = SapResult(step="va01")
        try:
            # VA01 头部数据写入。
            self.session.set_text("wnd[0]/tbar[0]/okcd", "/nva01")
            self.session.send_vkey(0)
            self.session.set_text("wnd[0]/usr/ctxtVBAK-AUART", self.config.order_type)
            self.session.set_text("wnd[0]/usr/ctxtVBAK-VKORG", self.config.sales_organization)
            self.session.set_text("wnd[0]/usr/ctxtVBAK-VTWEG", self.config.distribution_channels)
            self.session.set_text("wnd[0]/usr/ctxtVBAK-VKBUR", self.config.sales_office)
            self.session.set_text("wnd[0]/usr/ctxtVBAK-VKGRP", order.sales_group)
            self.session.send_vkey(0)

            customer_id = (
                "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/"
                "subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR"
            )
            self.session.set_text(customer_id, order.sap_no)
            self.session.focus(customer_id, 6)
            self.session.send_vkey(0)
            self.session.send_vkey(0)

            self.session.set_text("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD", order.project_no)
            self.session.set_text("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBKD-BSTDK", self.today)
            self.session.set_text(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/"
                "ssubSUBSCREEN_BODY:SAPMV45A:4415/ctxtVBKD-FBUDA",
                self.today,
            )
            self.session.focus("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD", 17)
            self.session.send_vkey(0)
            self.session.press("wnd[1]/tbar[0]/btn[0]")
            self.session.press("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD")

            currency_id = (
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/"
                "ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBAK-WAERK"
            )
            self.session.set_text(currency_id, order.currency_type)
            self.session.focus(currency_id, 3)
            self.session.send_vkey(0)
            self.session.try_send_vkey(0, window_id="wnd[1]")

            if order.currency_type != "CNY":
                rate_id = (
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/"
                    "ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBKD-KURSK"
                )
                self.session.set_text(rate_id, order.exchange_rate)
                self.session.focus(rate_id, 8)
                self.session.send_vkey(0)

            self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\06")
            accounting_id = (
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\06/"
                "ssubSUBSCREEN_BODY:SAPMV45A:4311/txtVBAK-XBLNR"
            )
            self.session.set_text(accounting_id, "*")
            self.session.focus(accounting_id, 1)
            self.session.send_vkey(0)

            self._fill_partners(order, options)
            self._fill_header_text(order)

            # DATA A / DATA B 是订单头上的两组业务字段。
            self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13")
            self.session.set_key(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/"
                "ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR1",
                resolve_data_a_key(order, self.config),
            )

            self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14")
            self.session.set_text(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/"
                "ssubSUBSCREEN_BODY:SAPMV45A:4312/ctxtVBAK-ZZAUART",
                "WO",
            )
            self.session.set_text(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/"
                "ssubSUBSCREEN_BODY:SAPMV45A:4312/ctxtVBAK-ZZUNLIMITLIAB",
                "N",
            )
            self.session.set_text(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/"
                "ssubSUBSCREEN_BODY:SAPMV45A:4312/ctxtZAUFTD-VORAUS_AUFENDE",
                order.ecd,
            )
            if should_fill_auftragswert(revenue, self.config):
                self.session.set_text(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/"
                    "ssubSUBSCREEN_BODY:SAPMV45A:4312/txtZAUFTD-AUFTRAGSWERT",
                    format(revenue.revenue_cny, ".2f"),
                )
        except Exception as exc:
            return SapResult.fail(f"Order No未创建成功，{exc}", step="va01")
        return result

    def _fill_partners(self, order: OrderData, options: PartnerOptions) -> None:
        """Fill partner tab."""
        self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09")
        partner_prefix = (
            "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/"
            "ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:"
            "SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW"
        )
        four_name = self.session.read_text(f"{partner_prefix}/cmbGVS_TC_DATA-REC-PARVW[0,4]")
        e_row, g_row = (4, 5) if four_name in {"负责雇员", "Employee respons."} else (5, 4)

        self.session.set_key(f"{partner_prefix}/cmbGVS_TC_DATA-REC-PARVW[0,{g_row}]", "ZG")
        self.session.set_text(f"{partner_prefix}/ctxtGVS_TC_DATA-REC-PARTNER[1,{g_row}]", order.global_partner_code)
        self.session.focus(f"{partner_prefix}/ctxtGVS_TC_DATA-REC-PARTNER[1,{g_row}]", 8)
        self.session.set_text(f"{partner_prefix}/ctxtGVS_TC_DATA-REC-PARTNER[1,{e_row}]", self.config.cs_code)
        self.session.send_vkey(0)

        if options.add_contact:
            self.session.set_key(f"{partner_prefix}/cmbGVS_TC_DATA-REC-PARVW[0,6]", "AP")
            self.session.focus(f"{partner_prefix}/ctxtGVS_TC_DATA-REC-PARTNER[1,6]", 0)
            self.session.send_vkey(4)
            self.session.press("wnd[1]/tbar[0]/btn[0]")
            self.session.press("wnd[1]/tbar[0]/btn[0]")
            self.session.send_vkey(0)

        if options.add_sales_partner and self.config.sales_code:
            self.session.set_key(f"{partner_prefix}/cmbGVS_TC_DATA-REC-PARVW[0,7]", "VE")
            self.session.set_text(f"{partner_prefix}/ctxtGVS_TC_DATA-REC-PARTNER[1,7]", self.config.sales_code)
            self.session.focus(f"{partner_prefix}/ctxtGVS_TC_DATA-REC-PARTNER[1,7]", 4)
            self.session.send_vkey(0)

    def _fill_header_text(self, order: OrderData) -> None:
        """Fill order header short text."""
        self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10")
        text_id = (
            "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10/"
            "ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/"
            "cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell"
        )
        lang_id = (
            "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10/"
            "ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cmbLV70T-SPRAS"
        )
        self.session.set_text(text_id, order.short_text)
        self.session.set_selection_indexes(text_id, 11, 11)
        self.session.set_key(lang_id, "EN")
        self.session.focus(lang_id)
        self.session.send_vkey(0)

    def fill_lab_cost_entries(
        self,
        entries: list[DataBEntry],
        *,
        auftragswert_cny: float = 0.0,
    ) -> SapResult:
        """按已计算好的 Data B 明细写入人工成本。

        Args:
            entries: Data B 明细列表（DataBEntry）。
            auftragswert_cny: 所有 item 加和金额（CNY），≥ 阈值时回填订单价值字段。

        Returns:
            SapResult: 写入成功或失败信息。
        """
        result = SapResult(step="lab_cost")
        try:
            # 进入售达方，data b：最大化主窗口，进入抬头视图，切换到 T\14 页签。
            self.session.press("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD")
            self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14")
            # 所有 item 加和金额（CNY）达到阈值时回填订单价值字段。
            if auftragswert_cny >= self.config.revenue_threshold:
                self.session.set_text(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/"
                    "ssubSUBSCREEN_BODY:SAPMV45A:4312/txtZAUFTD-AUFTRAGSWERT",
                    format(auftragswert_cny, ".2f"),
                )
            for row, entry in enumerate(entries):
                performer_cost_center = entry.performer_cost_center.strip()
                rate_cost_center = (entry.rate_cost_center or performer_cost_center).strip()
                # 单条 Data B 只能对应一个 item，若上游传 "1000;3000" 这种多 item，
                # 取第一个 ";" 之前的部分，保留 SAP POSNR 字段单值约束。
                raw_item = (entry.item or "").strip()
                item_no = raw_item.split(";", 1)[0].strip() if raw_item else ""
                if not performer_cost_center and not rate_cost_center:
                    continue
                # Data B 页签中同一行需要同时写执行部门、费率成本中心、item 号和固定价格。
                self.session.set_text(
                    f"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/"
                    f"tblSAPMV45AZULEISTENDE/ctxtTABL-KOSTL[0,{row}]",
                    performer_cost_center,
                )
                self.session.set_text(
                    f"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/"
                    f"tblSAPMV45AKOSTENSAETZE/ctxtTABD-KOSTL[0,{row}]",
                    rate_cost_center,
                )
                if item_no:
                    # 新增 item 号写入；缺失时跳过，由 SAP 默认行为兜底。
                    self.session.set_text(
                        f"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/"
                        f"tblSAPMV45AKOSTENSAETZE/txtTABD-POSNR[1,{row}]",
                        item_no,
                    )
                self.session.set_text(
                    f"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/"
                    f"tblSAPMV45AKOSTENSAETZE/txtTABD-FESTPREIS[5,{row}]",
                    entry.amount,
                )
        except Exception as exc:
            return SapResult.fail(f"Data B未填写，{exc}", step="lab_cost")
        return result

    def save(self, info: str) -> SapResult:
        """Save current order page and verify status."""
        result = SapResult(step="save")
        save_error: Exception | None = None
        try:
            # 现有业务页面保存前通常需要先回退到可确认的层级。
            self.session.press("wnd[0]/tbar[0]/btn[3]")
            self.session.press("wnd[0]/tbar[0]/btn[3]")
            self.session.press("wnd[1]/usr/btnSPOP-OPTION1")
        except Exception as exc:
            save_error = exc
            try:
                self.session.press("wnd[0]/tbar[0]/btn[3]")
                self.session.press("wnd[0]/tbar[0]/btn[3]")
                self.session.press("wnd[1]/usr/btnSPOP-OPTION1")
                self.session.press("wnd[1]/usr/btnSPOP-OPTION1")
                save_error = None
            except Exception as retry_exc:
                save_error = retry_exc
        else:
            try:
                self.session.press("wnd[1]/usr/btnSPOP-OPTION1")
                self.session.press("wnd[1]/usr/btnSPOP-OPTION1")
            except Exception:
                pass

        try:
            save_msg = self.session.read_status()
        except Exception as exc:
            message = f"{info}保存失败，无法读取状态栏: {exc}"
            if save_error:
                message += f"；保存操作异常: {save_error}"
            return SapResult.fail(message, step="save")

        if "saved" not in save_msg.lower() and "保存" not in save_msg:
            message = f"{info}保存失败，{save_msg}"
            if save_error:
                message += f"；保存操作异常: {save_error}"
            return SapResult.fail(message, step="save")
        return result

    def open(self, order_no: str) -> SapResult:
        """Open an existing order in VA02."""
        result = SapResult(step="open_va02")
        try:
            # 先读用户在 SAP VA02 输入框中预填的值；/NVA02 命令会清空界面，需提前保存。
            prefilled_order_no = ""
            try:
                prefilled_order_no = (
                    self.session.read_text("wnd[0]/usr/ctxtVBAK-VBELN") or ""
                ).strip()
            except Exception:
                prefilled_order_no = ""

            self.session.set_text("wnd[0]/tbar[0]/okcd", "/NVA02")
            self.session.send_vkey(0)

            # 优先级：调用方传入的 order_no > 用户预填值；两者皆空时让 SAP 自己报错。
            target_order_no = (str(order_no).strip() if order_no else "") or prefilled_order_no
            if target_order_no:
                self.session.set_text("wnd[0]/usr/ctxtVBAK-VBELN", target_order_no)
            self.session.send_vkey(0)

            # SAP 左下角状态栏出现 E 类错误时，说明输入框为空或订单号错误。
            try:
                msg_type = self.session.find("wnd[0]/sbar").messageType
                msg_text = self.session.read_status()
            except Exception:
                msg_type, msg_text = "", ""
            if msg_type == "E":
                return SapResult.fail(
                    f"Order No 打开失败: {msg_text or '订单号为空或不存在'}",
                    step="open_va02",
                )
            # 进入 VA02 抬头后多重兜底读取订单号，避免抬头 VBELN 控件读取失败时丢失 order_no。
            result.order_no = self._safe_read_order_no(fallback=target_order_no)
        except Exception as exc:
            return SapResult.fail(f"该Order No {order_no} 未打开，{exc}", step="open_va02")
        return result

    def add_items(self, order: OrderData, revenue: RevenueData) -> SapResult:
        """Add item rows to current order."""
        result = SapResult(step="va02")
        try:
            # 进入 item 前先把 order_no 兜底进 result，确保从 VA02 直接开始的场景也能回传订单号。
            result.order_no = self._safe_read_order_no()
            result.sap_amount_vat = self._write_item_rows(order, result)
            # item 写完后 SAP 回到 VA02 抬头，再次刷新 order_no，覆盖更准确的值。
            result.order_no = self._safe_read_order_no(fallback=result.order_no)
            return result

        except Exception as exc:
            return SapResult.fail(f"Order add item failed: {exc}", step="va02")

    def update_items(self, order: OrderData, revenue: RevenueData) -> SapResult:
        """Update current order items after VA02 is open."""
        result = SapResult(step="va02_update")
        try:
            result.order_no = self._safe_read_order_no()
            result.sap_amount_vat = self._write_item_rows(order, result)
            result.order_no = self._safe_read_order_no(fallback=result.order_no)
        except Exception as exc:
            return SapResult.fail(f"Order update item failed: {exc}", step="va02_update")
        return result

    def _resolve_order_items(self, order: OrderData) -> list[OrderItemData]:
        items = [item for item in order.items if item.material_code]
        if not items:
            raise ValueError("order.items is required")
        return items

    def _write_item_rows(self, order: OrderData, result: SapResult) -> str:
        items = self._resolve_order_items(order)
        sap_amount_total = 0.0
        sap_amount_text = ""

        for row, item in enumerate(items):
            self._write_item_row(row, item)
        self.session.send_vkey(0)

        for row, item in enumerate(items):
            self.session.focus(self._material_id(row), 10)
            self.session.send_vkey(2)
            amount_text = self._write_item_condition(format(item.revenue, ".2f"))
            sap_amount_text = amount_text
            sap_amount_total += self._parse_amount(amount_text)
            # condition 写完仍处于 item 详情视图，借机写入 Long Text 后再返回 item 列表。
            if item.long_text:
                self._write_item_long_text(item.long_text, result)
            self.session.press("wnd[0]/tbar[0]/btn[3]")

        if len(items) > 1:
            return self._format_amount(sap_amount_total)
        return sap_amount_text

    def _write_item_row(self, row: int, item: OrderItemData) -> None:
        """Write one item row."""
        if item.item:
            self.session.set_text(self._item_id(row), item.item)
        self.session.set_text(self._material_id(row), item.material_code)
        self.session.set_text(self._quantity_id(row), item.quantity)
        self.session.set_text(self._unit_id(row), item.unit)

    @staticmethod
    def _item_id(row: int) -> str:
        """Return item number field id for row."""
        return (
            "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/"
            "ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/"
            f"tblSAPMV45ATCTRL_U_ERF_GUTLAST/txtVBAP-POSNR[0,{row}]"
        )

    @staticmethod
    def _material_id(row: int) -> str:
        """Return material field id for row."""
        return (
            "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/"
            "ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/"
            f"tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtRV45A-MABNR[1,{row}]"
        )

    @staticmethod
    def _quantity_id(row: int) -> str:
        """Return quantity field id for row."""
        return (
            "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/"
            "ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/"
            f"tblSAPMV45ATCTRL_U_ERF_GUTLAST/txtVBAP-ZMENG[2,{row}]"
        )

    @staticmethod
    def _unit_id(row: int) -> str:
        """Return unit field id for row."""
        return (
            "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/"
            "ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/"
            f"tblSAPMV45ATCTRL_U_ERF_GUTLAST/ctxtVBAP-ZIEME[3,{row}]"
        )

    def _write_item_condition(self, value) -> str:
        """Open item condition tab and write amount."""
        condition_id = (
            "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/"
            "ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/"
            "txtKOMV-KBETR[3,5]"
        )
        self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06")
        self.session.set_text(condition_id, value)
        self.session.focus(condition_id, 16)
        self.session.send_vkey(0)
        return self.session.read_text(condition_id)

    def _write_item_long_text(self, long_text: str, result: SapResult) -> None:
        """Write item long text."""
        text_id = (
            "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09/"
            "ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/"
            "cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell"
        )
        lang_id = (
            "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09/"
            "ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cmbLV70T-SPRAS"
        )
        self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09")
        self.session.set_text(text_id, long_text)
        self.session.set_selection_indexes(text_id, 4, 4)
        try:
            self.session.set_key(lang_id, "EN")
            self.session.focus(lang_id)
            self.session.send_vkey(0)
            self.session.set_selection_indexes(text_id, 0, 0)
        except Exception:
            result.append_message("Long Text 添加失败")

    @staticmethod
    def _parse_amount(amount_text: str) -> float:
        """Parse SAP amount text."""
        return float(amount_text.replace(",", ""))

    @staticmethod
    def _format_amount(amount: float) -> str:
        """Format amount with thousands separators."""
        return re.sub(r"(\d)(?=(\d\d\d)+(?!\d))", r"\1,", format(amount, ".2f"))

    def apply_plan_cost_entries(
        self,
        entries: list[PlanCostEntry],
        *,
        focus_row: int = 0,
    ) -> SapResult:
        """按已计算好的计划成本明细写入计划成本。

        Args:
            entries: 计划成本明细列表（PlanCostEntry）。
            focus_row: SAP item 表格中需要进入计划成本界面的行号。

        Returns:
            SapResult: 写入成功或失败信息。
        """
        result = SapResult(step="plan_cost")
        try:
            # 计划成本入口依赖当前 item 行焦点，先回到 item 概览页并聚焦目标行。
            self.session.press("wnd[0]/tbar[0]/btn[3]")
            self._open_plan_cost_editor(self._material_id(focus_row))
            for row, entry in enumerate(entries):
                if not entry.cost_center:
                    continue
                self._apply_single_plan_cost_entry(row, entry)
            self.session.press("wnd[0]/tbar[0]/btn[3]")
            self.session.press("wnd[1]/usr/btnSPOP-OPTION1")
        except Exception as exc:
            return SapResult.fail(f"plan cost未添加成功，{exc}", step="plan_cost")
        return result

    def _open_plan_cost_editor(self, focus_element_id: str) -> None:
        """Open plan cost editor for focused item."""
        self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02")
        # 计划成本菜单依赖当前焦点 item，必须先把光标放到目标物料行。
        self.session.focus(focus_element_id, 10)
        self.session.find("wnd[0]/mbar/menu[3]/menu[7]").select()
        self.session.press("wnd[1]/usr/btnSPOP-VAROPTION1")
        self.session.press("wnd[1]/tbar[0]/btn[0]")

    def _apply_single_plan_cost_entry(self, row: int, entry: PlanCostEntry) -> None:
        """在 plan cost 编辑器表格的指定 row 写入一条 PlanCostEntry。

        TYPPS=E 表示单条 entry；HERK2 为成本中心；HERK3 为类别（FREMDL/T01AST）；
        MENGE 为数量/金额。统一使用 .2f 保留两位小数，避免之前对 T01AST 调用 round(0.4, 0)=0 把工时截空。
        """
        self.session.set_text(f"wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,{row}]", "E")
        self.session.set_text(
            f"wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,{row}]",
            entry.cost_center,
        )
        self.session.set_text(
            f"wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,{row}]",
            entry.category,
        )
        self.session.set_text(
            f"wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,{row}]",
            format(float(entry.amount), ".2f"),
        )
        self.session.focus(f"wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,{row}]", 20)
        self.session.send_vkey(0)

    def set_lock_state(self, *, unlocked: bool) -> SapResult:
        """Switch order lock state."""
        result = SapResult(step="lock")
        action = "Unlock" if unlocked else "Lock"
        try:
            self.session.send_vkey(0, window_id="wnd[1]")
            label_id = (
                "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/"
                "subPART-SUB:SAPMV45A:4701/lblKUAGV-KUNNR"
            )
            self.session.focus(label_id, 3)
            self.session.send_vkey(2)
            self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\12")
            self.session.press(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\12/"
                "ssubSUBSCREEN_BODY:SAPMV45A:4305/btnBT_KSTC"
            )
            if unlocked:
                self.session.set_selected(
                    "wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/"
                    "tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,1]",
                    False,
                )
                self.session.set_selected(
                    "wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/"
                    "tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,0]",
                    False,
                )
            else:
                self.session.set_selected(
                    "wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/"
                    "tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,0]",
                    True,
                )
            self.session.focus(
                "wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/"
                "tblSAPLBSVATC_EO/chkJ_STMAINT-ANWSO[0,0]"
            )
            self.session.press("wnd[0]/tbar[0]/btn[3]")
            self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13")
            self.session.set_key(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/"
                "ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR4",
                "100" if unlocked else " ",
            )
            self.session.focus(
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/"
                "ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR4"
            )
            self.session.press("wnd[0]/tbar[0]/btn[11]")
            result.message = f"{action} 成功"
        except Exception as exc:
            return SapResult.fail(f"{action} 未成功，{exc}", step="lock")
        return result
