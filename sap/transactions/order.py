"""Order transactions."""

from __future__ import annotations

import re
import time

from sap.models import CostOptions, OrderData, OrderItemData, PartnerOptions, RevenueData, SapConfig, SapResult
from sap.rules import (
    build_fremdl_entry,
    build_lab_cost_entries,
    build_single_plan_cost_entries,
    build_split_plan_cost_entries,
    has_430_subcode,
    is_a2_material,
    is_d_split_material,
    resolve_a2_materials,
    resolve_data_a_key,
    should_apply_plan_cost,
    should_fill_auftragswert,
)
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

    def create(self, order: OrderData, revenue: RevenueData, options: PartnerOptions) -> SapResult:
        """Create order header in VA01."""
        result = SapResult(step="va01")
        try:
            # VA01 澶撮儴鏁版嵁鍐欏叆銆?            self.session.set_text("wnd[0]/tbar[0]/okcd", "/nva01")
            self.session.send_vkey(0)
            self.session.set_text("wnd[0]/usr/ctxtVBAK-AUART", self.config.order_type)
            self.session.set_text("wnd[0]/usr/ctxtVBAK-VKORG", self.config.sales_organization)
            self.session.set_text("wnd[0]/usr/ctxtVBAK-VTWEG", self.config.distribution_channels)
            self.session.set_text("wnd[0]/usr/ctxtVBAK-VKBUR", self.config.sales_office)
            # TODO 杩欎釜浣嶇疆搴旇鏄鍗曚俊鎭殑锛屽鏋滆鍗曚俊鎭病鏈夎瀛楁锛屼娇鐢ㄩ粯璁ょ殑
            self.session.set_text("wnd[0]/usr/ctxtVBAK-VKGRP", self.config.cost_center)
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

            # DATA A / DATA B 鏄鍗曞ご涓婄殑涓ょ粍涓氬姟瀛楁銆?            self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13")
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
            return SapResult.fail(f"Order No鏈垱寤烘垚鍔燂紝{exc}", step="va01")
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
        e_row, g_row = (4, 5) if four_name in {"璐熻矗闆囧憳", "Employee respons."} else (5, 4)

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

        if options.add_sales_partner and order.sales_name:
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
        # TODO 鍚屾牱鐨勫簲璇ユ槸鐩存帴鏇存柊涓烘渶鏂扮殑鏂囨湰鍐呭
        self.session.set_text(text_id, order.short_text)
        self.session.set_selection_indexes(text_id, 11, 11)
        self.session.set_key(lang_id, "EN")
        self.session.focus(lang_id)
        self.session.send_vkey(0)

    def fill_lab_cost(self, order: OrderData, revenue: RevenueData) -> SapResult:
        """Write Data B labor cost."""
        result = SapResult(step="lab_cost")
        try:
            entries = build_lab_cost_entries(order, revenue, self.config)
            for entry in entries:
                self.session.set_text(
                    f"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/"
                    f"tblSAPMV45AZULEISTENDE/ctxtTABL-KOSTL[0,{entry.row}]",
                    entry.performer_cost_center,
                )
                self.session.set_text(
                    f"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/"
                    f"tblSAPMV45AKOSTENSAETZE/ctxtTABD-KOSTL[0,{entry.row}]",
                    entry.rate_cost_center,
                )
                self.session.set_text(
                    f"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/"
                    f"tblSAPMV45AKOSTENSAETZE/txtTABD-FESTPREIS[5,{entry.row}]",
                    entry.amount,
                )
        except Exception as exc:
            return SapResult.fail(f"Data B鏈～鍐欙紝{exc}", step="lab_cost")
        return result

    def save(self, info: str) -> SapResult:
        """Save current order page and verify status."""
        result = SapResult(step="save")
        save_error: Exception | None = None
        try:
            # 鐜版湁涓氬姟椤甸潰淇濆瓨鍓嶉€氬父闇€瑕佸厛鍥為€€鍒板彲纭鐨勫眰绾с€?            self.session.press("wnd[0]/tbar[0]/btn[3]")
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
            message = f"{info}淇濆瓨澶辫触锛屾棤娉曡鍙栫姸鎬佹爮: {exc}"
            if save_error:
                message += f"锛涗繚瀛樻搷浣滃紓甯? {save_error}"
            return SapResult.fail(message, step="save")

        if "saved" not in save_msg.lower() and "保存" not in save_msg:
            message = f"{info}保存失败，{save_msg}"
            if save_error:
                message += f"锛涗繚瀛樻搷浣滃紓甯? {save_error}"
            return SapResult.fail(message, step="save")
        return result

    def open(self, order_no: str) -> SapResult:
        """Open an existing order in VA02."""
        result = SapResult(step="open_va02")
        try:
            self.session.set_text("wnd[0]/tbar[0]/okcd", "/NVA02")
            self.session.send_vkey(0)
            self.session.set_text("wnd[0]/usr/ctxtVBAK-VBELN", order_no)
            self.session.send_vkey(0)
        except Exception as exc:
            return SapResult.fail(f"璇rder No {order_no} 鏈紑鍚紝{exc}", step="open_va02")
        return result

    def add_items(self, order: OrderData, revenue: RevenueData) -> SapResult:
        """Add item rows to current order."""
        result = SapResult(step="va02")
        try:
            order_no = self.session.read_text("wnd[0]/usr/ctxtVBAK-VBELN")
            result.order_no = order_no
            result.sap_amount_vat = self._write_item_rows(order)
            return result

        except Exception as exc:
            return SapResult.fail(f"Order add item failed: {exc}", step="va02")
        return result

    def update_items(self, order: OrderData, revenue: RevenueData) -> SapResult:
        """Update current order items after VA02 is open."""
        result = SapResult(step="va02_update")
        try:
            result.order_no = self.session.read_text("wnd[0]/usr/ctxtVBAK-VBELN")
            result.sap_amount_vat = self._write_item_rows(order)
        except Exception as exc:
            return SapResult.fail(f"Order update item failed: {exc}", step="va02_update")
        return result

    def _resolve_order_items(self, order: OrderData) -> list[OrderItemData]:
        items = [item for item in order.items if item.material_code]
        if not items:
            raise ValueError("order.items is required")
        return items

    def _write_item_rows(self, order: OrderData) -> str:
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
            if order.long_text and row == 0:
                self._write_item_long_text(order.long_text, SapResult())
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
            result.append_message("Long Text 娣诲姞澶辫触")

    @staticmethod
    def _parse_amount(amount_text: str) -> float:
        """Parse SAP amount text."""
        return float(amount_text.replace(",", ""))

    @staticmethod
    def _format_amount(amount: float) -> str:
        """Format amount with thousands separators."""
        return re.sub(r"(\d)(?=(\d\d\d)+(?!\d))", r"\1,", format(amount, ".2f"))

    def apply_plan_cost(self, order: OrderData, revenue: RevenueData, options: CostOptions) -> SapResult:
        """Apply plan cost by material type."""
        result = SapResult(step="plan_cost")
        try:
            if not should_apply_plan_cost(revenue, self.config):
                return result

            if is_d_split_material(order.material_code):
                # D2/D3 鏄崟 item + 鎷嗗垎鎴愭湰瑙嗗浘銆?                self.session.press("wnd[0]/tbar[0]/btn[3]")
                self._open_plan_cost_editor(self._material_id(0))
                entries = build_split_plan_cost_entries(revenue, self.config, options)
                next_row = self._apply_plan_cost_entries(entries)
                fremdl_entry = build_fremdl_entry(next_row, order, self.config)
                if fremdl_entry:
                    self._apply_single_plan_cost_entry(fremdl_entry)
                self.session.press("wnd[0]/tbar[0]/btn[3]")
                self.session.press("wnd[1]/usr/btnSPOP-OPTION1")
                return result

            if is_a2_material(order.material_code):
                # A2 闇€瑕佸垎鍒繘鍏ヤ袱涓?item 鐨勮鍒掓垚鏈晫闈€?                self.session.press("wnd[0]/tbar[0]/btn[3]")
                first_target = self._material_id(0) if has_430_subcode(order.material_code) else self._material_id(1)
                second_target = self._material_id(1) if has_430_subcode(order.material_code) else self._material_id(0)

                self._open_plan_cost_editor(first_target)
                first_entries = []
                row = 0
                if options.include_cs:
                    phy_cs = round(float(revenue.phy_cs_cost), 0)
                    if phy_cs > 0:
                        from sap.rules import PlanCostEntry

                        first_entries.append(PlanCostEntry(row, self.config.sub_cost_center_cs, "T01AST", phy_cs))
                        row += 1
                if options.include_phy:
                    phy_lab = round(float(revenue.phy_lab_cost), 0)
                    if phy_lab > 0:
                        from sap.rules import PlanCostEntry

                        first_entries.append(PlanCostEntry(row, self.config.sub_cost_center_phy, "T01AST", phy_lab))
                        row += 1
                next_row = self._apply_plan_cost_entries(first_entries)
                if has_430_subcode(order.material_code):
                    # 430 瀛愮爜鐨勫鍖呮垚鏈寕鍦ㄧ涓€娈佃鍒掓垚鏈噷銆?                    fremdl_entry = build_fremdl_entry(next_row, order, self.config)
                    if fremdl_entry:
                        self._apply_single_plan_cost_entry(fremdl_entry)
                self.session.press("wnd[0]/tbar[0]/btn[3]")
                self.session.press("wnd[0]/tbar[0]/btn[3]")
                self.session.press("wnd[1]/usr/btnSPOP-OPTION1")

                self._open_plan_cost_editor(second_target)
                second_entries = []
                row = 0
                if options.include_cs:
                    chm_cs = round(float(revenue.chm_cs_cost), 0)
                    if chm_cs > 0:
                        from sap.rules import PlanCostEntry

                        second_entries.append(PlanCostEntry(row, self.config.sub_cost_center_cs, "T01AST", chm_cs))
                        row += 1
                if options.include_chm:
                    chm_lab = round(float(revenue.chm_lab_cost), 0)
                    if chm_lab > 0:
                        from sap.rules import PlanCostEntry

                        second_entries.append(PlanCostEntry(row, self.config.sub_cost_center_chm, "T01AST", chm_lab))
                        row += 1
                next_row = self._apply_plan_cost_entries(second_entries)
                if not has_430_subcode(order.material_code):
                    # 闈?430 瀛愮爜鍒欐妸澶栧寘鎴愭湰鎸傚埌绗簩娈点€?                    fremdl_entry = build_fremdl_entry(next_row, order, self.config)
                    if fremdl_entry:
                        self._apply_single_plan_cost_entry(fremdl_entry)
                self.session.press("wnd[0]/tbar[0]/btn[3]")
                self.session.press("wnd[1]/usr/btnSPOP-OPTION1")
                return result

            self.session.press("wnd[0]/tbar[0]/btn[3]")
            self._open_plan_cost_editor(self._material_id(0))
            entries = build_single_plan_cost_entries(order, revenue, self.config, options)
            next_row = self._apply_plan_cost_entries(entries)
            fremdl_entry = build_fremdl_entry(next_row, order, self.config)
            if fremdl_entry:
                self._apply_single_plan_cost_entry(fremdl_entry)
            self.session.press("wnd[0]/tbar[0]/btn[3]")
            self.session.press("wnd[1]/usr/btnSPOP-OPTION1")
        except Exception as exc:
            return SapResult.fail(f"plan cost鏈坊鍔犳垚鍔?{exc}", step="plan_cost")
        return result

    def _open_plan_cost_editor(self, focus_element_id: str) -> None:
        """Open plan cost editor for focused item."""
        self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02")
        # 璁″垝鎴愭湰鑿滃崟渚濊禆褰撳墠鐒︾偣 item锛屽繀椤诲厛鎶婂厜鏍囨斁鍒扮洰鏍囩墿鏂欒銆?        self.session.focus(focus_element_id, 10)
        self.session.find("wnd[0]/mbar/menu[3]/menu[7]").select()
        self.session.press("wnd[1]/usr/btnSPOP-VAROPTION1")
        self.session.press("wnd[1]/tbar[0]/btn[0]")

    def _apply_plan_cost_entries(self, entries) -> int:
        """Apply plan cost entries and return next row."""
        row = 0
        for entry in entries:
            self._apply_single_plan_cost_entry(entry)
            row = entry.row + 1
        # 杩斿洖涓嬩竴鍙敤琛岋紝渚?FREMDL 杩欑被闄勫姞椤圭户缁啓鍏ャ€?        return row

    def _apply_single_plan_cost_entry(self, entry) -> None:
        """Apply one plan cost entry."""
        self.session.set_text(f"wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-TYPPS[2,{entry.row}]", "E")
        self.session.set_text(
            f"wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,{entry.row}]",
            entry.cost_center,
        )
        self.session.set_text(
            f"wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,{entry.row}]",
            entry.category,
        )
        self.session.set_text(
            f"wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,{entry.row}]",
            round(float(entry.amount), 0) if entry.category == "T01AST" else format(entry.amount, ".2f"),
        )
        self.session.focus(f"wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,{entry.row}]", 20)
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
            return SapResult.fail(f"{action} 鏈垚鍔燂紝{exc}", step="lock")
        return result
