"""Order transactions."""

from __future__ import annotations

import re
import time

from sap.exceptions import SapUiError
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
from sap.rules import (
    resolve_data_a_key,
    should_fill_ic_transaction,
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
            self._fill_submission_if_needed(order)

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
            # 命中 Data_B_TUV（TUV IC 订单）时写入 IC 交易类型 O1。
            if should_fill_ic_transaction(order, self.config):
                self.session.set_text(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/"
                    "ssubSUBSCREEN_BODY:SAPMV45A:4312/ctxtZAUFTD-IC_TRANSAKTION",
                    "O1",
                )
            # 订单价值(AUFTRAGSWERT) 不在此处写入：item 尚未录入，此时无法按"Σ item 净值 × 汇率"
            # 计算。改由 fill_order_value() 在 item 全部录入后统一回填（创建/编辑同口径）。
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

    def _fill_submission_if_needed(self, order: OrderData) -> None:
        """404 Power driven Furniture 订单需要在 Additional data B 写入 submission 标识。"""
        if order.product_sub_category != "404 Power driven Furniture":
            return

        submission_id = (
            "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\11/"
            "ssubSUBSCREEN_BODY:SAPMV45A:4351/txtVBAK-SUBMI"
        )
        self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\11")
        self.session.set_text(submission_id, "EF")
        self.session.focus(submission_id, 2)

    def fill_lab_cost_entries(
        self,
        entries: list[DataBEntry],
        order: OrderData,
    ) -> SapResult:
        """按已计算好的 Data B 明细写入人工成本。

        Args:
            entries: Data B 明细列表（DataBEntry）。
            order: 订单数据，用于读取 sales_group 决定是否写入 item 号。

        Returns:
            SapResult: 写入成功或失败信息。

        Note:
            订单价值(AUFTRAGSWERT) 已从本方法剥离，改由 fill_order_value() 独立步骤回填。
        """
        result = SapResult(step="lab_cost")
        try:
            # 进入售达方，data b：最大化主窗口，进入抬头视图，切换到 T\14 页签。
            self.session.press("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD")
            self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14")
            self._write_lab_cost_rows(entries, order)
        except Exception as exc:
            return SapResult.fail(f"Data B未填写，{exc}", step="lab_cost")
        return result

    def _write_lab_cost_rows(self, entries: list[DataBEntry], order: OrderData) -> None:
        """在已进入 T\\14 页签的前提下，从 row 0 起逐行写入 Data B 明细。

        创建(VA01)与编辑(VA02 删空后重建)共用；不含导航/删除，调用方须保证已在正确页签、
        且表格为可从 row 0 顺序追加的状态。正常行写全字段（执行部门 / ZULEISTENDE·KOSTENSAETZE
        双表 item 号 / 费率成本中心 / 固定价格）；config 强制成本中心行只写执行部门并回车让 SAP
        带出该行，绝不碰费率成本中心 / item 号 / 固定价格。
        """
        for row, entry in enumerate(entries):
            performer_cost_center = entry.performer_cost_center.strip()
            kostl_id = (
                f"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/"
                f"tblSAPMV45AZULEISTENDE/ctxtTABL-KOSTL[0,{row}]"
            )

            if entry.kostl_only:
                # config 强制成本中心行：只录执行部门，回车让 SAP 带出该行，
                # 不写 item 号 / 费率成本中心 / 固定价格。
                if not performer_cost_center:
                    continue
                self.session.set_text(kostl_id, performer_cost_center)
                self.session.focus(kostl_id, len(performer_cost_center))
                self.session.send_vkey(0)
                continue

            rate_cost_center = (entry.rate_cost_center or performer_cost_center).strip()
            # 单条 Data B 只能对应一个 item，若上游传 "1000;3000" 这种多 item，
            # 取第一个 ";" 之前的部分，保留 SAP POSNR 字段单值约束。
            raw_item = (entry.item or "").strip()
            item_no = raw_item.split(";", 1)[0].strip() if raw_item else ""
            if not performer_cost_center and not rate_cost_center:
                continue
            # Data B 页签中同一行需要同时写执行部门、ZULEISTENDE/KOSTENSAETZE 双表 item 号、费率成本中心和固定价格。
            self.session.set_text(kostl_id, performer_cost_center)
            if item_no and order.sales_group != '240':
                # ZULEISTENDE 表格 item 号写入；缺失和 sales_group 为 240 时跳过，由 SAP 默认行为兜底。
                self.session.set_text(
                    f"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/"
                    f"tblSAPMV45AZULEISTENDE/txtTABL-ZPOSITION[1,{row}]",
                    item_no,
                )
            self.session.set_text(
                f"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/"
                f"tblSAPMV45AKOSTENSAETZE/ctxtTABD-KOSTL[0,{row}]",
                rate_cost_center,
            )
            if item_no and order.sales_group != '240':
                # KOSTENSAETZE 表格 item 号写入；缺失和 sales_group 为 240 时跳过，由 SAP 默认行为兜底。
                self.session.set_text(
                    f"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/"
                    f"tblSAPMV45AKOSTENSAETZE/txtTABD-POSNR[1,{row}]",
                    item_no,
                )
            self.session.set_text(
                f"wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/"
                f"tblSAPMV45AKOSTENSAETZE/txtTABD-FESTPREIS[5,{row}]",
                format(float(entry.amount), ".2f"),
            )

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
        """写入全部 item 行 → 回车 → 重读定位 → 逐条进详情写金额/长文本。

        第二轮**绝不复用第一轮的写入行号**：回车会让 SAP 按 POSNR 升序重排，写入顺序
        与物理行顺序不再等价（Excel 侧 `_sort_items_for_sap` 的预排序只对"全部 item 号
        都是数字且不与已有号冲突"成立，SAP 自动分配号的行不在其保证范围内）。
        故回车后立刻重读概览、按 (item 号, 物料) 解析每条的真实行。
        """
        items = self._resolve_order_items(order)
        sap_amount_total = 0.0
        sap_amount_text = ""

        for row, item in enumerate(items):
            self._write_item_row(row, item)
        # 回车让 SAP 落行并按 POSNR 升序重排——写入时的行号自此失效。
        self.session.send_vkey(0)

        # 重排后立刻重新确认 item 排序，把每个 item 映射到它的真实物理行。
        snapshot = self.read_item_rows()
        claimed: set[int] = set()
        targets: list[tuple[int | None, OrderItemData]] = []
        for item in items:
            row = self.find_item_row(
                item.item, item.material_code, rows=snapshot, skip_rows=claimed
            )
            if row is not None:
                claimed.add(row)
            targets.append((row, item))

        for row, item in targets:
            # 写前再确认一次：条件写入会触发 SAP 重算，行位有可能再次变化。
            row = None if row is None else self.verify_item_row(row, item.item, item.material_code)
            if row is None:
                # 定位不到宁可跳过：写到猜出来的行上会覆盖另一个 item 的金额。
                result.warning = True
                result.append_message(
                    f"item {self._norm_text(item.item) or '(空)'} "
                    f"物料 {self._norm_text(item.material_code)} 未能定位物理行，金额未写入"
                )
                continue
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

    def _write_item_row(self, row: int, item: OrderItemData, *, write_item_no: bool = True) -> None:
        """Write one item row.

        write_item_no=False 时跳过 POSNR 写入，交由 SAP 自动分配号——
        编辑场景新增行且该 item 号已存在时用，避免 POSNR 重号报错。
        """
        if write_item_no and item.item:
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

    @staticmethod
    def _net_value_id(row: int) -> str:
        """Return net value (金额) field id for row —— 概览行第5格，与 item/material 同行。"""
        return (
            "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02/"
            "ssubSUBSCREEN_BODY:SAPMV45A:4415/subSUBSCREEN_TC:SAPMV45A:4902/"
            f"tblSAPMV45ATCTRL_U_ERF_GUTLAST/txtVBAP-NETWR[4,{row}]"
        )

    @staticmethod
    def _auftragswert_id() -> str:
        """Return 订单价值(AUFTRAGSWERT) 字段 id —— 抬头 Data B(T\\14) 页签。"""
        return (
            "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/"
            "ssubSUBSCREEN_BODY:SAPMV45A:4312/txtZAUFTD-AUFTRAGSWERT"
        )

    # ------------------------------------------------------------------ #
    # item 概览行的实时读取与身份定位（创建/编辑共用）
    #
    # SAP 在 item 概览页回车后按 POSNR 升序**强制重排**，写入时的物理行号随即失效
    # （实测：已有 1000/1001/3000/5000 时新增 2000，2000 落到物理行 2，后两条顺延）。
    # 故本节确立统一口径：物理行号只在"同一次快照内"有效，跨越任何 send_vkey 一律作废；
    # 任何"进 item 详情 / 按 item 开计划成本"之前，都必须先重读、再按身份（item 号 + 物料）
    # 定位。定位不到宁可跳过并告警，也绝不写到猜出来的行上——写错行会覆盖别的 item 的金额。
    # ------------------------------------------------------------------ #
    @staticmethod
    def _norm_text(value) -> str:
        """控件文本归一化：None → 空串，其余去首尾空白。行定位的统一比对口径。"""
        return "" if value is None else str(value).strip()

    def read_item_rows(self, max_rows: int = 50) -> list[tuple[int, str, str, str]]:
        """读 item 概览页现有行，返回 [(物理 row, item 号, 物料, 净值)]；空行处停止。

        item / 物料 / 净值同在概览一行（第 1/2/5 格）。净值列读不到时退回空串，
        不阻断扫描——净值只用于消歧兜底，缺失不应让整次定位失败。
        """
        rows: list[tuple[int, str, str, str]] = []
        for row in range(max_rows):
            try:
                item_no = self._norm_text(self.session.read_text(self._item_id(row)))
                material = self._norm_text(self.session.read_text(self._material_id(row)))
            except SapUiError:
                break
            if not item_no and not material:
                break
            try:
                amount = self._norm_text(self.session.read_text(self._net_value_id(row)))
            except SapUiError:
                amount = ""
            rows.append((row, item_no, material, amount))
        return rows

    def find_item_row_by_no(self, item_no) -> int | None:
        """实时重读概览页，返回 item 号等于 item_no 的物理行；找不到返回 None。

        供"按 item 打开计划成本编辑器"使用：ODM 编号与 SAP 实际编号可能不同，
        调用方须传 SAP 侧的真实号。
        """
        target = self._norm_text(item_no)
        if not target:
            return None
        for row, no, _material, _amount in self.read_item_rows():
            if no == target:
                return row
        return None

    def find_item_row(
        self,
        item_no,
        material,
        *,
        rows: list[tuple[int, str, str, str]] | None = None,
        skip_rows: set[int] | None = None,
    ) -> int | None:
        """按 (item 号, 物料) 双键定位物理行；item 号为空时退化为物料匹配。

        Args:
            rows: 已读到的概览快照；不传则实时重读（默认即"实时"）。
            skip_rows: 已被其他 item 认领的行。**仅在同一次快照内有效**——
                跨 send_vkey 复用会重蹈"物理行号当身份"的覆辙。
        """
        snapshot = self.read_item_rows() if rows is None else rows
        target_no = self._norm_text(item_no)
        target_material = self._norm_text(material)
        skip = skip_rows or set()
        for row, no, mat, _amount in snapshot:
            if row in skip or mat != target_material:
                continue
            if not target_no or no == target_no:
                return row
        return None

    def verify_item_row(self, row: int, item_no, material) -> int | None:
        """写前校验：确认 row 上确实是目标 item；不符则实时重定位；仍找不到返回 None。

        这是"写行内数据前必须确认排序"的落地点：即使调用方刚定位过，其间任何一次
        回车都可能让 SAP 重排，故进详情前再确认一次，成本仅两次读取。
        """
        try:
            current_no = self._norm_text(self.session.read_text(self._item_id(row)))
            current_material = self._norm_text(self.session.read_text(self._material_id(row)))
        except SapUiError:
            current_no = current_material = ""
        target_no = self._norm_text(item_no)
        target_material = self._norm_text(material)
        if current_material == target_material and (not target_no or current_no == target_no):
            return row
        return self.find_item_row(item_no, material)

    def read_plan_cost_rows(self, max_rows: int = 50) -> list[tuple[int, str, str, str]]:
        """读计划成本编辑器现有行，返回 [(row, 成本中心, 类别, 金额)]；空行处停止。

        与 read_item_rows 同款口径：编辑器内每次回车后行位都可能变化，调用方须重读。
        """
        rows: list[tuple[int, str, str, str]] = []
        for row in range(max_rows):
            try:
                cost_center = self._norm_text(
                    self.session.read_text(f"wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK2[3,{row}]")
                )
                category = self._norm_text(
                    self.session.read_text(f"wnd[0]/usr/tblSAPLKKDI1301_TC/ctxtRK70L-HERK3[4,{row}]")
                )
            except SapUiError:
                break
            if not cost_center and not category:
                break
            try:
                amount = self._norm_text(
                    self.session.read_text(f"wnd[0]/usr/tblSAPLKKDI1301_TC/txtRK70L-MENGE[6,{row}]")
                )
            except SapUiError:
                amount = ""
            rows.append((row, cost_center, category, amount))
        return rows

    def _sum_item_net_values(self, max_rows: int = 200) -> tuple[float, bool]:
        """读 item 概览各行未税净值(VBAP-NETWR)加和；遇空行(POSNR 为空/读不到)停止。

        金额为单据币种(如 USD/EUR)口径，换算 CNY 由调用方 × 汇率完成。
        max_rows 仅作防跑飞的安全上限；正常终止条件是遇到空行。

        Returns:
            (total, truncated): total 为净值加和；truncated 为 True 表示扫到 max_rows
            上限时各行仍非空——可能有 item 未计入，调用方应据此告警而非静默少算。
        """
        total = 0.0
        for row in range(max_rows):
            try:
                item_no = (self.session.read_text(self._item_id(row)) or "").strip()
            except Exception:
                return total, False
            if not item_no:
                return total, False
            total += self._parse_amount(self.session.read_text(self._net_value_id(row)))
        return total, True

    def fill_order_value(self, order: OrderData) -> SapResult:
        """回填订单价值(AUFTRAGSWERT)：Σ SAP item 未税净值 × 汇率，达阈值才写。

        必须在 item 全部录入(概览页可读 NETWR)后调用。流程：切 item 概览读净值加和 →
        换算 CNY → 进抬头 Data B(T\\14) 页 → ≥ revenue_threshold 时写入。

        与旧"建单头即写 revenue_cny"的错误口径彻底分离，消除双重汇率（详见 fill_order_value 注释）。
        """
        result = SapResult(step="order_value")
        try:
            # 幂等切回 item 概览再读净值，消除对前置步骤(Add Item / Plan Cost)页面状态的依赖。
            self._ensure_item_overview()
            net_total, truncated = self._sum_item_net_values()
            order_value_cny = net_total * (order.exchange_rate or 1.0)

            # 进抬头 Data B 页签写入。
            self.session.press("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD")
            self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14")
            if order_value_cny >= self.config.revenue_threshold:
                self.session.set_text(self._auftragswert_id(), format(order_value_cny, ".2f"))
                result.message = (
                    f"订单价值 {format(order_value_cny, '.2f')} "
                    f"(净值 {format(net_total, '.2f')} × {order.exchange_rate or 1.0})"
                )
            else:
                result.message = (
                    f"订单价值 {format(order_value_cny, '.2f')} < 阈值 "
                    f"{format(self.config.revenue_threshold, '.2f')}，跳过写入"
                )
            # 截断告警放到最后 append，避免被上面的 result.message 直接赋值覆盖。
            if truncated:
                result.warning = True
                result.append_message("item 行数超过扫描上限，订单价值可能少算，请人工核对")
        except Exception as exc:
            return SapResult.fail(f"订单价值回填失败，{exc}", step="order_value")
        return result

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
        """Parse SAP amount text. 空串/非法值统一回退 0.0，避免 SAP 把 0.00 condition 清空回读后触发崩溃。"""
        if not amount_text or not amount_text.strip():
            return 0.0
        try:
            return float(amount_text.replace(",", ""))
        except (TypeError, ValueError):
            return 0.0

    @staticmethod
    def _format_amount(amount: float) -> str:
        """Format amount with thousands separators."""
        return re.sub(r"(\d)(?=(\d\d\d)+(?!\d))", r"\1,", format(amount, ".2f"))

    def apply_plan_cost_entries(
        self,
        entries: list[PlanCostEntry],
        *,
        focus_row: int = 0,
        target_item: str | None = None,
    ) -> SapResult:
        """按已计算好的计划成本明细写入计划成本。

        Args:
            entries: 计划成本明细列表（PlanCostEntry）。
            target_item: 目标 item 号。传入时**按号在概览页实时定位物理行**（推荐口径）——
                SAP 回车后按 POSNR 重排，调用方的列表索引不等于物理行；SAP 无该 item 时
                标 warning 并跳过，绝不开着编辑器写到别的 item 上。
            focus_row: 仅在 target_item 为空（Excel 未提供 item 号，由 SAP 自动分配）时
                使用的兜底行号。

        Returns:
            SapResult: 写入成功或失败信息。
        """
        result = SapResult(step="plan_cost")
        try:
            # 幂等切到 item 概览页：消除对前置步骤（Data B / Add Item / 多 item 循环）页面状态的隐式依赖。
            # 旧实现盲按 btn[3]，多 item 循环第二次会从 item 概览再退一级，触发"保存？"弹窗并把数据写错行。
            self._ensure_item_overview()
            if self._norm_text(target_item):
                # 先重新确认 item 排序，再决定进哪一行的计划成本。
                located = self.find_item_row_by_no(target_item)
                if located is None:
                    result.warning = True
                    result.message = f"SAP 无对应 item {self._norm_text(target_item)}，plan cost 已跳过"
                    return result
                focus_row = located
            self._open_plan_cost_editor(self._material_id(focus_row))
            for entry in entries:
                if not entry.cost_center:
                    continue
                # 每条前重读行数取追加位置：编辑器内每次回车后 SAP 都可能重排/合并行，
                # 用 enumerate 递推会把后续条目写到已有行上。
                self._apply_single_plan_cost_entry(len(self.read_plan_cost_rows()), entry)
            self.session.press("wnd[0]/tbar[0]/btn[3]")
            self.session.press("wnd[1]/usr/btnSPOP-OPTION1")
        except Exception as exc:
            return SapResult.fail(f"plan cost未添加成功，{exc}", step="plan_cost")
        return result

    def _ensure_item_overview(self) -> None:
        """幂等切到 VA02 item 概览页。

        item 概览页的 tab strip ID 是 ``tabsTAXI_TABSTRIP_OVERVIEW``，
        抬头页则是 ``tabsTAXI_TABSTRIP_HEAD``；前者能 find 到即说明当前在 item 概览，可直接返回。
        否则按一次 btn[3] 从抬头页（Data B 完成后的 T\\14 tab）退回 item 概览。
        """
        try:
            self.session.find("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW")
            return
        except SapUiError:
            pass
        # 当前不在 item 概览：可能停在抬头页（刚做完 Data B）或其他子视图，退一级即可。
        self.session.press("wnd[0]/tbar[0]/btn[3]")

    def _open_plan_cost_editor(self, focus_element_id: str) -> None:
        """Open plan cost editor for focused item."""
        self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02")
        # 计划成本菜单依赖当前焦点 item，必须先 setFocus 锁定目标行：
        # ctxtRV45A-MABNR[1,0] 对应 item 1000，[1,1] 对应 item 2000，以此类推。
        # caretPosition 与录制脚本保持一致取 0，避免物料编号短于光标位置时的越界。
        self.session.focus(focus_element_id, 0)
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
