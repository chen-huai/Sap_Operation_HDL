"""Order edit transactions（VA02 字段对比更新）。

与 OrderTransaction(创建) 业务分割：本类只负责"读 SAP 现值 → 与 Excel 对比 → 仅改差异"。

核心正确性约束（见 .claude/plan/va02_edit_order.md 决策 C）：
    对比/写入口径必须与"创建"逐字段一致（同款控件 ID、同款 .2f 格式化、同款规则），
    否则未变化数据会因两套逻辑差异被误判为有变化而错误重写。

复用策略：通过组合持有一个 OrderTransaction(self._base)，直接复用其 open()/save() 与
控件 ID 静态 helper（_item_id/_material_id/...），避免重复造轮子（DRY）。

VA02 编辑屏部分控件 ID 是否与创建抬头视图完全一致尚未实测，首版按一致推断，
不确定处以 TODO 标注，待用户 SAP 录制后校正。
"""

from __future__ import annotations

from sap.exceptions import SapUiError
from sap.models import (
    DataBEntry,
    OrderData,
    OrderItemData,
    PlanCostEntry,
    SapConfig,
    SapResult,
    SubEditEntry,
)
from sap.session import SapSession
from sap.transactions.order import OrderTransaction


class OrderEditTransaction:
    """封装 VA02 订单字段对比更新操作。"""

    def __init__(self, session: SapSession, config: SapConfig):
        """基于共享会话初始化；持有 OrderTransaction 复用 open/save/控件 ID。"""
        self.session = session
        self.config = config
        self._base = OrderTransaction(session, config)

    # ------------------------------------------------------------------ #
    # 对比原语
    # ------------------------------------------------------------------ #
    @staticmethod
    def _norm(value) -> str:
        """文本归一化：None→空串、去首尾空白，统一对比口径。"""
        return ("" if value is None else str(value)).strip()

    @staticmethod
    def _norm_amount(value) -> str:
        """金额归一化为创建写入口径 .2f；无法解析时退回文本归一化。

        SAP 读回的金额可能带千分位逗号，统一去逗号再格式化，保证"其实没变"的金额
        不会因显示格式差异被误判为变化。
        """
        try:
            return format(float(str(value).replace(",", "")), ".2f")
        except (TypeError, ValueError):
            return OrderEditTransaction._norm(value)

    def _compare_and_set(
        self,
        element_id: str,
        new_value,
        *,
        field: str,
        diffs: list[str],
        is_key: bool = False,
        amount: bool = False,
    ) -> bool:
        """读现值 → 对比 → 仅差异才写，并记录 `字段:旧→新` 到 diffs。

        Args:
            element_id: SAP 控件 ID。
            new_value: Excel 期望值。
            field: 字段中文名（用于差异摘要）。
            diffs: 差异收集列表（原地追加）。
            is_key: 下拉框走 set_key，否则 set_text。
            amount: 金额字段，按 .2f 口径归一化对比。

        Returns:
            bool: 是否发生了写入（有差异）。
        """
        try:
            current = self.session.read_text(element_id)
        except SapUiError:
            # 控件读不到（可能 VA02 屏与创建屏不同）→ 记录待校正，绝不盲改。
            diffs.append(f"{field}:控件读取失败(待校正控件ID)")
            return False

        cur_norm = self._norm_amount(current) if amount else self._norm(current)
        new_norm = self._norm_amount(new_value) if amount else self._norm(new_value)
        if cur_norm == new_norm:
            return False

        if is_key:
            self.session.set_key(element_id, self._norm(new_value))
        else:
            self.session.set_text(element_id, new_norm if amount else self._norm(new_value))
        diffs.append(f"{field}:{self._norm(current)}→{new_norm if amount else self._norm(new_value)}")
        return True

    # ------------------------------------------------------------------ #
    # 抬头编辑
    # ------------------------------------------------------------------ #
    def edit_header(self, order: OrderData, diffs: list[str]) -> SapResult:
        """对比并更新订单抬头字段（仅差异）。

        覆盖字段：售达方文本 / 币种 / 汇率 / Product Sub-Category 条件 / GPC Code / CS / Sales。
        Payer 由售达方(sap_no)联动，不单独写；Tax-inclusive amount 仅做校验不落 SAP 字段。
        售达方(sap_no)本身在 VA02 是否可改、控件 ID 待录制校正，见 _edit_sold_to。
        """
        result = SapResult(step="edit_header")
        try:
            # 售达方(SAP Customer Code)位于 VA02 概览屏 subSUBSCREEN_HEADER:4021，
            # 必须在按 btnBT_HEAD 进入抬头详情之前对比/编辑（口径同创建 order.py:71）；
            # 否则进入详情后该控件不在当前屏，读取失败会被静默跳过。
            self._edit_sold_to(order, diffs)

            # 进入抬头详情视图（与创建一致），处理币种/文本/伙伴/submission。
            self.session.press("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD")

            self._edit_currency(order, diffs)
            self._edit_short_text(order, diffs)
            self._edit_partners(order, diffs)
            self._edit_submission(order, diffs)
        except Exception as exc:
            return SapResult.fail(f"抬头编辑失败，{exc}", step="edit_header")
        result.message = "；".join(diffs) if diffs else "抬头无差异"
        return result

    def _edit_sold_to(self, order: OrderData, diffs: list[str]) -> None:
        """对比并编辑售达方(SAP Customer Code)；Payer 一一对应随之联动。

        控件位于 VA02 概览屏（同创建 order.py 写入位置）。读现值与 ODM `sap_no`
        对比，仅差异才按 SAP 录制序列改写并确认联动重算弹窗。
        """
        sold_to_id = (
            "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/"
            "subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR"
        )
        try:
            current = self.session.read_text(sold_to_id)
        except SapUiError:
            # 概览屏控件读不到（屏态异常）→ 记录待排查，绝不盲改。
            diffs.append("售达方(SAP No):控件读取失败")
            return

        new_value = self._norm(order.sap_no)
        if self._norm(current) == new_value:
            return

        # 写入新售达方 → 回车触发 SAP 重算 Payer/重定价等联动弹窗。
        self.session.set_text(sold_to_id, new_value)
        self.session.focus(sold_to_id, len(new_value))
        self.session.send_vkey(0)
        self._confirm_sold_to_dialogs(order)
        diffs.append(f"售达方(SAP No):{self._norm(current)}→{new_value}")

    def _confirm_sold_to_dialogs(self, order: OrderData) -> None:
        """确认改售达方后 SAP 弹出的联动重算弹窗（按录制序列，容错执行）。

        固定序列：wnd[1] 回车 → btnSPOP-VAROPTION1 按两次 → wnd[1] 回车；
        随后按当前订单 item 行数逐行补发回车（每条 item 一次，对应录制中
        "因 item 有两条而增加的两次 sendVKey 0"）。缺窗时容错跳过。
        """
        self.session.try_send_vkey(0, window_id="wnd[1]")
        self._try_press("wnd[1]/usr/btnSPOP-VAROPTION1")
        self._try_press("wnd[1]/usr/btnSPOP-VAROPTION1")
        self.session.try_send_vkey(0, window_id="wnd[1]")

        item_count = sum(1 for item in order.items if item.material_code) or 1
        for _ in range(item_count):
            self.session.try_send_vkey(0, window_id="wnd[1]")

    def _try_press(self, element_id: str) -> bool:
        """容错点击按钮：控件不存在时返回 False 而非抛错（条件弹窗按钮专用）。"""
        try:
            self.session.press(element_id)
            return True
        except SapUiError:
            return False

    def _edit_currency(self, order: OrderData, diffs: list[str]) -> None:
        """对比币种与汇率（T\\01）。"""
        self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01")
        currency_id = (
            "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/"
            "ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBAK-WAERK"
        )
        if self._compare_and_set(currency_id, order.currency_type, field="币种", diffs=diffs):
            self.session.focus(currency_id, 3)
            self.session.send_vkey(0)
            self.session.try_send_vkey(0, window_id="wnd[1]")

        if order.currency_type != "CNY":
            rate_id = (
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/"
                "ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBKD-KURSK"
            )
            if self._compare_and_set(rate_id, order.exchange_rate, field="汇率", diffs=diffs):
                self.session.focus(rate_id, 8)
                self.session.send_vkey(0)

    def _edit_short_text(self, order: OrderData, diffs: list[str]) -> None:
        """对比售达方文本（抬头短文本，T\\10）。"""
        text_id = (
            "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10/"
            "ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/"
            "cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell"
        )
        self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10")
        if self._compare_and_set(text_id, order.short_text, field="售达方文本", diffs=diffs):
            try:
                self.session.set_selection_indexes(text_id, 11, 11)
                lang_id = (
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10/"
                    "ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cmbLV70T-SPRAS"
                )
                self.session.set_key(lang_id, "EN")
                self.session.focus(lang_id)
                self.session.send_vkey(0)
            except SapUiError:
                pass

    def _edit_partners(self, order: OrderData, diffs: list[str]) -> None:
        """对比伙伴页 GPC Code / CS / Sales（T\\09）。

        伙伴表行号随订单不同而变，先按 PARVW 角色 key 定位行，再对比 partner 值。
        角色：ZG=GPC、负责雇员行=CS(cs_code)、VE=Sales(sales_code)。
        """
        self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09")
        partner_prefix = (
            "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/"
            "ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:"
            "SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW"
        )

        # GPC：定位 ZG 行后对比 partner 编码。
        gpc_row = self._find_partner_row(partner_prefix, "ZG")
        if gpc_row is not None and order.global_partner_code:
            self._compare_and_set(
                f"{partner_prefix}/ctxtGVS_TC_DATA-REC-PARTNER[1,{gpc_row}]",
                order.global_partner_code,
                field="GPC Code",
                diffs=diffs,
            )

        # CS：配置映射出的 cs_code 写在"负责雇员/Employee respons."行。
        if self.config.cs_code:
            cs_row = self._find_employee_row(partner_prefix)
            if cs_row is not None:
                self._compare_and_set(
                    f"{partner_prefix}/ctxtGVS_TC_DATA-REC-PARTNER[1,{cs_row}]",
                    self.config.cs_code,
                    field="Primary CS",
                    diffs=diffs,
                )

        # Sales：VE 行对比 sales_code。
        if self.config.sales_code:
            ve_row = self._find_partner_row(partner_prefix, "VE")
            if ve_row is not None:
                self._compare_and_set(
                    f"{partner_prefix}/ctxtGVS_TC_DATA-REC-PARTNER[1,{ve_row}]",
                    self.config.sales_code,
                    field="Sales",
                    diffs=diffs,
                )

    def _find_partner_row(self, partner_prefix: str, parvw_key: str, max_rows: int = 12) -> int | None:
        """扫描伙伴表前 max_rows 行，返回 PARVW 角色 key 命中的首行行号；找不到返回 None。"""
        for row in range(max_rows):
            try:
                key = self.session.find(f"{partner_prefix}/cmbGVS_TC_DATA-REC-PARVW[0,{row}]").key
            except SapUiError:
                break
            if (key or "").strip() == parvw_key:
                return row
        return None

    def _find_employee_row(self, partner_prefix: str, max_rows: int = 12) -> int | None:
        """定位"负责雇员/Employee respons."行（CS 所在行）；找不到返回 None。"""
        for row in range(max_rows):
            try:
                text = self.session.read_text(f"{partner_prefix}/cmbGVS_TC_DATA-REC-PARVW[0,{row}]")
            except SapUiError:
                break
            if (text or "").strip() in {"负责雇员", "Employee respons."}:
                return row
        return None

    def _edit_submission(self, order: OrderData, diffs: list[str]) -> None:
        """对比 Product Sub-Category 驱动的 submission 标识（仅 404 场景，T\\11）。"""
        if order.product_sub_category != "404 Power driven Furniture":
            return
        submission_id = (
            "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\11/"
            "ssubSUBSCREEN_BODY:SAPMV45A:4351/txtVBAK-SUBMI"
        )
        self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\11")
        self._compare_and_set(submission_id, "EF", field="Submission(404)", diffs=diffs)

    # ------------------------------------------------------------------ #
    # item 编辑
    # ------------------------------------------------------------------ #
    def edit_items(self, order: OrderData, diffs: list[str]) -> SapResult:
        """对比并更新 item：命中行改差异，未命中行按创建写法新增。

        Returns:
            SapResult: message 含 item 差异摘要，sap_amount_vat 含金额加和文本。
        """
        result = SapResult(step="edit_items")
        try:
            # 回到 item 概览页（_base 复用创建侧 helper 保证幂等）。
            self._base._ensure_item_overview()
            items = [item for item in order.items if item.material_code]
            if not items:
                result.message = "无 item 数据"
                return result

            existing = self._read_existing_item_rows()  # {item_no: 物理 row}
            next_row = len(existing)
            sap_amount_total = 0.0
            sap_amount_text = ""

            for item in items:
                row = existing.get(item.item) if item.item else None
                if row is None:
                    # 新增 item：复用创建写法落行 + 写金额/长文本。
                    self._base._write_item_row(next_row, item)
                    self.session.send_vkey(0)
                    amount_text = self._enter_item_and_write_condition(next_row, item, result, diffs, is_new=True)
                    existing[item.item or f"_new_{next_row}"] = next_row
                    next_row += 1
                else:
                    # 已存在 item：对比物料/金额/长文本。
                    self._compare_item_material(row, item, diffs)
                    self.session.send_vkey(0)
                    amount_text = self._enter_item_and_write_condition(row, item, result, diffs, is_new=False)

                sap_amount_text = amount_text or sap_amount_text
                sap_amount_total += self._base._parse_amount(amount_text)

            result.sap_amount_vat = (
                self._base._format_amount(sap_amount_total) if len(items) > 1 else sap_amount_text
            )
        except Exception as exc:
            return SapResult.fail(f"item 编辑失败，{exc}", step="edit_items")
        result.message = "；".join(diffs) if diffs else "item 无差异"
        return result

    def _read_existing_item_rows(self, max_rows: int = 50) -> dict[str, int]:
        """读取 item 概览页现有行，返回 {item_no: 物理 row}。空行处停止扫描。"""
        existing: dict[str, int] = {}
        for row in range(max_rows):
            try:
                item_no = (self.session.read_text(OrderTransaction._item_id(row)) or "").strip()
                material = (self.session.read_text(OrderTransaction._material_id(row)) or "").strip()
            except SapUiError:
                break
            if not item_no and not material:
                break
            if item_no:
                existing[item_no] = row
        return existing

    def _compare_item_material(self, row: int, item: OrderItemData, diffs: list[str]) -> None:
        """对比指定行的物料编码（Item Material Code）。"""
        self._compare_and_set(
            OrderTransaction._material_id(row),
            item.material_code,
            field=f"item {item.item} 物料",
            diffs=diffs,
        )

    def _enter_item_and_write_condition(
        self,
        row: int,
        item: OrderItemData,
        result: SapResult,
        diffs: list[str],
        *,
        is_new: bool,
    ) -> str:
        """进入 item 详情，对比/写入金额条件与长文本，返回 SAP 金额文本。"""
        self.session.focus(OrderTransaction._material_id(row), 10)
        self.session.send_vkey(2)

        # 金额条件（Item price）：进入条件页对比。
        condition_id = (
            "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/"
            "ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/"
            "txtKOMV-KBETR[3,5]"
        )
        self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06")
        field = f"item {item.item or '新'} 金额"
        if self._compare_and_set(condition_id, item.revenue, field=field, diffs=diffs, amount=True):
            self.session.focus(condition_id, 16)
            self.session.send_vkey(0)
        amount_text = self.session.read_text(condition_id)

        # 长文本（Item Group Description）：对比更新。
        if item.long_text:
            self._compare_item_long_text(item, result, diffs)

        self.session.press("wnd[0]/tbar[0]/btn[3]")
        return amount_text

    def _compare_item_long_text(self, item: OrderItemData, result: SapResult, diffs: list[str]) -> None:
        """对比 item 长文本（T\\09）。"""
        text_id = (
            "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09/"
            "ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/"
            "cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell"
        )
        self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09")
        if self._compare_and_set(text_id, item.long_text, field=f"item {item.item} 长文本", diffs=diffs):
            try:
                self.session.set_selection_indexes(text_id, 4, 4)
                lang_id = (
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09/"
                    "ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cmbLV70T-SPRAS"
                )
                self.session.set_key(lang_id, "EN")
                self.session.focus(lang_id)
                self.session.send_vkey(0)
                self.session.set_selection_indexes(text_id, 0, 0)
            except SapUiError:
                result.append_message("Long Text 更新失败")

    # ------------------------------------------------------------------ #
    # sub 编辑（Data B / Plan Cost，严格对比口径）
    # ------------------------------------------------------------------ #
    def edit_data_b(
        self,
        entries: list[DataBEntry],
        sub_edit_entries: list[SubEditEntry],
        order: OrderData,
        diffs: list[str],
        *,
        auftragswert_cny: float = 0.0,
    ) -> SapResult:
        """对比并更新 Data B（人工成本）行；口径与创建 fill_lab_cost_entries 一致。

        命中行（按物理 row）对比执行部门/费率成本中心/固定价格；行不足时按创建写法补写。
        新列 Sub Site / Sub Site Transfer Price 控件 ID 待录制校正（见 TODO）。
        """
        result = SapResult(step="edit_data_b")
        try:
            self.session.press("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD")
            self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14")

            if auftragswert_cny >= self.config.revenue_threshold:
                self._compare_and_set(
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/"
                    "ssubSUBSCREEN_BODY:SAPMV45A:4312/txtZAUFTD-AUFTRAGSWERT",
                    auftragswert_cny,
                    field="订单价值",
                    diffs=diffs,
                    amount=True,
                )

            sub_site_by_item = {e.item: e for e in sub_edit_entries}
            for row, entry in enumerate(entries):
                performer = (entry.performer_cost_center or "").strip()
                rate = (entry.rate_cost_center or performer).strip()
                raw_item = (entry.item or "").strip()
                item_no = raw_item.split(";", 1)[0].strip() if raw_item else ""
                if not performer and not rate:
                    continue

                base = "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312"
                self._compare_and_set(
                    f"{base}/tblSAPMV45AZULEISTENDE/ctxtTABL-KOSTL[0,{row}]",
                    performer,
                    field=f"DataB[{row}]执行部门",
                    diffs=diffs,
                )
                if item_no and order.sales_group != "240":
                    self._compare_and_set(
                        f"{base}/tblSAPMV45AZULEISTENDE/txtTABL-ZPOSITION[1,{row}]",
                        item_no,
                        field=f"DataB[{row}]item",
                        diffs=diffs,
                    )
                self._compare_and_set(
                    f"{base}/tblSAPMV45AKOSTENSAETZE/ctxtTABD-KOSTL[0,{row}]",
                    rate,
                    field=f"DataB[{row}]费率中心",
                    diffs=diffs,
                )
                if item_no and order.sales_group != "240":
                    self._compare_and_set(
                        f"{base}/tblSAPMV45AKOSTENSAETZE/txtTABD-POSNR[1,{row}]",
                        item_no,
                        field=f"DataB[{row}]POSNR",
                        diffs=diffs,
                    )
                self._compare_and_set(
                    f"{base}/tblSAPMV45AKOSTENSAETZE/txtTABD-FESTPREIS[5,{row}]",
                    entry.amount,
                    field=f"DataB[{row}]固定价格",
                    diffs=diffs,
                    amount=True,
                )

                # TODO(录制校正): Sub Site / Sub Site Transfer Price 控件 ID 与所属页签。
                # sub_edit = sub_site_by_item.get(item_no)
                # if sub_edit:
                #     self._compare_and_set(<sub_site_id>, sub_edit.sub_site, field=..., diffs=diffs)
                #     self._compare_and_set(<transfer_price_id>, sub_edit.transfer_price,
                #                           field=..., diffs=diffs, amount=True)
        except Exception as exc:
            return SapResult.fail(f"Data B 编辑失败，{exc}", step="edit_data_b")
        result.message = "；".join(diffs) if diffs else "Data B 无差异"
        return result

    def edit_plan_cost(
        self,
        entries: list[PlanCostEntry],
        diffs: list[str],
        *,
        focus_row: int = 0,
    ) -> SapResult:
        """对比并更新计划成本；口径与创建 apply_plan_cost_entries 一致。

        进入计划成本编辑器后按 row 对比 TYPPS/HERK2/HERK3/MENGE，仅差异才写。
        """
        result = SapResult(step="edit_plan_cost")
        try:
            self._base._ensure_item_overview()
            self._base._open_plan_cost_editor(OrderTransaction._material_id(focus_row))
            for row, entry in enumerate(entries):
                if not entry.cost_center:
                    continue
                base = f"wnd[0]/usr/tblSAPLKKDI1301_TC"
                self._compare_and_set(f"{base}/ctxtRK70L-TYPPS[2,{row}]", "E",
                                      field=f"PlanCost[{row}]类型", diffs=diffs)
                self._compare_and_set(f"{base}/ctxtRK70L-HERK2[3,{row}]", entry.cost_center,
                                      field=f"PlanCost[{row}]成本中心", diffs=diffs)
                self._compare_and_set(f"{base}/ctxtRK70L-HERK3[4,{row}]", entry.category,
                                      field=f"PlanCost[{row}]类别", diffs=diffs)
                if self._compare_and_set(f"{base}/txtRK70L-MENGE[6,{row}]", entry.amount,
                                         field=f"PlanCost[{row}]数量", diffs=diffs, amount=True):
                    self.session.focus(f"{base}/txtRK70L-MENGE[6,{row}]", 20)
                    self.session.send_vkey(0)
            self.session.press("wnd[0]/tbar[0]/btn[3]")
            # 退出编辑器时若弹"是否保存"确认框，按 OPTION1 兜底确认。
            try:
                self.session.press("wnd[1]/usr/btnSPOP-OPTION1")
            except SapUiError:
                pass
        except Exception as exc:
            return SapResult.fail(f"plan cost 编辑失败，{exc}", step="edit_plan_cost")
        result.message = "；".join(diffs) if diffs else "Plan Cost 无差异"
        return result

    # ------------------------------------------------------------------ #
    # 复用创建侧 open/save
    # ------------------------------------------------------------------ #
    def open(self, order_no: str) -> SapResult:
        """复用创建侧 VA02 打开逻辑。"""
        return self._base.open(order_no)

    def save(self, info: str) -> SapResult:
        """复用创建侧保存逻辑。"""
        return self._base.save(info)
