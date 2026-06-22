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
            self._edit_partners(order, diffs)
            self._edit_short_text(order, diffs)
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

    def _dismiss_popups(self, max_rounds: int = 3) -> None:
        """连续回车关闭残留的 wnd[1] 确认弹窗，直到无弹窗或达上限。

        SAP 改字段后可能弹出链式确认框（如重新定价提示），残留的模态弹窗会让
        后续 wnd[0] 操作（select_tab 等）抛错并被 edit_header 的 try 吞掉，导致
        "后面方法全不触发"。这里兜底逐个关闭。
        """
        for _ in range(max_rounds):
            if not self.session.try_send_vkey(0, window_id="wnd[1]"):
                return

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
            self._dismiss_popups()

        if order.currency_type != "CNY":
            rate_id = (
                "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\01/"
                "ssubSUBSCREEN_BODY:SAPMV45A:4301/ctxtVBKD-KURSK"
            )
            if self._compare_and_set(rate_id, order.exchange_rate, field="汇率", diffs=diffs):
                self.session.focus(rate_id, 8)
                self.session.send_vkey(0)
                self._dismiss_popups()

        # 兜底：币种/汇率改动后若仍残留确认弹窗，逐个关闭，避免阻塞后续抬头字段编辑。
        self._dismiss_popups()

    def _edit_short_text(self, order: OrderData, diffs: list[str]) -> None:
        """对比售达方文本（抬头短文本，T\\10）。"""
        text_id = (
            "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10/"
            "ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/"
            "cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell"
        )
        self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10")
        if self._compare_and_set(text_id, order.short_text, field="售达方文本", diffs=diffs):
            # 文本本体已由 _compare_and_set 写入；以下语言设置为"尽力而为"。
            # 注意：set_key 设 .key 抛的是原始 COM 异常（非 SapUiError），
            # 这里须用 Exception 兜底，否则语言设置失败会让整个抬头编辑被判失败。
            try:
                self.session.set_selection_indexes(text_id, 11, 11)
                lang_id = (
                    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\10/"
                    "ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cmbLV70T-SPRAS"
                )
                self.session.set_key(lang_id, "EN")
                self.session.focus(lang_id)
                self.session.send_vkey(0)
            except Exception:
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

        # GPC：定位 ZG 行后对比 partner 编码；改动后提交并确认 SAP 校验弹窗。
        gpc_row = self._find_partner_row(partner_prefix, "ZG")
        if gpc_row is not None and order.global_partner_code:
            self._compare_partner_and_confirm(
                f"{partner_prefix}/ctxtGVS_TC_DATA-REC-PARTNER[1,{gpc_row}]",
                order.global_partner_code,
                field="GPC Code",
                diffs=diffs,
            )

        # CS：配置映射出的 cs_code 写在"负责雇员/Employee respons."行；改动后提交并确认。
        if self.config.cs_code:
            cs_row = self._find_employee_row(partner_prefix)
            if cs_row is not None:
                self._compare_partner_and_confirm(
                    f"{partner_prefix}/ctxtGVS_TC_DATA-REC-PARTNER[1,{cs_row}]",
                    self.config.cs_code,
                    field="Primary CS",
                    diffs=diffs,
                )

        # Sales：VE 行对比 sales_code；改动后提交并确认。
        # 销售从无到有时订单原本无 VE 行（创建侧 add_sales_partner 未触发），此时新增一行。
        if self.config.sales_code:
            ve_row = self._find_partner_row(partner_prefix, "VE")
            if ve_row is not None:
                self._compare_partner_and_confirm(
                    f"{partner_prefix}/ctxtGVS_TC_DATA-REC-PARTNER[1,{ve_row}]",
                    self.config.sales_code,
                    field="Sales",
                    diffs=diffs,
                )
            else:
                self._add_partner_row(
                    partner_prefix, "VE", self.config.sales_code, field="Sales", diffs=diffs
                )

    def _compare_partner_and_confirm(
        self, element_id: str, new_value, *, field: str, diffs: list[str]
    ) -> None:
        """对比并写入伙伴行编码；有差异才提交：聚焦→wnd[0]回车触发校验→关闭确认弹窗。

        伙伴字段用 set_text 写入后不会自动校验，须先在 wnd[0] 回车提交，SAP 才会
        弹出确认框（wnd[1]），随后回车关闭。无差异则不提交、不回车。
        """
        if not self._compare_and_set(element_id, new_value, field=field, diffs=diffs):
            return
        self.session.focus(element_id, len(self._norm(new_value)))
        self.session.send_vkey(0)
        self._dismiss_popups()

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

    def _find_empty_partner_row(self, partner_prefix: str, max_rows: int = 12) -> int | None:
        """扫描伙伴表，返回首个空行（PARVW 与 PARTNER 均为空）行号；无空行返回 None。

        用于"角色原本不存在需新增"场景（如 Sales 从无到有）。越界(SapUiError)即停。
        """
        for row in range(max_rows):
            try:
                key = (self.session.find(f"{partner_prefix}/cmbGVS_TC_DATA-REC-PARVW[0,{row}]").key or "").strip()
                partner = (self.session.read_text(f"{partner_prefix}/ctxtGVS_TC_DATA-REC-PARTNER[1,{row}]") or "").strip()
            except SapUiError:
                break
            if not key and not partner:
                return row
        return None

    def _add_partner_row(
        self, partner_prefix: str, parvw_key: str, partner_value, *, field: str, diffs: list[str]
    ) -> None:
        """在伙伴表空行上新增一行：设角色 key + 写编码并提交确认（口径同创建 order.py:189-193）。

        无空行可用时记录待校正、绝不盲写。新增动作记入 diffs（`字段:(空)→新值`）。
        """
        value = self._norm(partner_value)
        if not value:
            return
        row = self._find_empty_partner_row(partner_prefix)
        if row is None:
            diffs.append(f"{field}:无空行可新增(待校正)")
            return
        self.session.set_key(f"{partner_prefix}/cmbGVS_TC_DATA-REC-PARVW[0,{row}]", parvw_key)
        self.session.set_text(f"{partner_prefix}/ctxtGVS_TC_DATA-REC-PARTNER[1,{row}]", value)
        self.session.focus(f"{partner_prefix}/ctxtGVS_TC_DATA-REC-PARTNER[1,{row}]", len(value))
        self.session.send_vkey(0)
        self._dismiss_popups()
        diffs.append(f"{field}:(空)→{value}")

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
        """按 item+物料 双键对比更新 item。

        规则（见 .claude/plan/va02_edit_items_match.md）：
            - item 与 物料 均一致 → 仅更新金额（绝不改写物料，已落盘行物料只读会报错）；
            - item 或 物料 有一个不同 → 新增一条；
            - SAP 有、ODM 表无 → 提示并记 log，不删不改。

        Returns:
            SapResult: message 含 item 差异/提示摘要，sap_amount_vat 含金额加和文本。
        """
        result = SapResult(step="edit_items")
        try:
            # 回到 item 概览页（_base 复用创建侧 helper 保证幂等）。
            self._base._ensure_item_overview()
            items = [item for item in order.items if item.material_code]
            if not items:
                result.message = "无 item 数据"
                return result

            existing = self._read_existing_item_rows()  # [(物理 row, item_no, material, 金额)]
            existing_item_nos = {item_no for _, item_no, _, _ in existing if item_no}
            matched_rows: set[int] = set()
            next_row = len(existing)
            sap_amount_total = 0.0
            sap_amount_text = ""

            for item in items:
                row = self._match_item_row(existing, item, matched_rows)
                if row is not None:
                    # item + 物料 一致 → 进详情逐字段比对金额/长文本，仅差异才写（不碰物料）。
                    # 长文本(Item Group Description)只能在详情页读写，故命中行一律进详情比对。
                    matched_rows.add(row)
                    amount_text, summary = self._enter_item_and_write_condition(row, item, result, is_new=False)
                else:
                    # item 或 物料 不一致 → 新增一条；item 号已存在则让 SAP 自动分配，避免重号。
                    write_item_no = not (self._norm(item.item) and self._norm(item.item) in existing_item_nos)
                    self._base._write_item_row(next_row, item, write_item_no=write_item_no)
                    self.session.send_vkey(0)
                    amount_text, summary = self._enter_item_and_write_condition(next_row, item, result, is_new=True)
                    next_row += 1

                if summary:  # 仅有更新/新增的 item 才输出，无变化静默
                    diffs.append(summary)
                sap_amount_text = amount_text or sap_amount_text
                sap_amount_total += self._base._parse_amount(amount_text)

            # SAP 有但 ODM 表格没有的行 → 每行一条提示，不删不改。
            for row, item_no, material, amount in existing:
                if row not in matched_rows:
                    diffs.append(
                        f"item {item_no} 物料 {material} 金额 {amount}：SAP 有、Excel 无，已跳过"
                    )

            result.sap_amount_vat = (
                self._base._format_amount(sap_amount_total) if len(items) > 1 else sap_amount_text
            )
        except Exception as exc:
            return SapResult.fail(f"item 编辑失败，{exc}", step="edit_items")
        result.message = "；".join(diffs) if diffs else "item 无差异"
        return result

    def _read_existing_item_rows(self, max_rows: int = 50) -> list[tuple[int, str, str, str]]:
        """读取 item 概览页现有行的三要素 item/material/金额。

        item/material/金额同在概览一行（第1/2/5格），返回 [(物理 row, item_no, material, 金额)]。
        空行处停止扫描。金额列读不到时退回空串，不阻断扫描。
        """
        rows: list[tuple[int, str, str, str]] = []
        for row in range(max_rows):
            try:
                item_no = (self.session.read_text(OrderTransaction._item_id(row)) or "").strip()
                material = (self.session.read_text(OrderTransaction._material_id(row)) or "").strip()
            except SapUiError:
                break
            if not item_no and not material:
                break
            try:
                amount = (self.session.read_text(OrderTransaction._net_value_id(row)) or "").strip()
            except SapUiError:
                amount = ""
            rows.append((row, item_no, material, amount))
        return rows

    def _match_item_row(
        self, existing: list[tuple[int, str, str, str]], item: OrderItemData, matched_rows: set[int]
    ) -> int | None:
        """在现有行中找 item_no 与 material 同时一致且未被占用的物理 row；找不到返回 None。"""
        target_item = self._norm(item.item)
        target_material = self._norm(item.material_code)
        for row, item_no, material, _amount in existing:
            if row in matched_rows:
                continue
            if self._norm(item_no) == target_item and self._norm(material) == target_material:
                return row
        return None

    # T\09 item 长文本编辑器控件（仅文本，语言不动——编辑屏语言下拉不可改）。
    _ITEM_LONG_TEXT_ID = (
        "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09/"
        "ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/"
        "cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell"
    )

    def _enter_item_and_write_condition(
        self,
        row: int,
        item: OrderItemData,
        result: SapResult,
        *,
        is_new: bool,
    ) -> tuple[str, str]:
        """进入 item 详情写金额/长文本，返回 (概览权威净值, 单条 item 汇总文本)。

        新增 item 与已存在 item 的条件表布局不同，分别处理：
            - is_new：完全沿用创建侧新增步骤（物料格进详情、价格条件 [3,5] 直接写）；
            - 已存在：编辑对比（数量格进详情、价格条件 [3,1]，仅差异才写）。
        """
        if is_new:
            # 新增 item：与创建侧新增步骤一致（复用 _base 的条件/长文本写法，价格条件 [3,5]）。
            self.session.focus(OrderTransaction._material_id(row), 10)
            self.session.send_vkey(2)
            self._base._write_item_condition(format(item.revenue, ".2f"))
            if item.long_text:
                self._base._write_item_long_text(item.long_text, result)
            summary = self._new_item_summary(item)
        else:
            # 已存在 item：按 SAP 录制聚焦数量格(ZMENG[2,row])双击进详情；价格条件在编辑屏第 1 行
            # KBETR[3,1]（创建侧的 [3,5] 在编辑屏是空行/只读行，写入会抛 Property '.text' can not be set）。
            self.session.focus(OrderTransaction._quantity_id(row), 16)
            self.session.send_vkey(2)
            condition_id = (
                "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/"
                "ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/"
                "txtKOMV-KBETR[3,1]"
            )
            self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06")
            amount_diff = self._diff_set(condition_id, item.revenue, amount=True)
            if amount_diff and amount_diff[0]:
                self.session.focus(condition_id, 16)
                self.session.send_vkey(0)

            # 长文本：仅写文本本身，不动语言（编辑屏 cmbLV70T-SPRAS 不可改，set_key 会抛
            # Property '.key' can not be set，且为原始 COM 异常）。
            text_diff = None
            if item.long_text:
                self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09")
                text_diff = self._diff_set(self._ITEM_LONG_TEXT_ID, item.long_text)

            summary = self._matched_item_summary(item, amount_diff, text_diff)

        # 返回概览页后，从同一行第5格（NETWR）读取权威净值金额；条件页 KBETR 对部分单读不到。
        self.session.press("wnd[0]/tbar[0]/btn[3]")
        try:
            amount_text = self.session.read_text(OrderTransaction._net_value_id(row))
        except SapUiError:
            amount_text = ""
        return amount_text, summary

    # ------------------------------------------------------------------ #
    # item 差异原语与单行汇总
    # ------------------------------------------------------------------ #
    def _diff_set(self, element_id: str, new_value, *, amount: bool = False) -> tuple[bool, str, str] | None:
        """读现值→对比→仅差异才写。返回 (是否变化, 旧值, 新值)；控件读不到返回 None。"""
        try:
            current = self.session.read_text(element_id)
        except SapUiError:
            return None
        old = self._norm(current)
        if amount:
            cur_cmp, new_disp = self._norm_amount(current), self._norm_amount(new_value)
        else:
            cur_cmp, new_disp = old, self._norm(new_value)
        if cur_cmp == (self._norm_amount(new_value) if amount else self._norm(new_value)):
            return (False, old, new_disp)
        self.session.set_text(element_id, new_disp)
        return (True, old, new_disp)

    def _matched_item_summary(
        self, item: OrderItemData, amount_diff: tuple[bool, str, str] | None,
        text_diff: tuple[bool, str, str] | None,
    ) -> str:
        """命中 item 的单行汇总：只列真正变化的字段；无变化返回空串(调用方不输出)。

        例：`item 10 物料 ABC 金额 500 → 400.00`。
        """
        parts: list[str] = []
        if amount_diff is None:
            parts.append("金额读取失败")
        elif amount_diff[0]:
            parts.append(f"金额 {amount_diff[1]} → {amount_diff[2]}")
        if item.long_text:
            if text_diff is None:
                parts.append("文本读取失败")
            elif text_diff[0]:
                parts.append(f"文本 {text_diff[1]} → {text_diff[2]}")
        if not parts:
            return ""
        head = f"item {self._norm(item.item)} 物料 {self._norm(item.material_code)}"
        return f"{head} " + "，".join(parts)

    def _new_item_summary(self, item: OrderItemData) -> str:
        """新增 item 的单行汇总：`item X 物料 Y 新增金额 V[，文本 T]`。"""
        head = f"item {self._norm(item.item) or '新'} 物料 {self._norm(item.material_code)}"
        parts = [f"新增金额 {format(item.revenue, '.2f')}"]
        if item.long_text:
            parts.append(f"文本 {self._norm(item.long_text)}")
        return f"{head} " + "，".join(parts)

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

    _PLAN_COST_TABLE = "wnd[0]/usr/tblSAPLKKDI1301_TC"

    def edit_plan_cost(
        self,
        entries: list[PlanCostEntry],
        diffs: list[str],
        *,
        focus_row: int = 0,
    ) -> SapResult:
        """按 成本中心+类别 匹配更新计划成本（每条 entry 一行汇总）。

        规则（同 item 编辑）：
            - 成本中心 与 类别 均一致 → 仅更新不同的金额/时间(MENGE)，不碰类型/中心/类别；
            - 成本中心 或 类别 有一个不同 → 新增一行（写全 TYPPS/HERK2/HERK3/MENGE，同创建）；
            - SAP 有、ODM 表无 → 记录提示，不删不改。
        """
        result = SapResult(step="edit_plan_cost")
        try:
            self._base._ensure_item_overview()
            self._open_plan_cost_editor_for_edit(OrderTransaction._material_id(focus_row))

            valid = [e for e in entries if e.cost_center]
            existing = self._read_existing_plan_cost_rows()  # [(row, cost_center, category, amount)]
            matched_rows: set[int] = set()
            next_row = len(existing)

            for entry in valid:
                row = self._match_plan_cost_row(existing, entry, matched_rows)
                if row is not None:
                    # 成本中心+类别一致 → 仅更新金额/时间(MENGE)，无变化不输出。
                    matched_rows.add(row)
                    menge_id = f"{self._PLAN_COST_TABLE}/txtRK70L-MENGE[6,{row}]"
                    amount_diff = self._diff_set(menge_id, entry.amount, amount=True)
                    if amount_diff is None:
                        diffs.append(f"{self._plan_cost_head(entry.cost_center, entry.category)} 数量读取失败")
                    elif amount_diff[0]:
                        self.session.focus(menge_id, 20)
                        self.session.send_vkey(0)
                        diffs.append(self._plan_cost_changed_summary(entry, amount_diff))
                else:
                    # 成本中心或类别不同 → 新增一行（复用创建侧单行写法）。
                    self._base._apply_single_plan_cost_entry(next_row, entry)
                    diffs.append(self._plan_cost_new_summary(entry))
                    next_row += 1

            # SAP 有但 ODM 表格没有的计划成本行 → 每行一条提示，不删不改。
            for row, cost_center, category, amount in existing:
                if row not in matched_rows:
                    label = self._plan_cost_label(category)
                    diffs.append(
                        f"{self._plan_cost_head(cost_center, category)} {label} {amount}：SAP 有、Excel 无，已跳过"
                    )

            self.session.press("wnd[0]/tbar[0]/btn[3]")
            # 退出编辑器时若弹"是否保存"确认框，按 OPTION1 兜底确认（无则跳过）。
            self._try_press("wnd[1]/usr/btnSPOP-OPTION1")
        except Exception as exc:
            return SapResult.fail(f"plan cost 编辑失败，{exc}", step="edit_plan_cost")
        result.message = "；".join(diffs) if diffs else "Plan Cost 无差异"
        return result

    def _open_plan_cost_editor_for_edit(self, focus_element_id: str) -> None:
        """编辑场景打开计划成本编辑器：容错处理"选择计算变式"弹窗。

        已有计划成本的 item 再次进入编辑器通常不弹 btnSPOP-VAROPTION1 选择框，
        故对该弹窗按钮用 _try_press 容错（缺失即跳过），避免硬按报"找不到 SAP 元素"。
        """
        self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\02")
        self.session.focus(focus_element_id, 0)
        self.session.find("wnd[0]/mbar/menu[3]/menu[7]").select()
        self._try_press("wnd[1]/usr/btnSPOP-VAROPTION1")
        self._try_press("wnd[1]/tbar[0]/btn[0]")

    def _read_existing_plan_cost_rows(self, max_rows: int = 50) -> list[tuple[int, str, str, str]]:
        """读取计划成本编辑器现有行，返回 [(row, cost_center, category, amount)]。空行处停止。"""
        rows: list[tuple[int, str, str, str]] = []
        for row in range(max_rows):
            try:
                cost_center = (self.session.read_text(f"{self._PLAN_COST_TABLE}/ctxtRK70L-HERK2[3,{row}]") or "").strip()
                category = (self.session.read_text(f"{self._PLAN_COST_TABLE}/ctxtRK70L-HERK3[4,{row}]") or "").strip()
            except SapUiError:
                break
            if not cost_center and not category:
                break
            try:
                amount = (self.session.read_text(f"{self._PLAN_COST_TABLE}/txtRK70L-MENGE[6,{row}]") or "").strip()
            except SapUiError:
                amount = ""
            rows.append((row, cost_center, category, amount))
        return rows

    def _match_plan_cost_row(
        self, existing: list[tuple[int, str, str, str]], entry: PlanCostEntry, matched_rows: set[int]
    ) -> int | None:
        """找 成本中心 与 类别 同时一致且未占用的行；找不到返回 None。"""
        cc, cat = self._norm(entry.cost_center), self._norm(entry.category)
        for row, cost_center, category, _amount in existing:
            if row in matched_rows:
                continue
            if self._norm(cost_center) == cc and self._norm(category) == cat:
                return row
        return None

    def _plan_cost_label(self, category) -> str:
        """计划成本数量列标签：T01AST(工时)→"时间"，其余→"金额"。"""
        return "时间" if self._norm(category) == "T01AST" else "金额"

    def _plan_cost_head(self, cost_center, category) -> str:
        """计划成本行抬头：`计划成本 成本中心{中心}({类别})`。"""
        return f"计划成本 成本中心{self._norm(cost_center)}({self._norm(category)})"

    def _plan_cost_changed_summary(
        self, entry: PlanCostEntry, amount_diff: tuple[bool, str, str]
    ) -> str:
        """命中且有变化：`计划成本 成本中心X(类别) 金额|时间 旧 → 新`。"""
        _changed, old, new = amount_diff
        label = self._plan_cost_label(entry.category)
        return f"{self._plan_cost_head(entry.cost_center, entry.category)} {label} {old} → {new}"

    def _plan_cost_new_summary(self, entry: PlanCostEntry) -> str:
        """新增行：`计划成本 成本中心X(类别) 新增金额|时间 值`。"""
        label = self._plan_cost_label(entry.category)
        head = self._plan_cost_head(entry.cost_center, entry.category)
        return f"{head} 新增{label} {format(float(entry.amount), '.2f')}"

    # ------------------------------------------------------------------ #
    # 复用创建侧 open/save
    # ------------------------------------------------------------------ #
    def open(self, order_no: str) -> SapResult:
        """复用创建侧 VA02 打开逻辑。"""
        return self._base.open(order_no)

    def save(self, info: str) -> SapResult:
        """复用创建侧保存逻辑。"""
        return self._base.save(info)
