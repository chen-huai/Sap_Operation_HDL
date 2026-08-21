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

from collections import Counter
from dataclasses import replace

from sap.exceptions import SapUiError
from sap.models import (
    DataBEntry,
    ItemAddInfo,
    OrderData,
    OrderItemData,
    PlanCostEntry,
    SapConfig,
    SapResult,
)
from sap.rules import resolve_data_a_key, should_fill_ic_transaction
from sap.session import SapSession
from sap.transactions.order import OrderTransaction


class OrderEditTransaction:
    """封装 VA02 订单字段对比更新操作。"""

    # 售达方/付款方联动重算弹窗排空的最大轮数（死循环兜底，见 _confirm_sold_to_dialogs）。
    _MAX_DIALOG_ROUNDS = 8

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
        """文本归一化：None→空串、换行统一为 \\n、去首尾空白，统一对比口径。

        换行归一是必需的：SAP 多行文本控件（售达方文本 T\\10、item 长文本）回读用 \\r\\n，
        Excel 侧是 \\n，只 strip 首尾会让"一字不差"的文本每次都判为差异并重写（实测已确认）。
        """
        if value is None:
            return ""
        return str(value).replace("\r\n", "\n").replace("\r", "\n").strip()

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

    @staticmethod
    def _norm_number(value) -> str:
        """数值归一化：去千分位后按 float 值比对，尾随零不影响；无法解析时退回文本归一化。

        与 _norm_amount 的区别：不强制 .2f——汇率等字段需要保留原始精度语义。
        SAP 回读汇率带尾随零（`7.10000`）而 Excel 是 `7.1`，文本比对会每次误判为差异、
        重写汇率并连带触发币种/汇率联动弹窗排空（实测已确认）。
        """
        try:
            return repr(float(str(value).replace(",", "")))
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
        numeric: bool = False,
        prefix_field: bool = True,
    ) -> bool:
        """读现值 → 对比 → 仅差异才写，并记录 `字段:旧→新` 到 diffs（空值统一显示 `(空)`）。

        Args:
            element_id: SAP 控件 ID。
            new_value: Excel 期望值。
            field: 字段中文名（用于差异摘要）。
            diffs: 差异收集列表（原地追加）。
            is_key: 下拉框走 read_key/set_key（key 与显示文本不同口径），否则 read_text/set_text。
            amount: 金额字段，按 .2f 口径归一化对比。
            numeric: 数值字段（如汇率），按 float 值比对（尾随零不算差异），写入仍用原值。
            prefix_field: 摘要是否带 `字段:` 前缀。多字段共段（如 Header）需要区分故为 True；
                单字段独立成段（如订单价值，段名已由 mixin 提供）传 False 避免 `订单价值:订单价值:` 双重前缀。

        Returns:
            bool: 是否发生了写入（有差异）。
        """
        try:
            current = self.session.read_key(element_id) if is_key else self.session.read_text(element_id)
        except SapUiError:
            # 控件读不到（可能 VA02 屏与创建屏不同）→ 记录待校正，绝不盲改。
            diffs.append(f"{field}:控件读取失败(待校正控件ID)")
            return False

        if amount:
            cur_norm, new_norm = self._norm_amount(current), self._norm_amount(new_value)
        elif numeric:
            # 只用于"是否相等"的判定；写入与摘要仍用原值，避免把 7.1 写成 repr 形态。
            cur_norm, new_norm = self._norm_number(current), self._norm_number(new_value)
        else:
            cur_norm, new_norm = self._norm(current), self._norm(new_value)
        if cur_norm == new_norm:
            return False

        written = new_norm if amount else self._norm(new_value)
        if is_key:
            self.session.set_key(element_id, self._norm(new_value))
        else:
            self.session.set_text(element_id, written)
        # 空值统一渲染 `(空)`，避免 `订单价值:3,021.26→` 结尾空白看不出目标被清空。
        label = f"{field}:" if prefix_field else ""
        old_disp = self._norm(current) or "(空)"
        new_disp = written or "(空)"
        diffs.append(f"{label}{old_disp}→{new_disp}")
        return True

    # ------------------------------------------------------------------ #
    # 抬头编辑
    # ------------------------------------------------------------------ #
    def edit_header(self, order: OrderData, diffs: list[str]) -> SapResult:
        """对比并更新订单抬头字段（仅差异）。

        覆盖字段：售达方文本 / 币种 / 汇率 / Product Sub-Category 条件 / GPC Code / CS / Sales /
        DATA A 客户组(T\\13) / ECD 与 IC 交易类型(T\\14)。
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
            self._edit_data_a(order, diffs)
            self._edit_data_b_header(order, diffs)
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
        self._confirm_sold_to_dialogs()
        diffs.append(f"售达方(SAP No):{self._norm(current) or '(空)'}→{new_value or '(空)'}")

    def _confirm_sold_to_dialogs(self) -> None:
        """动态排空改售达方后 SAP 弹出的联动重算弹窗。

        弹窗数量不是 item 行数的函数，而是随本次变更连带触发的重算种类而变：
        付款方联动（固定）+ 是否改汇率（条件性，如汇率变会多一个确认框）+ 是否重定价。
        故不再按固定次数/item 数硬发回车，改为逐轮排空：

            每轮 → 选项框 btnSPOP-VAROPTION1 存在则按按钮（禁止用回车替代，
                    否则会静默选中默认项、选错还不报错）；
                  → 否则普通 wnd[1] 确认框回车；
                  → 都没有则结束。

        缺窗时 _try_press/try_send_vkey 幂等返回 False，天然给出退出信号；
        再以 _MAX_DIALOG_ROUNDS 兜底，杜绝异常弹窗导致死循环。
        """
        for _ in range(self._MAX_DIALOG_ROUNDS):
            if self._try_press("wnd[1]/usr/btnSPOP-VAROPTION1"):
                continue
            if self.session.try_send_vkey(0, window_id="wnd[1]"):
                continue
            return

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
            # numeric 比对：SAP 回读 7.10000、Excel 是 7.1，文本比对会每次重写汇率并触发联动弹窗。
            if self._compare_and_set(
                rate_id, order.exchange_rate, field="汇率", diffs=diffs, numeric=True
            ):
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

    def _edit_data_a(self, order: OrderData, diffs: list[str]) -> None:
        """对比 DATA A 客户组（T\\13 cmbVBAK-KVGR1），判定复用创建同一条规则。

        客户号命中 Data_A_E1 → E1、Data_A_Z2 → Z2，否则兜底 00；客户改动后该值随之变化，
        故编辑时必须重算回填，否则保留旧客户的分组。下拉框按 key 对比（见 _compare_and_set）。
        """
        self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13")
        self._compare_and_set(
            "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\13/"
            "ssubSUBSCREEN_BODY:SAPMV45A:4309/cmbVBAK-KVGR1",
            resolve_data_a_key(order, self.config),
            field="Data A",
            diffs=diffs,
            is_key=True,
        )

    def _edit_data_b_header(self, order: OrderData, diffs: list[str]) -> None:
        """对比 T\\14 抬头两项：ECD 与 IC 交易类型（口径同创建 fill_header）。

        - ECD(VORAUS_AUFENDE)：仅 Excel 有值才同步。Excel 缺 ECD 属数据不全，不视为"要清空"。
        - IC 交易类型(IC_TRANSAKTION)：命中 Data_B_TUV → O1，未命中 → 清空。config 为权威，
          客户从清单移除后 SAP 上的 O1 须随之撤销，故走全量对比覆盖而非"命中才写"。
        """
        base = "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312"
        self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14")

        if order.ecd:
            self._compare_and_set(
                f"{base}/ctxtZAUFTD-VORAUS_AUFENDE", order.ecd, field="ECD", diffs=diffs,
            )

        self._compare_and_set(
            f"{base}/ctxtZAUFTD-IC_TRANSAKTION",
            "O1" if should_fill_ic_transaction(order, self.config) else "",
            field="IC交易类型",
            diffs=diffs,
        )

    # ------------------------------------------------------------------ #
    # item 编辑
    # ------------------------------------------------------------------ #
    def edit_items(
        self,
        order: OrderData,
        diffs: list[str],
        added_out: list[ItemAddInfo] | None = None,
    ) -> SapResult:
        """按 item+物料 双键对比更新 item。

        规则（见 .claude/plan/va02_edit_items_match.md）：
            - item 与 物料 均一致 → 仅更新金额（绝不改写物料，已落盘行物料只读会报错）；
            - item 或 物料 有一个不同 → 新增一条；
            - SAP 有、ODM 表无 → 提示并记 log，不删不改。

        Args:
            added_out: 若传入，则把本次发生新增的 item 明细（ItemAddInfo）原地追加，
                供调用方 save+open 后 build_item_no_mapping 建立 ODM→SAP 号映射
                （见 .claude/plan/va02_edit_item_add_midsave.md）。

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

            # 分两批处理，规避"新增回车后 SAP 按 POSNR 重排导致行号失效"：
            #   ① 命中项：改金额/长文本，POSNR 不变、不触发重排，行号稳定；
            #   ② 新增项：每加一条 SAP 都会重排，故回车后必须重读概览重新定位当前行，
            #      绝不能沿用写入时的追加行号（否则金额/文本写到别的 item 上）。
            matched: list[tuple[int, "OrderItemData"]] = []
            new_items: list["OrderItemData"] = []
            matched_rows: set[int] = set()
            for item in items:
                row = self._match_item_row(existing, item, matched_rows)
                if row is not None:
                    matched_rows.add(row)
                    matched.append((row, item))
                else:
                    new_items.append(item)

            # 批次一：命中项（进详情逐字段比对金额/长文本，仅差异才写，绝不碰只读物料）。
            for row, item in matched:
                _amount_text, summary = self._enter_item_and_write_condition(
                    row, item, result, is_new=False, currency=order.currency_type
                )
                if summary:
                    diffs.append(summary)

            # 批次二：新增项。写入末尾行 → 回车（SAP 重排）→ 重读定位当前行 → 写金额/文本。
            known_item_nos = set(existing_item_nos)
            used_rows = set(matched_rows)
            for item in new_items:
                # item 号已存在则让 SAP 自动分配（write_item_no=False），避免重号。
                write_item_no = not (self._norm(item.item) and self._norm(item.item) in known_item_nos)
                append_row = len(self._read_existing_item_rows())
                self._base._write_item_row(append_row, item, write_item_no=write_item_no)
                self.session.send_vkey(0)
                actual_row = self._relocate_new_item_row(
                    item, write_item_no=write_item_no,
                    known_item_nos=known_item_nos, used_rows=used_rows,
                )
                if actual_row is None:
                    actual_row = append_row  # 兜底：重读定位失败时退回写入行
                amount_text, summary = self._enter_item_and_write_condition(actual_row, item, result, is_new=True)
                used_rows.add(actual_row)
                # 记录新增明细：write_item_no=False 时 SAP 自动改号，需 save+open 后回读映射。
                if added_out is not None:
                    added_out.append(ItemAddInfo(
                        odm_item=self._norm(item.item),
                        material=self._norm(item.material_code),
                        amount=self._norm(amount_text),
                        auto_numbered=not write_item_no,
                    ))
                # 重排后刷新已知号集合，供下一条新增的号冲突判定与定位。
                known_item_nos = {n for _, n, _, _ in self._read_existing_item_rows() if n}
                if summary:
                    diffs.append(summary)

            # SAP 有但 ODM 表格没有的行 → 每行一条提示，不删不改（用初始概览与 matched_rows，二者同基准一致）。
            for row, item_no, material, amount in existing:
                if row not in matched_rows:
                    diffs.append(
                        f"item {item_no} 物料 {material} 金额 {amount}：SAP 有、Excel 无，已跳过"
                    )

            # 未税加和：全部 item 写完后回概览页重读一次全量净值(Σ VBAP-NETWR)，与
            # edit_order_value / 创建路径 fill_order_value 完全同源。不用"边写边累加"的中间值：
            #   ① 那会漏掉"SAP 有、Excel 无"的 item，加和小于 SAP 订单真实总额；
            #   ② 先前 item 的净值会被后续新增触发的 SAP 重排/重算改变，累加值不会回头更新；
            #   ③ 单 item 曾走"取最后一次读到的文本"分支，读不到时直接落空串 → 校验按 0 比对。
            self._base._ensure_item_overview()
            net_total, truncated = self._base._sum_item_net_values()
            result.sap_amount_vat = self._base._format_amount(net_total)
            if truncated:
                result.warning = True
                diffs.append("item 行数超过扫描上限，未税加和可能少算，请人工核对")
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

    def _relocate_new_item_row(
        self, item: OrderItemData, *, write_item_no: bool, known_item_nos: set[str], used_rows: set[int]
    ) -> int | None:
        """回车重排后重读概览，定位刚新增 item 的当前物理行。

        SAP VA02 概览页在回车确认后按 POSNR 升序重排，写入时的追加行号随即失效，
        必须按 item 身份重新定位：
            - write_item_no=True：写了 POSNR，新行 item 号 == ODM 号，按号定位；
            - write_item_no=False：SAP 自动改号，按"物料一致 且 item 号此前未出现"定位。
        找不到返回 None（调用方退回写入行兜底）。
        """
        rows = self._read_existing_item_rows()
        target_item = self._norm(item.item)
        target_material = self._norm(item.material_code)
        if write_item_no and target_item:
            for row, item_no, _material, _amount in rows:
                if row not in used_rows and self._norm(item_no) == target_item:
                    return row
        for row, item_no, material, _amount in rows:
            if row in used_rows:
                continue
            if self._norm(material) == target_material and self._norm(item_no) not in known_item_nos:
                return row
        return None

    def _find_item_physical_row(self, target_item) -> int | None:
        """在 SAP item 概览页找 item 号等于 target_item 的物理 row；找不到返回 None。

        用于计划成本按 item 定位（ODM item 编号与 SAP 实际编号可能不同，需精确匹配）。
        """
        target = self._norm(target_item)
        if not target:
            return None
        for row, item_no, _material, _amount in self._read_existing_item_rows():
            if self._norm(item_no) == target:
                return row
        return None

    def build_item_no_mapping(self, added: list[ItemAddInfo]) -> dict[str, str]:
        """中途 save+open 后，为新增 item 建立 ODM→SAP 号映射。

        调用前提：调用方已 save 使新增 item 落库、并重进订单（/NVA02）。此处读回
        当前 item 概览页，把每个新增项关联到其 SAP 实际号：
            - auto_numbered=False：ODM 号即 SAP 号 → 恒等映射，并预占该 SAP 行避免被抢；
            - auto_numbered=True：SAP 自动改号 → 按物料匹配、金额兜底、出现顺序兜底定位。

        无匹配的新增项退回恒等映射并记 log（下游 plan cost 会 warning 跳过，不静默）。

        Returns:
            dict[str, str]: {ODM item 号: SAP 实际 item 号}；仅含新增项，未新增 item 的
                下游定位直接用原号（调用方 .get(odm, odm) 兜底）。
        """
        mapping: dict[str, str] = {}
        self._base._ensure_item_overview()
        existing = self._read_existing_item_rows()  # [(row, item_no, material, amount)]
        used_rows: set[int] = set()

        # 先处理恒等映射项：预占其 SAP 行，防止同物料的改号新增项误抢。
        identity_items = {info.odm_item for info in added if not info.auto_numbered and info.odm_item}
        for row, item_no, _material, _amount in existing:
            if self._norm(item_no) in identity_items:
                used_rows.add(row)

        for info in added:
            if not info.auto_numbered:
                if info.odm_item:
                    mapping[info.odm_item] = info.odm_item
                continue
            sap_item_no = self._match_added_row(existing, info, used_rows)
            if sap_item_no is None:
                # 无法定位改号后的行：退回恒等，交由下游 warning 跳过（不静默丢失）。
                mapping[info.odm_item] = info.odm_item
            else:
                mapping[info.odm_item] = sap_item_no
        return mapping

    def _match_added_row(
        self, existing: list[tuple[int, str, str, str]], info: ItemAddInfo, used_rows: set[int]
    ) -> str | None:
        """为一个改号新增项在概览行中定位 SAP 实际号：物料匹配→金额兜底→顺序兜底。

        命中后把该物理行标记为已用（同物料多条新增按出现顺序逐一消费），返回 SAP item 号；
        物料无任何未占用行匹配时返回 None。
        """
        material = self._norm(info.material)
        # 仅接受 item 号非空的行——空号行（新增行号尚未渲染等）不能作为映射目标，否则会把
        # ODM 号映射成空串，导致下游 Data B 的 POSNR/ZPOSITION 被写空。
        candidates = [
            (row, item_no, amount)
            for row, item_no, mat, amount in existing
            if row not in used_rows and self._norm(mat) == material and self._norm(item_no)
        ]
        if not candidates:
            return None

        # 金额兜底：同物料多条时优先净值一致的行。
        if info.amount:
            want = self._norm_amount(info.amount)
            for row, item_no, amount in candidates:
                if self._norm_amount(amount) == want:
                    used_rows.add(row)
                    return item_no

        # 顺序兜底：取首个未占用的同物料行。
        row, item_no, _amount = candidates[0]
        used_rows.add(row)
        return item_no

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
        currency: str | None = None,
    ) -> tuple[str, str]:
        """进入 item 详情写金额/币种/长文本，返回 (概览权威净值, 单条 item 汇总文本)。

        新增 item 与已存在 item 的条件表布局不同，分别处理：
            - is_new：完全沿用创建侧新增步骤（物料格进详情、价格条件 [3,5] 直接写，
              币种自动继承单据币种，无需写）；
            - 已存在：编辑对比（数量格进详情、价格条件 [3,1]，仅差异才写）；单据币种变了，
              item 条件币种 KOEIN[4,1] 也须随抬头同步（currency=order.currency_type）。
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
            condition_base = (
                "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/"
                "ssubSUBSCREEN_BODY:SAPLV69A:6201/tblSAPLV69ATCTRL_KONDITIONEN/"
            )
            condition_id = f"{condition_base}txtKOMV-KBETR[3,1]"
            self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06")
            amount_diff = self._diff_set(condition_id, item.revenue, amount=True)
            # 单据币种变了，同一条件行的币种列 KOEIN[4,1] 也随抬头同步（仅差异才写）。
            # 仅在能读到现有币种值时才对比：已定价 item 的 KOEIN 恒有币种，读到空串
            # 说明控件异常/无条件行，此时不盲写（口径同 _edit_sold_to 读不到即跳过）。
            currency_diff = None
            if currency is not None:
                currency_id = f"{condition_base}ctxtRV61A-KOEIN[4,1]"
                try:
                    current_currency = self._norm(self.session.read_text(currency_id))
                except SapUiError:
                    current_currency = ""
                if current_currency:
                    currency_diff = self._diff_set(currency_id, currency)
            # 金额或币种任一变化才回车提交整屏（只改币种时金额数字不变，仍须提交）。
            if (amount_diff and amount_diff[0]) or (currency_diff and currency_diff[0]):
                self.session.focus(condition_id, 16)
                self.session.send_vkey(0)

            # 长文本：仅写文本本身，不动语言（编辑屏 cmbLV70T-SPRAS 不可改，set_key 会抛
            # Property '.key' can not be set，且为原始 COM 异常）。
            text_diff = None
            if item.long_text:
                self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09")
                text_diff = self._diff_set(self._ITEM_LONG_TEXT_ID, item.long_text)

            summary = self._matched_item_summary(item, amount_diff, text_diff, currency_diff)

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
        currency_diff: tuple[bool, str, str] | None = None,
    ) -> str:
        """命中 item 的单行汇总：只列真正变化的字段；无变化返回空串(调用方不输出)。

        箭头格式与全局统一：`旧→新`（无空格），空值渲染 `(空)`。
        例：`item 10 物料 ABC 金额 500.00→400.00，币种 USD→CNY`。
        """
        def _arrow(diff: tuple[bool, str, str]) -> str:
            return f"{diff[1] or '(空)'}→{diff[2] or '(空)'}"

        parts: list[str] = []
        if amount_diff is None:
            parts.append("金额读取失败")
        elif amount_diff[0]:
            parts.append(f"金额 {_arrow(amount_diff)}")
        if currency_diff and currency_diff[0]:
            parts.append(f"币种 {_arrow(currency_diff)}")
        if item.long_text:
            if text_diff is None:
                parts.append("文本读取失败")
            elif text_diff[0]:
                parts.append(f"文本 {_arrow(text_diff)}")
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
    def clear_data_b(
        self,
        entries: list[DataBEntry],
        order: OrderData,
        diffs: list[str],
        item_no_map: dict[str, str] | None = None,
    ) -> SapResult:
        """Data B 同步第一段：与 Excel 对比，有差异则两阶段删空（不写入任何数据）。

        与 write_data_b() 之间**必须隔一次 save + open_order**，这是 SAP 硬约束：
        成本表(KOSTENSAETZE) 每行都是对执行部门表(ZULEISTENDE) 成本中心的引用，SAP 校验
        "该费率行成本中心是否为本订单贡献成本中心"，同一 dialog 内不认未落库的新增行。
        删除已消耗一次 PAI 往返 → 此后同屏写成本表必报 ZR520
        (No contributing cost centre exists for the order specific hourly rate)。
        删除单独提交后表回到干净空状态，重进再写即等价于创建路径(VA01)，实测通过。

        一致性短路：SAP 与 Excel 内容一致时不删不写，返回 changed=False，调用方据此免掉
        中途保存（纯更新且无变化的订单零额外开销）。

        Returns:
            SapResult: changed=True 表示已删空、待 write_data_b 重建；changed=False 表示无差异跳过。
        """
        base = "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312"
        try:
            self.session.press("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD")
            self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14")

            # 对比用已 remap 的 entries：POSNR 存的是 SAP 真实号。映射此刻可能尚未建立
            # （改号 item 的真实号只有落库后才有），比不上只会判"不一致"→ 走重建，方向安全。
            expected = self._remap_data_b_items(entries, item_no_map)
            zul, kos, truncated = self._read_data_b_snapshot(base)
            # 费率成本中心列是否真读到，直接写进消息：这是"该列可否读"的实机结论，
            # 读到即比对四组（费率被人工改过也能发现），读不到则降级三组。
            rate_note = "含费率中心" if self._rate_centers_readable(kos) else "费率中心未读到"
            if truncated:
                diffs.append(
                    f"Data B 读满 {len(zul)}/{len(kos)} 行（判停失效，疑似空行回读非空），按有差异处理"
                )
            else:
                diff_desc = self._data_b_diff(zul, kos, expected, order)
                if not diff_desc:
                    return SapResult(
                        step="clear_data_b",
                        changed=False,
                        message=f"Data B 与 Excel 一致，已跳过（{rate_note}）",
                    )
                diffs.append(f"{diff_desc}（{rate_note}）")
            # 判定为有差异时一并输出两表原文快照与 Excel 期望值：归一化后的比对结果看不出
            # 根因（回读格式、编号体系、多余空行都会导致"看起来一样却判不一致"）。
            diffs.append(self._data_b_snapshot_note(zul, kos, expected))

            # 分两阶段删空，规避两个子表行数不一致：强制成本中心行只在 ZULEISTENDE
            # (执行部门) 有行、KOSTENSAETZE (费率/POSNR/固定价格) 无行。
            # 先按费率(KOSTENSAETZE)行数删配对行（两表同选同删），再删剩余执行部门
            # (ZULEISTENDE) 独占的强制行（只选执行部门表）。若沿用"同行号双表同删"，
            # 删到强制行时会去选 KOSTENSAETZE 上不存在/无值的行 → 费率成本中心报错。
            kos_rows = self._count_data_b_kos_rows(base)
            for row in range(kos_rows - 1, -1, -1):
                self._delete_data_b_row(base, row)
            remaining = self._count_existing_data_b_rows(base)
            for row in range(remaining - 1, -1, -1):
                self._delete_data_b_zul_row(base, row)
        except Exception as exc:
            return SapResult.fail(f"Data B 清空失败，{exc}", step="clear_data_b")

        diffs.append(f"Data B 有差异：已删除旧 {len(zul)} 行待重建")
        return SapResult(step="clear_data_b", changed=True, message="；".join(diffs))

    def write_data_b(
        self,
        entries: list[DataBEntry],
        order: OrderData,
        diffs: list[str],
        item_no_map: dict[str, str] | None = None,
    ) -> SapResult:
        """Data B 同步第二段：从空表重建全部行；写入口径与创建 fill_lab_cost_entries 同源。

        前置条件（调用方保证）：clear_data_b() 已删空、且其后已 save + open_order——
        执行部门行落库成为贡献成本中心后，写成本表才不会报 ZR520（见 clear_data_b 文档）。
        open_order 会重置页面，故此处重新导航进抬头 T\\14 页签。

        Args:
            item_no_map: 中途 save+open 后建立的 ODM→SAP item 号映射；新增被 SAP 改号的
                item 需用真实号写 POSNR/ZPOSITION，否则指向错误 item
                （见 .claude/plan/va02_edit_item_add_midsave.md）。未传或未命中时用原号。

        Note:
            订单价值(AUFTRAGSWERT) 已从本方法剥离，改由 edit_order_value() 独立步骤对比更新。
        """
        try:
            self.session.press("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD")
            self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14")
            rebuilt = self._remap_data_b_items(entries, item_no_map)
            self._base._write_lab_cost_rows(rebuilt, order)
        except Exception as exc:
            return SapResult.fail(f"Data B 重建失败，{exc}", step="write_data_b")

        normal_count = sum(1 for e in entries if not e.kostl_only)
        forced_count = sum(1 for e in entries if e.kostl_only)
        diffs.append(f"Data B 重建：正常 {normal_count} 行 + 强制 {forced_count} 行")
        return SapResult(step="write_data_b", message="；".join(diffs))

    def _read_data_b_snapshot(
        self,
        base: str,
        max_rows: int = 50,
    ) -> tuple[list[tuple[str, str]], list[tuple[str, str, str]], bool]:
        """读 Data B 当前内容快照，供与 Excel 比对。

        Returns:
            (zul, kos, truncated)：
              - zul: [(执行部门成本中心, item)]，含强制行；以执行部门列为空判停。
              - kos: [(费率成本中心, POSNR, 固定价格原文)]，不含强制行；以固定价格列为空判停。
              - truncated: 任一表读满 max_rows，行数可能被可见行数截断，结果不可信。

        费率成本中心 `ctxtTABD-KOSTL` 只读、只读**配对行**：以固定价格判停保证 row 落在成本表
        真实存在的行上（强制成本中心行在成本表没有行，永远不会被读到）。原"该列不可读"的说法
        缺实测依据——历史证据只支持"**聚焦/写入**强制行的该格会中断 SAP"，而正常行的写入
        (`order.py::_write_lab_cost_rows`) 一直在 set_text 这个控件。
        单独兜住该列的读取异常：读不到就退回空串，比对侧据此跳过费率组（降级为原三组口径），
        绝不因读不到而中断快照或恒判不一致。
        """
        zul: list[tuple[str, str]] = []
        kos: list[tuple[str, str, str]] = []
        for row in range(max_rows):
            try:
                kostl = self._norm(self.session.read_text(
                    f"{base}/tblSAPMV45AZULEISTENDE/ctxtTABL-KOSTL[0,{row}]"
                ))
                if not kostl:
                    break
                item = self._norm(self.session.read_text(
                    f"{base}/tblSAPMV45AZULEISTENDE/txtTABL-ZPOSITION[1,{row}]"
                ))
            except SapUiError:
                break
            zul.append((kostl, item))
        for row in range(max_rows):
            try:
                festpreis = self._norm(self.session.read_text(
                    f"{base}/tblSAPMV45AKOSTENSAETZE/txtTABD-FESTPREIS[5,{row}]"
                ))
                if not festpreis:
                    break
                posnr = self._norm(self.session.read_text(
                    f"{base}/tblSAPMV45AKOSTENSAETZE/txtTABD-POSNR[1,{row}]"
                ))
            except SapUiError:
                break
            try:
                rate = self._norm(self.session.read_text(
                    f"{base}/tblSAPMV45AKOSTENSAETZE/ctxtTABD-KOSTL[0,{row}]"
                ))
            except SapUiError:
                rate = ""
            kos.append((rate, posnr, festpreis))
        return zul, kos, len(zul) >= max_rows or len(kos) >= max_rows

    @classmethod
    def _data_b_diff(
        cls,
        zul: list[tuple[str, str]],
        kos: list[tuple[str, str, str]],
        entries: list[DataBEntry],
        order: OrderData,
    ) -> str:
        """比对 SAP 快照与 Excel entries，返回差异描述；**空串表示一致**（无需删空重建）。

        两张表口径不同（用户 2026-08-20 定稿），都不比行序（SAP 按成本中心号重排）：

        ① **执行部门表按集合"包含"判定**——只看成本中心，要求 Excel 正常行 + config 强制行的
           全部成本中心都出现在 SAP 侧即可。不用多重集：SAP 的贡献成本中心清单同一成本中心
           通常只保留一行，而 Excel 可能有多行同中心（不同 item/金额），数量天然对不上。
           item 列(ZPOSITION) 不参与——item 的正确性由成本表 POSNR 覆盖。
           **代价**：SAP 多出的执行部门成本中心不触发重建（与"多一个贡献中心不影响业务"的
           强制成本中心行语义一致）。

        ② **成本表按行三字段严格比对**——(费率成本中心, POSNR, 固定价格) 三元组的多重集必须
           完全相等。用三元组而非三列各自比：三列独立比会把"两行之间金额互换"误判为一致。
           降级规则：费率列读不到 → 两侧该位置置空（退化为 POSNR+金额）；
           sales_group == '240' 不写 item 号 → 两侧 POSNR 置空。

        差异描述带上两侧实际值：SAP 回读格式与 Excel 不同（如金额小数分隔符、item 前导零）
        会让比对恒不相等、短路永不生效，日志里能直接看出是格式问题还是真实数据差异。
        """
        normal = [e for e in entries if not e.kostl_only]

        # ① 执行部门表：Excel(含强制行) 的成本中心必须都在 SAP 侧。
        #    编号一律走 _norm_no 去前导零——SAP 的 KOSTL 是 CHAR10 定长，回读带前导零。
        sap_centers = {cls._norm_no(z[0]) for z in zul if cls._norm_no(z[0])}
        excel_centers = {
            cls._norm_no(e.performer_cost_center)
            for e in entries
            if cls._norm_no(e.performer_cost_center)
        }
        missing = excel_centers - sap_centers
        if missing:
            return (
                f"执行部门缺少={cls._fmt_bag(sorted(missing))} "
                f"SAP={cls._fmt_bag(sorted(sap_centers))} Excel={cls._fmt_bag(sorted(excel_centers))}"
            )

        # ② 成本表：三字段成行比对。读不到的费率列 / 240 订单的 POSNR 两侧同时置空，等价于不比该字段。
        use_rate = cls._rate_centers_readable(kos)
        use_item = order.sales_group != "240"
        sap_bag = Counter(
            (
                cls._norm_no(k[0]) if use_rate else "",
                cls._norm_no(k[1]) if use_item else "",
                cls._norm_amount(k[2]),
            )
            for k in kos
        )
        excel_bag = Counter(
            (
                (cls._norm_no(e.rate_cost_center) or cls._norm_no(e.performer_cost_center)) if use_rate else "",
                cls._norm_no(cls._first_item_no(e.item)) if use_item else "",
                cls._norm_amount(e.amount),
            )
            for e in normal
        )
        if sap_bag != excel_bag:
            return f"成本表不同 SAP={cls._fmt_bag(sap_bag)} Excel={cls._fmt_bag(excel_bag)}"
        return ""

    @classmethod
    def _data_b_snapshot_note(
        cls,
        zul: list[tuple[str, str]],
        kos: list[tuple[str, str, str]],
        entries: list[DataBEntry],
    ) -> str:
        """渲染两表 SAP 原文快照 + Excel 期望值，供实机定位"看起来一样却判不一致"。

        输出**未归一化的原文**（金额不套 .2f、item 不裁剪），这样回读格式差异（千分位、
        小数分隔符、前导零）能直接暴露；Excel 侧标出强制行(F)，便于核对执行部门表多出的行。
        """
        zul_txt = ",".join(f"{k}|{i}" for k, i in zul) or "空"
        kos_txt = ",".join(f"{r}|{p}|{f}" for r, p, f in kos) or "空"
        excel_txt = ",".join(
            f"{cls._norm(e.performer_cost_center)}|{cls._norm(e.rate_cost_center)}"
            f"|{cls._norm(e.item)}|{e.amount}{'|F' if e.kostl_only else ''}"
            for e in entries
        ) or "空"
        return (
            f"Data B 快照 执行部门[中心|item]={zul_txt} "
            f"成本表[费率|POSNR|金额]={kos_txt} "
            f"Excel[执行部门|费率|item|金额]={excel_txt}"
        )

    @classmethod
    def _rate_centers_readable(cls, kos: list[tuple[str, str, str]]) -> bool:
        """成本表费率成本中心列是否全部读到（有任一为空即视为读不到，跳过该组比对）。

        无成本表行时返回 True——没有行可比，等价于"该组无差异"。
        """
        return all(cls._norm(k[0]) for k in kos)

    @staticmethod
    def _fmt_bag(bag) -> str:
        """多重集/集合/列表渲染为稳定可读文本（排序去随机性），供差异日志逐项比对。

        元素为元组（成本表的 费率/POSNR/金额 三字段行）时用 "/" 连接；Counter 按次数展开重复项。
        """
        raw = bag.elements() if isinstance(bag, Counter) else bag
        return "[" + ",".join(sorted(
            "/".join(x) if isinstance(x, tuple) else str(x) for x in raw
        )) + "]"

    @staticmethod
    def _norm_no(value) -> str:
        """SAP 定长编号归一化：去前导零后再比对。

        SAP 的成本中心 `KOSTL`(CHAR10) 与 item 号 `POSNR`(6 位) 回读恒带前导零
        （实测：`0048601258` / `001000`），Excel 侧不带 ⇒ 直接比对必然不等、一致性短路
        永不生效（症状：数据一样也走删除+保存+重建）。
        只用于**比对**，写入一律仍用原值——SAP 会自行补零。
        全零串归一为 "0" 而非空串，避免被后续"非空"判定当成无值。
        """
        raw = OrderEditTransaction._norm(value)
        stripped = raw.lstrip("0")
        return stripped or ("0" if raw else "")

    @staticmethod
    def _first_item_no(item) -> str:
        """取首个 ";" 前的 item 号，与写入侧 _write_lab_cost_rows 的单值裁剪口径一致。"""
        raw = OrderEditTransaction._norm(item)
        return raw.split(";", 1)[0].strip() if raw else ""

    @staticmethod
    def _remap_data_b_items(
        entries: list[DataBEntry],
        item_no_map: dict[str, str] | None,
    ) -> list[DataBEntry]:
        """把 entries 中新增改号 item 的 ODM 号替换为 SAP 真实号，供 POSNR/ZPOSITION 定位。

        item_no_map 为空/未命中时原样返回（绝不把 item 写成空）；强制行 item 恒为空，不受影响。
        取首个 ";" 前的 item 号做映射，与写入侧 _write_lab_cost_rows 的单值裁剪口径一致。
        """
        if not item_no_map:
            return entries
        rebuilt: list[DataBEntry] = []
        for entry in entries:
            raw = (entry.item or "").strip()
            first = raw.split(";", 1)[0].strip() if raw else ""
            mapped = item_no_map.get(first) if first else None
            rebuilt.append(replace(entry, item=mapped) if mapped else entry)
        return rebuilt

    def _delete_data_b_row(self, base: str, row: int) -> None:
        """删除 Data B 指定行（见用户 SAP 录屏）：选中 ZULEISTENDE/KOSTENSAETZE 两个子表的该行
        → 聚焦执行部门格 → 按专用删除按钮 btnTABLOESCH → 确认弹窗。

        Data B 删除非"置空控件"，须两个子表同时选中该行再按删除按钮。
        聚焦落在 ZULEISTENDE/ctxtTABL-KOSTL（执行部门）而非 KOSTENSAETZE/ctxtTABD-KOSTL：
        强制成本中心行的 KOSTENSAETZE/TABD-KOSTL 为空/不可编辑，聚焦该元素会导致 SAP 流程
        中断；删除按钮依"选中行"生效，聚焦仅定位光标，改聚焦执行部门列同样有效且安全。
        """
        zul = f"{base}/tblSAPMV45AZULEISTENDE"
        kos = f"{base}/tblSAPMV45AKOSTENSAETZE"
        self.session.find(zul).getAbsoluteRow(row).selected = True
        self.session.find(kos).getAbsoluteRow(row).selected = True
        self.session.focus(f"{zul}/ctxtTABL-KOSTL[0,{row}]", 0)
        self.session.press(f"{base}/btnTABLOESCH")
        self._try_press("wnd[1]/usr/btnSPOP-OPTION1")

    def _delete_data_b_zul_row(self, base: str, row: int) -> None:
        """删除只在 ZULEISTENDE(执行部门) 存在、KOSTENSAETZE 无对应行的强制成本中心行。

        强制行只写了执行部门，费率子表 KOSTENSAETZE 没有该行，故只选中 ZULEISTENDE 的该行、
        聚焦执行部门格再按删除按钮 btnTABLOESCH，绝不触碰 KOSTENSAETZE（否则费率列报错）。
        """
        zul = f"{base}/tblSAPMV45AZULEISTENDE"
        self.session.find(zul).getAbsoluteRow(row).selected = True
        self.session.focus(f"{zul}/ctxtTABL-KOSTL[0,{row}]", 0)
        self.session.press(f"{base}/btnTABLOESCH")
        self._try_press("wnd[1]/usr/btnSPOP-OPTION1")

    def _count_existing_data_b_rows(self, base: str, max_rows: int = 50) -> int:
        """扫 Data B 执行部门列(ZULEISTENDE/ctxtTABL-KOSTL)到空行，返回现有行数（含强制行）。"""
        count = 0
        for row in range(max_rows):
            try:
                kostl = (self.session.read_text(
                    f"{base}/tblSAPMV45AZULEISTENDE/ctxtTABL-KOSTL[0,{row}]"
                ) or "").strip()
            except SapUiError:
                break
            if not kostl:
                break
            count += 1
        return count

    def _count_data_b_kos_rows(self, base: str, max_rows: int = 50) -> int:
        """扫 Data B 固定价格列(KOSTENSAETZE/txtTABD-FESTPREIS)到空行，返回费率子表行数（不含强制行）。

        强制成本中心行在 KOSTENSAETZE 无行，故此计数 = 配对(正常)行数，用于两阶段删除的第一阶段。
        用固定价格列而非费率列(ctxtTABD-KOSTL) 计数：后者在**强制行上不可聚焦/写入**（历史实测），
        计数只需一个恒有值的列即可——正常行的固定价格总以 .2f 落 "0.00"，等价且不涉及任何写操作。
        （费率列的只读读取见 _read_data_b_snapshot：对配对行可读，"该列不可读"的旧说法无实测依据。）
        """
        count = 0
        for row in range(max_rows):
            try:
                festpreis = (self.session.read_text(
                    f"{base}/tblSAPMV45AKOSTENSAETZE/txtTABD-FESTPREIS[5,{row}]"
                ) or "").strip()
            except SapUiError:
                break
            if not festpreis:
                break
            count += 1
        return count

    def edit_order_value(self, order: OrderData, diffs: list[str]) -> SapResult:
        """对比并更新订单价值(AUFTRAGSWERT)：Σ SAP item 未税净值 × 汇率。

        口径与创建 fill_order_value 完全一致（Σ VBAP-NETWR × exchange_rate）。
        达阈值写换算值；低于阈值时目标为空——既符合"小额不回填"规则，又能把历史因
        双重汇率误写的脏值(如 35675)清空，供重跑编辑自愈。
        """
        result = SapResult(step="edit_order_value")
        try:
            # 读净值须在 item 概览页；复用创建事务的加和逻辑，保证与创建同口径。
            self._base._ensure_item_overview()
            net_total, truncated = self._base._sum_item_net_values()
            order_value_cny = net_total * (order.exchange_rate or 1.0)

            self.session.press("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD")
            self.session.select_tab("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14")
            target = (
                format(order_value_cny, ".2f")
                if order_value_cny >= self.config.revenue_threshold
                else ""
            )
            # prefix_field=False：本段只有这一个字段、段名已由 mixin 的 _append_remark 提供，
            # 再带 `订单价值:` 会输出 `订单价值:订单价值:3,021.26→(空)` 双重前缀。
            self._compare_and_set(
                self._base._auftragswert_id(),
                target,
                field="订单价值",
                diffs=diffs,
                amount=True,
                prefix_field=False,
            )
            if truncated:
                result.warning = True
                diffs.append("item 行数超过扫描上限，可能少算，请人工核对")
        except Exception as exc:
            return SapResult.fail(f"订单价值编辑失败，{exc}", step="edit_order_value")
        result.message = "；".join(diffs) if diffs else "订单价值无差异"
        return result

    _PLAN_COST_TABLE = "wnd[0]/usr/tblSAPLKKDI1301_TC"

    @classmethod
    def _pc_center_id(cls, row: int) -> str:
        """计划成本编辑器成本中心列(HERK2)控件 ID。"""
        return f"{cls._PLAN_COST_TABLE}/ctxtRK70L-HERK2[3,{row}]"

    @classmethod
    def _pc_category_id(cls, row: int) -> str:
        """计划成本编辑器类别列(HERK3，FREMDL/T01AST)控件 ID。"""
        return f"{cls._PLAN_COST_TABLE}/ctxtRK70L-HERK3[4,{row}]"

    @classmethod
    def _pc_amount_id(cls, row: int) -> str:
        """计划成本编辑器数量/金额列(MENGE)控件 ID。"""
        return f"{cls._PLAN_COST_TABLE}/txtRK70L-MENGE[6,{row}]"

    def edit_plan_cost(
        self,
        entries: list[PlanCostEntry],
        diffs: list[str],
        *,
        target_item: str,
    ) -> SapResult:
        """按 (成本中心,类别) 主键匹配指定 SAP item 的计划成本：仅差异才写，顺序不同不算变化。

        先按 item 号在 SAP item 概览页实时定位物理行（ODM 与 SAP 编号可能不同、SAP 也可能
        重排，故每次重读匹配，不用写入时的行号）。SAP 不存在该 item → 不开编辑器、成功跳过。
        跨 item 隔离：编辑器 tblSAPLKKDI1301_TC 只含该 item 的行，entries 也来自单 item，两侧同源。

        规则（用户 2026-08-21 定稿，顺序无关）：
            - 主键 = (_norm_no(成本中心), _norm(类别))，对行序免疫；
            - Excel entry 认领同键 SAP 行 → 仅金额有差异才改 MENGE 列，日志 `金额|时间 旧→新`；
            - Excel 有 SAP 无 → 追加到编辑器末尾（全写四列），日志 `… (空)→新`；
            - SAP 有 Excel 无 → Shift+F2 删除（倒序，避免行号位移），日志 `…：Excel 无，已删除`；
            - 全认领且金额全等（含顺序不同）→ 零写入零回车，`Plan Cost 无差异`。

        执行顺序：①改金额（行号稳定先做）→ ②倒序删多余 → ③追加新增（删后重读行数取起始 row）。
        """
        result = SapResult(step="edit_plan_cost")
        try:
            self._base._ensure_item_overview()
            focus_row = self._find_item_physical_row(target_item)
            if focus_row is None:
                # SAP 无此 item：不失败、不开编辑器，但标 warning 让 UI 用区别色提示。
                result.warning = True
                result.message = f"SAP 无对应 item {self._norm(target_item)}，已跳过"
                return result
            self._open_plan_cost_editor_for_edit(OrderTransaction._material_id(focus_row))

            valid = [e for e in entries if e.cost_center]
            existing = self._read_existing_plan_cost_rows()  # [(row, cost_center, category, amount)]

            # SAP 行按 (归一中心,类别) 主键分组，供 Excel entry 认领（对顺序免疫；同键多行按 row 升序消费）。
            sap_by_key: dict[tuple[str, str], list[tuple[int, str, str, str]]] = {}
            for row_tuple in existing:
                key = (self._norm_no(row_tuple[1]), self._norm(row_tuple[2]))
                sap_by_key.setdefault(key, []).append(row_tuple)

            # 阶段①：认领 + 仅金额差异才改（行号此时稳定，最先做）。认领不到的 entry 记为待新增。
            matched_rows: set[int] = set()
            to_add: list[PlanCostEntry] = []
            for entry in valid:
                key = (self._norm_no(entry.cost_center), self._norm(entry.category))
                bucket = sap_by_key.get(key)
                if bucket:
                    row, _center, _category, old_amount = bucket.pop(0)
                    matched_rows.add(row)
                    summary = self._overwrite_plan_cost_amount(row, entry, old_amount)
                    if summary:
                        diffs.append(summary)
                else:
                    to_add.append(entry)

            # 阶段②：SAP 有 Excel 无 → 倒序删除，避免删除后行号位移。
            leftover = sorted(
                (t for t in existing if t[0] not in matched_rows),
                key=lambda t: t[0],
                reverse=True,
            )
            for row, cost_center, category, amount in leftover:
                self._delete_plan_cost_row(row)
                label = self._plan_cost_label(category)
                diffs.append(
                    f"{self._plan_cost_head(cost_center, category)} {label} {amount}：Excel 无，已删除"
                )

            # 阶段③：Excel 有 SAP 无 → 删后重读行数，从当前末尾逐行追加（全写四列，口径同创建）。
            # 仅在确有新增时才重读，避免"无新增"常见场景多一次 SAP 往返。
            if to_add:
                next_row = len(self._read_existing_plan_cost_rows())
                for entry in to_add:
                    self._base._apply_single_plan_cost_entry(next_row, entry)
                    diffs.append(self._plan_cost_add_summary(entry))
                    next_row += 1

            self.session.press("wnd[0]/tbar[0]/btn[3]")
            # 退出编辑器时若弹"是否保存"确认框，按 OPTION1 兜底确认（无则跳过）。
            self._try_press("wnd[1]/usr/btnSPOP-OPTION1")
        except Exception as exc:
            return SapResult.fail(f"plan cost 编辑失败，{exc}", step="edit_plan_cost")
        result.message = "；".join(diffs) if diffs else "Plan Cost 无差异"
        return result

    def _delete_plan_cost_row(self, row: int) -> None:
        """删除计划成本编辑器指定行：聚焦该行成本中心格后 Shift+F2（sendVKey 14）。

        SAP 删除非"置空控件"，须定位行后触发 Shift+F2；聚焦 `HERK2[3,row]`（每行必有的
        成本中心列，即本类读写用的列）即定位该行。删除后若弹确认框，OPTION1 兜底。
        """
        self.session.focus(self._pc_center_id(row), 8)
        self.session.send_vkey(14)
        self._try_press("wnd[1]/usr/btnSPOP-OPTION1")

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
                cost_center = (self.session.read_text(self._pc_center_id(row)) or "").strip()
                category = (self.session.read_text(self._pc_category_id(row)) or "").strip()
            except SapUiError:
                break
            if not cost_center and not category:
                break
            try:
                amount = (self.session.read_text(self._pc_amount_id(row)) or "").strip()
            except SapUiError:
                amount = ""
            rows.append((row, cost_center, category, amount))
        return rows

    def _plan_cost_label(self, category) -> str:
        """计划成本数量列标签：T01AST(工时)→"时间"，其余→"金额"。"""
        return "时间" if self._norm(category) == "T01AST" else "金额"

    def _plan_cost_head(self, cost_center, category) -> str:
        """计划成本行抬头：`计划成本 成本中心{中心}({类别})`。"""
        return f"计划成本 成本中心{self._norm(cost_center)}({self._norm(category)})"

    def _overwrite_plan_cost_amount(
        self, row: int, entry: PlanCostEntry, old_amount: str
    ) -> str:
        """认领到同键 SAP 行后：仅金额有差异才改 MENGE 列 + 回车，返回箭头摘要；相等返回空串。

        金额按 `_norm_amount`(.2f、去千分位) 归一比对，避免回读格式差异误判；
        成本中心/类别已由主键保证一致，无需重写（也不触碰只读的强制行式布局）。
        """
        new_disp = format(float(entry.amount), ".2f")
        if self._norm_amount(old_amount) == self._norm_amount(new_disp):
            return ""
        self.session.set_text(self._pc_amount_id(row), new_disp)
        self.session.focus(self._pc_amount_id(row), 20)
        self.session.send_vkey(0)
        label = self._plan_cost_label(entry.category)
        head = self._plan_cost_head(entry.cost_center, entry.category)
        old_disp = self._norm(old_amount) or "(空)"
        return f"{head} {label} {old_disp}→{new_disp}"

    def _plan_cost_add_summary(self, entry: PlanCostEntry) -> str:
        """新增行摘要：`计划成本 成本中心X(类别) 金额|时间 (空)→值`。"""
        label = self._plan_cost_label(entry.category)
        head = self._plan_cost_head(entry.cost_center, entry.category)
        return f"{head} {label} (空)→{format(float(entry.amount), '.2f')}"

    # ------------------------------------------------------------------ #
    # 复用创建侧 open/save
    # ------------------------------------------------------------------ #
    def open(self, order_no: str) -> SapResult:
        """复用创建侧 VA02 打开逻辑。"""
        return self._base.open(order_no)

    def save(self, info: str) -> SapResult:
        """复用创建侧保存逻辑。"""
        return self._base.save(info)
