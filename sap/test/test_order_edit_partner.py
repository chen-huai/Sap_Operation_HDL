"""VA02 编辑：伙伴三角色（Buyer(GPC) / Primary CS / Sales）幂等同步的回归测试。

覆盖用户 2026-08-25 提的两种情况：
    ① 付款方变动（Excel SAP No. 改动 → SAP 重跑 partner determination）→ 三行被清空 → 补回；
    ② 付款方未变、三者之一自身变了 → 改。

关键约束：行位必须沿用**创建同源口径**（`order.py:_fill_partners`）——SAP 带出的行 4/5
一为负责雇员、另一为送达方，创建把送达方那行改成 ZG 写 Buyer 值，VE 固定行 7。
第一轮修法"找不到角色行就新增到第一个空行"实机无效，故本测试重点断言**落在哪一行**。
"""

from __future__ import annotations

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

import sap.test.helpers  # noqa: E402,F401  仅为注入 win32com mock
from sap.models import OrderData, SapConfig  # noqa: E402
from sap.session import SapSession  # noqa: E402
from sap.transactions.order_edit import OrderEditTransaction  # noqa: E402


def _make_config(cs_code: str = "", sales_code: str = "SA001") -> SapConfig:
    # cs_code 默认置空，便于把单个角色路径隔离出来测。
    return SapConfig(
        order_type="ZOR",
        sales_organization="3002",
        distribution_channels="10",
        sales_office="1000",
        sales_group="200",
        sub_cost_center_cs="1101",
        sub_cost_center_chm="1102",
        sub_cost_center_phy="1103",
        cs_code=cs_code,
        sales_code=sales_code,
    )

PARTNER_PREFIX = (
    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/"
    "ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:"
    "SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW"
)

PARVW_MARK = "cmbGVS_TC_DATA-REC-PARVW"

# SAP 侧的**系统级 PARVW 选项表**：key → combo 显示文本。同一屏所有 PARVW 下拉框共享，
# 既驱动桩控件的 key→text 联动，也作为 `Entries` 枚举的数据源。
ROLE_TEXTS = {"AG": "售达方", "ZG": "GPC", "ER": "负责雇员", "VE": "销售员", "WE": "送达方"}

# 负责雇员的真实 key 为未知值 ZP 的环境（`ER` 已被实机否决）。
ROLE_TEXTS_ZP = {"AG": "售达方", "ZG": "GPC", "ZP": "负责雇员", "VE": "销售员", "WE": "送达方"}

# 中文翻译歧义环境：Ship-to party(WE) 与 Global Partner(ZG) 的显示文本**都是"送达方"**。
# 这正是 `GPC Code:角色key ZG 无效(待校正)` 背后不能按文本认角色的原因。
ROLE_TEXTS_AMBIGUOUS = {"AG": "售达方", "WE": "送达方", "ZG": "送达方", "ER": "负责雇员"}


def _parvw_id(row: int) -> str:
    return f"{PARTNER_PREFIX}/{PARVW_MARK}[0,{row}]"


def _partner_id(row: int) -> str:
    return f"{PARTNER_PREFIX}/ctxtGVS_TC_DATA-REC-PARTNER[1,{row}]"


class _Element:
    """可控 SAP 控件：text/key 可读写，焦点/按键为 no-op。"""

    def __init__(self, text: str = "", key: str = ""):
        self.text = text
        self.key = key
        self.caretPosition = 0

    def setFocus(self):
        pass

    def sendVKey(self, _key):
        pass

    def press(self):
        pass

    def select(self):
        pass


class _ComboEntry:
    """GuiComboBoxEntry 桩：一个下拉选项的 (Key, Value)。"""

    def __init__(self, key: str, value: str):
        self.Key = key
        self.Value = value


class _ComboEntries:
    """GuiCollection 桩：按 Count / ElementAt(i) 访问，与 SAP GUI Scripting 一致。"""

    def __init__(self, role_texts: dict[str, str]):
        self._items = [_ComboEntry(k, v) for k, v in role_texts.items()]

    @property
    def Count(self) -> int:
        return len(self._items)

    def ElementAt(self, index: int) -> _ComboEntry:
        return self._items[index]


class _ComboElement(_Element):
    """PARVW 下拉框桩：写 key 时按 role_texts 联动显示文本，模拟真实 combo。

    `role_texts` 同时是该 combo 的选项表，经 `Entries` 暴露给
    `SapSession.list_combo_entries`——生产代码正是靠它把显示文本反查成 key。

    `invalid_keys` 里的 key 赋值抛错，模拟 SAP 的拒绝（选项表里没有 / 角色列只读）；
    `silent_keys` 里的 key 赋值被静默吞掉（不抛也不生效），模拟"写进去了但 SAP 改回来"，
    用于验证写后回读自证。
    """

    def __init__(
        self,
        role_texts: dict[str, str],
        invalid_keys: frozenset[str],
        key: str = "",
        *,
        silent_keys: frozenset[str] = frozenset(),
        entries_available: bool = True,
    ):
        # 须先于 super()，父类 __init__ 里的 `self.key = ...` 会走下面的 setter。
        self._role_texts = role_texts
        self._invalid_keys = invalid_keys
        self._silent_keys = silent_keys
        self._entries_available = entries_available
        super().__init__(key=key)

    def rebind(
        self,
        role_texts: dict[str, str],
        invalid_keys: frozenset[str],
        silent_keys: frozenset[str],
        entries_available: bool,
    ) -> None:
        """对齐到会话级选项表——真实 SAP 里同屏各行共享同一张 PARVW 系统配置表。

        `_combo()` 造桩时会话尚不存在，拿不到本用例的选项表，故由 `_RawSession` 事后统一
        绑定；否则 preset 行与 SAP 自动补出的行会看到两张不同的选项表，与实机不符。
        """
        self._role_texts = role_texts
        self._invalid_keys = invalid_keys
        self._silent_keys = silent_keys
        self._entries_available = entries_available

    @property
    def Entries(self) -> _ComboEntries:
        # 老版 SAP GUI 无此属性：用 AttributeError 模拟，验证生产代码的兜底路径。
        if not self._entries_available:
            raise AttributeError("Entries")
        return _ComboEntries(self._role_texts)

    @property
    def key(self) -> str:
        return self._key

    @key.setter
    def key(self, value: str) -> None:
        if value in self._invalid_keys:
            raise ValueError(f"invalid PARVW key: {value}")
        if value in self._silent_keys:
            return
        self._key = value
        self.text = self._role_texts.get(value, "") if value else ""


class _RawSession:
    """按 element_id 缓存控件的最小 raw session；未知 id 返回空控件。

    PARVW 下拉框走 `_ComboElement`（key→text 联动 + Entries 选项表），其余走普通 `_Element`。
    preset 中的 combo 由 `_make_transaction` 统一对齐到同一张选项表——真实 SAP 里选项表是
    系统级配置，同屏各行不可能不一致。
    """

    def __init__(
        self,
        preset: dict[str, _Element] | None = None,
        *,
        role_texts: dict[str, str] | None = None,
        invalid_keys: frozenset[str] = frozenset(),
        silent_keys: frozenset[str] = frozenset(),
        entries_available: bool = True,
    ):
        self._cache: dict[str, _Element] = preset or {}
        self._role_texts = ROLE_TEXTS if role_texts is None else role_texts
        self._invalid_keys = invalid_keys
        self._silent_keys = silent_keys
        self._entries_available = entries_available
        for element in self._cache.values():
            if isinstance(element, _ComboElement):
                element.rebind(
                    self._role_texts, invalid_keys, silent_keys, entries_available
                )

    def findById(self, element_id: str) -> _Element:
        if element_id not in self._cache:
            self._cache[element_id] = (
                _ComboElement(
                    self._role_texts, self._invalid_keys,
                    silent_keys=self._silent_keys, entries_available=self._entries_available,
                )
                if PARVW_MARK in element_id
                else _Element()
            )
        return self._cache[element_id]


def _combo(key: str, role_texts: dict[str, str] | None = None) -> _ComboElement:
    """构造一个已带角色 key 的 PARVW 桩控件（显示文本按映射联动）。

    role_texts 仅决定构造瞬间的显示文本；进入 `_RawSession` 后会被统一对齐到会话级选项表。
    """
    return _ComboElement(ROLE_TEXTS if role_texts is None else role_texts, frozenset(), key=key)


def _make_transaction(
    preset: dict[str, _Element] | None = None,
    *,
    cs_code: str = "",
    sales_code: str = "SA001",
    role_texts: dict[str, str] | None = None,
    invalid_keys: frozenset[str] = frozenset(),
    silent_keys: frozenset[str] = frozenset(),
    entries_available: bool = True,
):
    raw = _RawSession(
        preset, role_texts=role_texts, invalid_keys=invalid_keys,
        silent_keys=silent_keys, entries_available=entries_available,
    )
    session = SapSession(raw, raw, raw, raw)
    return OrderEditTransaction(session, _make_config(cs_code, sales_code)), raw


def _make_order(global_partner_code: str = "") -> OrderData:
    return OrderData(
        sap_no="123456",
        project_no="PRJ-001",
        currency_type="CNY",
        exchange_rate=1.0,
        short_text="Test",
        global_partner_code=global_partner_code,
    )


def _sap_determined_rows(gpc_key: str = "WE", employee_at_four: bool = True) -> dict[str, _Element]:
    """构造 SAP partner determination 带出的行 4/5：一行负责雇员、一行送达方(或指定角色)。

    Args:
        gpc_key: Buyer(GPC) 目标行当前的角色 key，默认 WE（SAP 重置后的原生送达方）。
        employee_at_four: 负责雇员在行 4（→ gpc 落行 5）还是行 5（→ gpc 落行 4）。
    """
    e_row, g_row = (4, 5) if employee_at_four else (5, 4)
    return {
        _parvw_id(e_row): _combo("ER"),
        _parvw_id(g_row): _combo(gpc_key),
    }


class ResolvePartnerRowsTest(unittest.TestCase):
    """行位判定必须与创建 order.py:157-159 完全同口径。"""

    def test_employee_at_row_four(self):
        tx, _ = _make_transaction(_sap_determined_rows(employee_at_four=True))
        self.assertEqual(tx._resolve_partner_rows(PARTNER_PREFIX), (4, 5))

    def test_employee_at_row_five(self):
        tx, _ = _make_transaction(_sap_determined_rows(employee_at_four=False))
        self.assertEqual(tx._resolve_partner_rows(PARTNER_PREFIX), (5, 4))

    def test_blank_table_falls_back_to_create_default(self):
        # 行 4 读不到角色文本 → 退回创建路径的默认分支 (5, 4)。
        tx, _ = _make_transaction()
        self.assertEqual(tx._resolve_partner_rows(PARTNER_PREFIX), (5, 4))


class ResolveSalesRowTest(unittest.TestCase):
    def test_existing_ve_row_wins(self):
        tx, _ = _make_transaction({_parvw_id(3): _combo("VE")})
        self.assertEqual(tx._resolve_sales_row(PARTNER_PREFIX), 3)

    def test_falls_back_to_create_row_seven(self):
        tx, _ = _make_transaction(_sap_determined_rows())
        self.assertEqual(tx._resolve_sales_row(PARTNER_PREFIX), OrderEditTransaction._SALES_ROW)

    def test_occupied_row_seven_falls_back_to_empty_row(self):
        # 行 7 已被别的角色占用 → 退回首个空行（此处行 8），绝不覆盖已有伙伴。
        # 行 0-3 填上 SAP 原生角色，贴近实机：行 0 是售达方，不可能为空。
        preset = _sap_determined_rows()
        for row, key in enumerate(["AG", "RE", "RG", "WE"]):
            preset[_parvw_id(row)] = _combo(key, {key: key})
            preset[_partner_id(row)] = _Element(text="C000")
        preset[_parvw_id(6)] = _combo("AP")
        preset[_partner_id(6)] = _Element(text="CT001")
        preset[_parvw_id(7)] = _combo("AP")
        preset[_partner_id(7)] = _Element(text="CT002")

        tx, _ = _make_transaction(preset)
        self.assertEqual(tx._resolve_sales_row(PARTNER_PREFIX), 8)


class SyncPartnerRowTest(unittest.TestCase):
    """四种 SAP 状态都要收敛到期望值。"""

    def test_row_present_value_empty_writes_value(self):
        # 状态①：行在、key 已对、值空（付款方变动被清空）→ 只写值，不动 key。
        tx, raw = _make_transaction({_parvw_id(5): _combo("ZG")})
        diffs: list[str] = []
        tx._sync_partner_row(PARTNER_PREFIX, 5, "ZG", "GP001", field="GPC Code", diffs=diffs)

        self.assertEqual(raw.findById(_partner_id(5)).text, "GP001")
        self.assertEqual(diffs, ["GPC Code:(空)→GP001"])

    def test_row_present_value_stale_writes_value(self):
        # 状态②：行在、值不对（字段自身变动）→ 只写值。
        preset = {_parvw_id(5): _combo("ZG"), _partner_id(5): _Element(text="OLD")}
        tx, raw = _make_transaction(preset)
        diffs: list[str] = []
        tx._sync_partner_row(PARTNER_PREFIX, 5, "ZG", "GP001", field="GPC Code", diffs=diffs)

        self.assertEqual(raw.findById(_partner_id(5)).text, "GP001")
        self.assertEqual(diffs, ["GPC Code:OLD→GP001"])

    def test_role_reset_to_we_is_corrected(self):
        # 状态③：SAP 把该行恢复成原生送达方 WE → 改回 ZG 再写值，两条 diff 都要有。
        preset = {_parvw_id(5): _combo("WE"), _partner_id(5): _Element(text="C999")}
        tx, raw = _make_transaction(preset)
        diffs: list[str] = []
        tx._sync_partner_row(PARTNER_PREFIX, 5, "ZG", "GP001", field="GPC Code", diffs=diffs)

        self.assertEqual(raw.findById(_parvw_id(5)).key, "ZG")
        self.assertEqual(raw.findById(_partner_id(5)).text, "GP001")
        self.assertEqual(diffs, ["GPC Code:角色 WE→ZG", "GPC Code:C999→GP001"])

    def test_blank_row_gets_key_and_value(self):
        # 状态④：整行空（key 也空）→ 补角色再写值。
        tx, raw = _make_transaction()
        diffs: list[str] = []
        tx._sync_partner_row(PARTNER_PREFIX, 5, "ZG", "GP001", field="GPC Code", diffs=diffs)

        self.assertEqual(raw.findById(_parvw_id(5)).key, "ZG")
        self.assertEqual(raw.findById(_partner_id(5)).text, "GP001")
        self.assertEqual(diffs, ["GPC Code:角色 (空)→ZG", "GPC Code:(空)→GP001"])

    def test_no_change_is_zero_write(self):
        # key 与值都已正确 → 零写入零 diff（大多数订单走这条路径）。
        preset = {_parvw_id(5): _combo("ZG"), _partner_id(5): _Element(text="GP001")}
        tx, _ = _make_transaction(preset)
        diffs: list[str] = []
        tx._sync_partner_row(PARTNER_PREFIX, 5, "ZG", "GP001", field="GPC Code", diffs=diffs)
        self.assertEqual(diffs, [])

    def test_empty_expected_value_is_noop(self):
        tx, raw = _make_transaction()
        diffs: list[str] = []
        tx._sync_partner_row(PARTNER_PREFIX, 5, "ZG", "", field="GPC Code", diffs=diffs)
        self.assertEqual(diffs, [])
        self.assertEqual(raw.findById(_parvw_id(5)).key, "")

    def test_none_row_records_pending(self):
        tx, _ = _make_transaction()
        diffs: list[str] = []
        tx._sync_partner_row(PARTNER_PREFIX, None, "VE", "SA001", field="Sales", diffs=diffs)
        self.assertEqual(diffs, ["Sales:无可用行位(待校正)"])

    def test_key_present_but_readonly_records_pending(self):
        # key 在选项表里却写不进 → 角色列只读（已保存订单的强制伙伴行常如此），不写编码。
        tx, raw = _make_transaction(invalid_keys=frozenset({"ER"}))
        diffs: list[str] = []
        tx._sync_partner_row(
            PARTNER_PREFIX, 4, "ER", "CS001", field="Primary CS", diffs=diffs,
        )

        self.assertEqual(raw.findById(_partner_id(4)).text, "")
        self.assertEqual(diffs, ["Primary CS:角色列不可改(当前(空))(待校正)"])

    def test_key_absent_from_entries_lists_options(self):
        # key 根本不在选项表 → 文案带出可选项，实机一眼看出该填哪个角色。
        tx, raw = _make_transaction(
            role_texts={"AG": "售达方", "WE": "送达方"}, invalid_keys=frozenset({"ZG"}),
        )
        diffs: list[str] = []
        tx._sync_partner_row(PARTNER_PREFIX, 5, "ZG", "GP001", field="GPC Code", diffs=diffs)

        self.assertEqual(raw.findById(_partner_id(5)).text, "")
        self.assertEqual(diffs, ["GPC Code:角色key ZG 不在可选项[AG/WE](待校正)"])

    def test_entries_unavailable_keeps_legacy_message(self):
        # 老版 SAP GUI 无 Entries → 退回原有的笼统文案，行为不回归。
        tx, raw = _make_transaction(invalid_keys=frozenset({"ZG"}), entries_available=False)
        diffs: list[str] = []
        tx._sync_partner_row(PARTNER_PREFIX, 5, "ZG", "GP001", field="GPC Code", diffs=diffs)

        self.assertEqual(raw.findById(_partner_id(5)).text, "")
        self.assertEqual(diffs, ["GPC Code:角色key ZG 无效(待校正)"])

    def test_ship_to_display_text_is_never_accepted_as_gpc(self):
        """核心安全回归：显示文本"送达方"绝不能被当成 ZG 而写入 Buyer 编码。

        旧实现在 set_key 被拒后回读显示文本，命中 `_GPC_TEXTS`（错误地含"送达方"）就
        继续写值，于是 Buyer(GPC) 编码被静默挂到 Ship-to party(WE) 行上。
        """
        preset = {_parvw_id(5): _combo("WE", ROLE_TEXTS_AMBIGUOUS)}
        tx, raw = _make_transaction(
            preset, role_texts=ROLE_TEXTS_AMBIGUOUS, invalid_keys=frozenset({"ZG"}),
        )
        diffs: list[str] = []
        tx._sync_partner_row(PARTNER_PREFIX, 5, "ZG", "GP001", field="GPC Code", diffs=diffs)

        self.assertEqual(raw.findById(_parvw_id(5)).key, "WE")   # 角色未被改动
        self.assertEqual(raw.findById(_partner_id(5)).text, "")  # 编码绝不落到送达方行
        self.assertEqual(diffs, ["GPC Code:角色列不可改(当前WE)(待校正)"])

    def test_key_readback_mismatch_skips_write(self):
        # set_key 不抛错但没生效（SAP 悄悄改回）→ 回读自证失败，不写编码。
        tx, raw = _make_transaction(silent_keys=frozenset({"ZG"}))
        diffs: list[str] = []
        tx._sync_partner_row(PARTNER_PREFIX, 5, "ZG", "GP001", field="GPC Code", diffs=diffs)

        self.assertEqual(raw.findById(_partner_id(5)).text, "")
        self.assertEqual(
            diffs,
            ["GPC Code:角色 (空)→ZG", "GPC Code:角色key回读不符(ZG→(空))(待校正)"],
        )


class EditPartnersTest(unittest.TestCase):
    def test_payer_change_refills_all_three(self):
        # 情况①：付款方变动 → SAP 带出行4=负责雇员、行5=送达方，三者值均为空 → 全部补回。
        tx, raw = _make_transaction(_sap_determined_rows(), cs_code="CS001")
        diffs: list[str] = []
        tx._edit_partners(_make_order("GP001"), diffs)

        # Buyer 落行 5（负责雇员在行 4），角色由 WE 纠正为 ZG。
        self.assertEqual(raw.findById(_parvw_id(5)).key, "ZG")
        self.assertEqual(raw.findById(_partner_id(5)).text, "GP001")
        # CS 落行 4，角色本就是 ER，只写值。
        self.assertEqual(raw.findById(_parvw_id(4)).key, "ER")
        self.assertEqual(raw.findById(_partner_id(4)).text, "CS001")
        # Sales 落创建口径行 7。
        self.assertEqual(raw.findById(_parvw_id(7)).key, "VE")
        self.assertEqual(raw.findById(_partner_id(7)).text, "SA001")

        self.assertEqual(
            diffs,
            [
                "GPC Code:角色 WE→ZG",
                "GPC Code:(空)→GP001",
                "Primary CS:(空)→CS001",
                "Sales:角色 (空)→VE",
                "Sales:(空)→SA001",
            ],
        )

    def test_employee_at_row_five_swaps_target_rows(self):
        # 负责雇员在行 5 时 Buyer 必须落行 4，与创建 e_row/g_row 判定一致。
        tx, raw = _make_transaction(
            _sap_determined_rows(employee_at_four=False), cs_code="CS001", sales_code="",
        )
        tx._edit_partners(_make_order("GP001"), [])

        self.assertEqual(raw.findById(_partner_id(4)).text, "GP001")
        self.assertEqual(raw.findById(_partner_id(5)).text, "CS001")

    def test_single_field_change_only_writes_that_field(self):
        # 情况②：付款方未变，仅 Sales 变了 → 只改 Sales，另两个零写入。
        preset = _sap_determined_rows(gpc_key="ZG")
        preset[_partner_id(5)] = _Element(text="GP001")
        preset[_partner_id(4)] = _Element(text="CS001")
        preset[_parvw_id(7)] = _combo("VE")
        preset[_partner_id(7)] = _Element(text="OLD_SA")

        tx, raw = _make_transaction(preset, cs_code="CS001")
        diffs: list[str] = []
        tx._edit_partners(_make_order("GP001"), diffs)

        self.assertEqual(raw.findById(_partner_id(7)).text, "SA001")
        self.assertEqual(diffs, ["Sales:OLD_SA→SA001"])

    def test_all_match_is_zero_write(self):
        preset = _sap_determined_rows(gpc_key="ZG")
        preset[_partner_id(5)] = _Element(text="GP001")
        preset[_partner_id(4)] = _Element(text="CS001")
        preset[_parvw_id(7)] = _combo("VE")
        preset[_partner_id(7)] = _Element(text="SA001")

        tx, _ = _make_transaction(preset, cs_code="CS001")
        diffs: list[str] = []
        tx._edit_partners(_make_order("GP001"), diffs)
        self.assertEqual(diffs, [])

    def test_sales_optional_skipped_when_config_empty(self):
        # Sales 选填：config 无值时整段跳过，不占用行 7。
        tx, raw = _make_transaction(_sap_determined_rows(), cs_code="CS001", sales_code="")
        tx._edit_partners(_make_order("GP001"), [])
        self.assertEqual(raw.findById(_parvw_id(7)).key, "")

    def test_no_gpc_value_leaves_row_untouched(self):
        # Excel 未给 Buyer(GPC) → 不改也不补，该行保持 SAP 原样（送达方 WE）。
        tx, raw = _make_transaction(_sap_determined_rows(), sales_code="")
        tx._edit_partners(_make_order(), [])
        self.assertEqual(raw.findById(_parvw_id(5)).key, "WE")


class PrimaryCsRoleKeyTest(unittest.TestCase):
    """CS 段：命中负责雇员行时只写值、绝不动 key。

    实机结论（2026-08-25 log: `Primary CS:角色key ER 无效(待校正)`）：负责雇员行的
    PARVW key 不是 `ER`，该 combo 拒绝 set_key。但创建路径 order.py:165 本来也只写值
    不动 key——命中即证明角色正确，无需知道 key 是什么。
    """

    def test_existing_employee_row_key_is_never_touched(self):
        # 负责雇员行 key 是未知值 ZP（非 ER）→ 只写 cs_code，key 保持 ZP，不记待校正。
        preset = {
            _parvw_id(4): _combo("ZP", ROLE_TEXTS_ZP),
            _parvw_id(5): _combo("WE", ROLE_TEXTS_ZP),
        }
        tx, raw = _make_transaction(
            preset, cs_code="CS001", sales_code="", role_texts=ROLE_TEXTS_ZP,
        )
        diffs: list[str] = []
        tx._edit_partners(_make_order(), diffs)

        self.assertEqual(raw.findById(_parvw_id(4)).key, "ZP")       # key 未被改写
        self.assertEqual(raw.findById(_partner_id(4)).text, "CS001")  # 值已写入
        self.assertEqual(diffs, ["Primary CS:(空)→CS001"])            # 无「待校正」

    def test_employee_row_outside_four_five_is_found(self):
        # 负责雇员行被 SAP 排到行 3 → 全表扫描仍能命中，CS 落行 3。
        preset = {
            _parvw_id(3): _combo("ZP", ROLE_TEXTS_ZP),
            _parvw_id(4): _combo("WE", ROLE_TEXTS_ZP),
        }
        tx, raw = _make_transaction(
            preset, cs_code="CS001", sales_code="", role_texts=ROLE_TEXTS_ZP,
        )
        diffs: list[str] = []
        tx._edit_partners(_make_order(), diffs)

        self.assertEqual(raw.findById(_partner_id(3)).text, "CS001")
        self.assertEqual(raw.findById(_parvw_id(3)).key, "ZP")

    def test_stale_cs_value_is_updated_without_key_change(self):
        preset = {
            _parvw_id(4): _combo("ZP", ROLE_TEXTS_ZP),
            _partner_id(4): _Element(text="OLD_CS"),
        }
        tx, raw = _make_transaction(
            preset, cs_code="CS001", sales_code="", role_texts=ROLE_TEXTS_ZP,
        )
        diffs: list[str] = []
        tx._edit_partners(_make_order(), diffs)

        self.assertEqual(raw.findById(_partner_id(4)).text, "CS001")
        self.assertEqual(diffs, ["Primary CS:OLD_CS→CS001"])

    def test_missing_employee_row_uses_key_from_entries(self):
        # 负责雇员行整行不存在 → 兜底新建。角色 key 取自选项表反查（ZP），非硬编码 ER。
        tx, raw = _make_transaction(cs_code="CS001", sales_code="", role_texts=ROLE_TEXTS_ZP)
        diffs: list[str] = []
        tx._edit_partners(_make_order(), diffs)

        self.assertEqual(raw.findById(_parvw_id(5)).key, "ZP")
        self.assertEqual(raw.findById(_partner_id(5)).text, "CS001")
        self.assertEqual(diffs, ["Primary CS:角色 (空)→ZP", "Primary CS:(空)→CS001"])

    def test_missing_employee_row_falls_back_to_key_guess(self):
        # 选项表不可用 → 只能用被实机否决的 ER 兜底；被拒时记待校正，绝不写脏数据。
        tx, raw = _make_transaction(
            cs_code="CS001", sales_code="",
            invalid_keys=frozenset({"ER"}), entries_available=False,
        )
        diffs: list[str] = []
        tx._edit_partners(_make_order(), diffs)

        self.assertEqual(raw.findById(_partner_id(5)).text, "")
        self.assertEqual(diffs, ["Primary CS:角色key ER 无效(待校正)"])


class LookupRoleKeyTest(unittest.TestCase):
    """显示文本 → PARVW key 的反查：唯一命中才采信。"""

    def test_unique_text_resolves_to_key(self):
        tx, _ = _make_transaction(role_texts=ROLE_TEXTS_ZP)
        self.assertEqual(
            tx._lookup_role_key(_parvw_id(0), OrderEditTransaction._EMPLOYEE_TEXTS), "ZP",
        )

    def test_ambiguous_text_returns_none(self):
        # "送达方"同时对应 WE 与 ZG → 歧义，弃权而不是猜一个。
        tx, _ = _make_transaction(role_texts=ROLE_TEXTS_AMBIGUOUS)
        self.assertIsNone(tx._lookup_role_key(_parvw_id(0), frozenset({"送达方"})))

    def test_entries_unavailable_returns_none(self):
        tx, _ = _make_transaction(entries_available=False)
        self.assertIsNone(
            tx._lookup_role_key(_parvw_id(0), OrderEditTransaction._EMPLOYEE_TEXTS),
        )


class ResolveGpcRowTest(unittest.TestCase):
    """GPC 行定位：已有 ZG 行优先，且命中时不再改角色。"""

    def test_existing_zg_row_wins_without_key_change(self):
        # ZG 被 SAP 排到行 8（不在 4/5）→ 仍命中，且返回 None 表示不必改角色。
        # 这正是 `GPC Code:角色key ZG 无效(待校正)` 的根因场景。
        preset = {_parvw_id(8): _combo("ZG")}
        tx, _ = _make_transaction(preset)
        self.assertEqual(tx._resolve_gpc_row(PARTNER_PREFIX), (8, None))

    def test_falls_back_to_create_row_with_key(self):
        # 无 ZG 行 → 落创建同源行位（负责雇员在行 4 → GPC 落行 5）并改角色。
        tx, _ = _make_transaction(_sap_determined_rows())
        self.assertEqual(tx._resolve_gpc_row(PARTNER_PREFIX), (5, "ZG"))

    def test_existing_zg_row_skips_set_key_entirely(self):
        # 端到端：ZG 行已在行 8 且角色列只读，仍能把 Buyer 编码写进去（零 set_key）。
        preset = {_parvw_id(8): _combo("ZG")}
        tx, raw = _make_transaction(preset, sales_code="", invalid_keys=frozenset({"ZG"}))
        diffs: list[str] = []
        tx._edit_partners(_make_order("GP001"), diffs)

        self.assertEqual(raw.findById(_partner_id(8)).text, "GP001")
        self.assertEqual(diffs, ["GPC Code:(空)→GP001"])


class FindEmployeeRowTest(unittest.TestCase):
    def test_matches_by_key_resolved_from_entries(self):
        preset = {_parvw_id(2): _combo("ZP", ROLE_TEXTS_ZP)}
        tx, _ = _make_transaction(preset, role_texts=ROLE_TEXTS_ZP)
        self.assertEqual(tx._find_employee_row(PARTNER_PREFIX), 2)

    def test_falls_back_to_display_text_without_entries(self):
        # 选项表不可用 → 退回显示文本扫描，老环境行为不回归。
        preset = {_parvw_id(2): _combo("ZP", ROLE_TEXTS_ZP)}
        tx, _ = _make_transaction(preset, role_texts=ROLE_TEXTS_ZP, entries_available=False)
        self.assertEqual(tx._find_employee_row(PARTNER_PREFIX), 2)

    def test_returns_none_when_absent(self):
        tx, _ = _make_transaction(_sap_determined_rows(employee_at_four=False))
        # _sap_determined_rows 用 ER→"负责雇员"，此处换成全非雇员行验证未命中。
        tx2, _ = _make_transaction({_parvw_id(0): _combo("AG")})
        self.assertIsNone(tx2._find_employee_row(PARTNER_PREFIX, max_rows=3))
        self.assertIsNotNone(tx._find_employee_row(PARTNER_PREFIX))


class FindEmptyPartnerRowTest(unittest.TestCase):
    def test_returns_first_empty_row(self):
        # 行0=ZG已占用，行1=空，断言返回 1。
        preset = {_parvw_id(0): _combo("ZG"), _partner_id(0): _Element(text="GP001")}
        tx, _ = _make_transaction(preset)
        self.assertEqual(tx._find_empty_partner_row(PARTNER_PREFIX, max_rows=4), 1)


if __name__ == "__main__":
    unittest.main()
