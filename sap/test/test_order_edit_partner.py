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

# SAP 侧 PARVW key → combo 显示文本，供角色自证分支的桩联动。
ROLE_TEXTS = {"ZG": "GPC", "ER": "负责雇员", "VE": "销售员", "WE": "送达方"}


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


class _ComboElement(_Element):
    """PARVW 下拉框桩：写 key 时按 role_texts 联动显示文本，模拟真实 combo。

    `invalid_keys` 里的 key 赋值抛错，模拟 SAP 对不在选项列表中的 key 的拒绝；
    `role_texts` 可覆盖映射，用于构造"key 合法但角色对不上"的场景。
    """

    def __init__(self, role_texts: dict[str, str], invalid_keys: frozenset[str], key: str = ""):
        # 须先于 super()，父类 __init__ 里的 `self.key = ...` 会走下面的 setter。
        self._role_texts = role_texts
        self._invalid_keys = invalid_keys
        super().__init__(key=key)

    @property
    def key(self) -> str:
        return self._key

    @key.setter
    def key(self, value: str) -> None:
        if value in self._invalid_keys:
            raise ValueError(f"invalid PARVW key: {value}")
        self._key = value
        self.text = self._role_texts.get(value, "") if value else ""


class _RawSession:
    """按 element_id 缓存控件的最小 raw session；未知 id 返回空控件。

    PARVW 下拉框走 `_ComboElement`（key→text 联动），其余走普通 `_Element`。
    """

    def __init__(
        self,
        preset: dict[str, _Element] | None = None,
        *,
        role_texts: dict[str, str] | None = None,
        invalid_keys: frozenset[str] = frozenset(),
    ):
        self._cache: dict[str, _Element] = preset or {}
        self._role_texts = ROLE_TEXTS if role_texts is None else role_texts
        self._invalid_keys = invalid_keys

    def findById(self, element_id: str) -> _Element:
        if element_id not in self._cache:
            self._cache[element_id] = (
                _ComboElement(self._role_texts, self._invalid_keys)
                if PARVW_MARK in element_id
                else _Element()
            )
        return self._cache[element_id]


def _combo(key: str, role_texts: dict[str, str] | None = None) -> _ComboElement:
    """构造一个已带角色 key 的 PARVW 桩控件（显示文本按映射联动）。"""
    return _ComboElement(ROLE_TEXTS if role_texts is None else role_texts, frozenset(), key=key)


def _make_transaction(
    preset: dict[str, _Element] | None = None,
    *,
    cs_code: str = "",
    sales_code: str = "SA001",
    role_texts: dict[str, str] | None = None,
    invalid_keys: frozenset[str] = frozenset(),
):
    raw = _RawSession(preset, role_texts=role_texts, invalid_keys=invalid_keys)
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

    def test_invalid_parvw_key_records_pending(self):
        # 闸①：key 不在 combo 选项里 → 记待校正，且绝不写编码。
        tx, raw = _make_transaction(invalid_keys=frozenset({"ER"}))
        diffs: list[str] = []
        tx._sync_partner_row(
            PARTNER_PREFIX, 4, "ER", "CS001", field="Primary CS", diffs=diffs,
            expect_texts=OrderEditTransaction._EMPLOYEE_TEXTS,
        )

        self.assertEqual(raw.findById(_partner_id(4)).text, "")
        self.assertEqual(diffs, ["Primary CS:角色key ER 无效(待校正)"])

    def test_gpc_display_text_allows_unknown_key(self):
        # 编辑屏实机存在：GPC 行显示文本正确，但 ZG key 被 combo 拒绝。此时不改 key，直接写 Buyer 值。
        preset = {_parvw_id(5): _ComboElement({"ZP": "GPC"}, frozenset({"ZG"}), key="ZP")}
        tx, raw = _make_transaction(preset, invalid_keys=frozenset({"ZG"}))
        diffs: list[str] = []
        tx._sync_partner_row(
            PARTNER_PREFIX, 5, "ZG", "GP001", field="GPC Code", diffs=diffs,
            expect_texts=OrderEditTransaction._GPC_TEXTS,
        )

        self.assertEqual(raw.findById(_parvw_id(5)).key, "ZP")
        self.assertEqual(raw.findById(_partner_id(5)).text, "GP001")
        self.assertEqual(diffs, ["GPC Code:(空)→GP001"])

    def test_unexpected_role_text_skips_write(self):
        # 闸②：key 合法但对应角色不是"负责雇员" → 不写编码。
        tx, raw = _make_transaction(role_texts={"ER": "开票方"})
        diffs: list[str] = []
        tx._sync_partner_row(
            PARTNER_PREFIX, 4, "ER", "CS001", field="Primary CS", diffs=diffs,
            expect_texts=OrderEditTransaction._EMPLOYEE_TEXTS,
        )

        self.assertEqual(raw.findById(_partner_id(4)).text, "")
        self.assertEqual(
            diffs,
            ["Primary CS:角色 (空)→ER", "Primary CS:角色key ER 对应「开票方」非预期(待校正)"],
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
            _parvw_id(4): _combo("ZP", {"ZP": "负责雇员"}),
            _parvw_id(5): _combo("WE"),
        }
        tx, raw = _make_transaction(preset, cs_code="CS001", sales_code="")
        diffs: list[str] = []
        tx._edit_partners(_make_order(), diffs)

        self.assertEqual(raw.findById(_parvw_id(4)).key, "ZP")       # key 未被改写
        self.assertEqual(raw.findById(_partner_id(4)).text, "CS001")  # 值已写入
        self.assertEqual(diffs, ["Primary CS:(空)→CS001"])            # 无「待校正」

    def test_employee_row_outside_four_five_is_found(self):
        # 负责雇员行被 SAP 排到行 3 → 全表扫描仍能命中，CS 落行 3。
        preset = {
            _parvw_id(3): _combo("ZP", {"ZP": "负责雇员"}),
            _parvw_id(4): _combo("WE"),
        }
        tx, raw = _make_transaction(preset, cs_code="CS001", sales_code="")
        diffs: list[str] = []
        tx._edit_partners(_make_order(), diffs)

        self.assertEqual(raw.findById(_partner_id(3)).text, "CS001")
        self.assertEqual(raw.findById(_parvw_id(3)).key, "ZP")

    def test_stale_cs_value_is_updated_without_key_change(self):
        preset = {
            _parvw_id(4): _combo("ZP", {"ZP": "负责雇员"}),
            _partner_id(4): _Element(text="OLD_CS"),
        }
        tx, raw = _make_transaction(preset, cs_code="CS001", sales_code="")
        diffs: list[str] = []
        tx._edit_partners(_make_order(), diffs)

        self.assertEqual(raw.findById(_partner_id(4)).text, "CS001")
        self.assertEqual(diffs, ["Primary CS:OLD_CS→CS001"])

    def test_missing_employee_row_falls_back_to_key_guess(self):
        # 负责雇员行整行不存在 → 走兜底分支，凭推断 key 新建；ER 被拒时记待校正不写脏数据。
        tx, raw = _make_transaction(cs_code="CS001", sales_code="", invalid_keys=frozenset({"ER"}))
        diffs: list[str] = []
        tx._edit_partners(_make_order(), diffs)

        self.assertEqual(raw.findById(_partner_id(5)).text, "")
        self.assertEqual(diffs, ["Primary CS:角色key ER 无效(待校正)"])


class FindEmployeeRowTest(unittest.TestCase):
    def test_matches_by_display_text_not_key(self):
        preset = {_parvw_id(2): _combo("ZP", {"ZP": "负责雇员"})}
        tx, _ = _make_transaction(preset)
        self.assertEqual(tx._find_employee_row(PARTNER_PREFIX), 2)

    def test_returns_none_when_absent(self):
        tx, _ = _make_transaction(_sap_determined_rows(employee_at_four=False))
        # _sap_determined_rows 用 ER→"负责雇员"，此处换成全非雇员行验证未命中。
        tx2, _ = _make_transaction({_parvw_id(0): _combo("AG", {"AG": "售达方"})})
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
