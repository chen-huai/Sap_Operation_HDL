"""SAP 业务规则与纯逻辑计算。"""

from __future__ import annotations

from dataclasses import dataclass

from sap.models import CostOptions, OrderData, RevenueData, SapConfig


@dataclass(slots=True, frozen=True)
class LabCostEntry:
    """Data B 人工成本行。"""

    # 在 SAP 表格中的目标行号。
    row: int
    # 执行部门成本中心。
    performer_cost_center: str
    # 费率成本中心。
    rate_cost_center: str
    # 金额。
    amount: float


@dataclass(slots=True, frozen=True)
class PlanCostEntry:
    """计划成本行。"""

    # 在计划成本表中的行号。
    row: int
    # 成本中心。
    cost_center: str
    # SAP 成本类别，如 T01AST / FREMDL。
    category: str
    # 金额。
    amount: float


def is_a2_material(material_code: str) -> bool:
    """判断是否为 A2 双 item 物料。"""
    return "A2" in material_code


def is_d_split_material(material_code: str) -> bool:
    """判断是否为 D2/D3 拆分物料。"""
    return "D2" in material_code or "D3" in material_code


def is_t75_material(material_code: str) -> bool:
    """判断是否属于 T75 系列物料。"""
    return "T75" in material_code


def has_430_subcode(material_code: str) -> bool:
    """判断物料代码中是否包含 430 子码。"""
    return "430" in material_code


def resolve_a2_materials(material_code: str, config: SapConfig) -> tuple[str, str]:
    """根据 A2 子码解析两个 item 的物料号。"""
    for sub_code, materials in config.a2_material_mapping.items():
        if sub_code in material_code:
            return materials
    return config.a2_material_mapping.get("441", ("T75-441-00", "T20-441-00"))


def resolve_data_a_key(order: OrderData, config: SapConfig) -> str:
    # DATA A 的判定只保留纯规则，不夹带任何 SAP 页面操作。
    if is_d_split_material(order.material_code):
        return "E1" if order.sap_no in config.data_ae1 else "Z0"
    if order.sap_no in config.data_az2:
        return "Z2"
    return "00"


def should_fill_auftragswert(revenue: RevenueData, config: SapConfig) -> bool:
    """判断是否需要填写订单价值。"""
    return revenue.revenue_cny >= config.revenue_threshold


def should_apply_plan_cost(revenue: RevenueData, config: SapConfig) -> bool:
    """判断当前订单是否需要进入计划成本流程。"""
    return revenue.revenue_cny >= config.plan_cost_min_threshold


def build_lab_cost_entries(order: OrderData, revenue: RevenueData, config: SapConfig) -> list[LabCostEntry]:
    # 这里返回“应该写什么”，由事务层决定“写到哪里”。
    if is_a2_material(order.material_code) or is_d_split_material(order.material_code):
        return [
            LabCostEntry(0, config.sub_cost_center_chm, config.sub_cost_center_chm, revenue.chm_cost),
            LabCostEntry(1, config.sub_cost_center_phy, config.sub_cost_center_phy, revenue.phy_cost),
        ]
    if "T20" in order.material_code or has_430_subcode(order.material_code):
        return [LabCostEntry(0, config.sub_cost_center_phy, config.sub_cost_center_phy, revenue.phy_cost)]
    return [LabCostEntry(0, config.sub_cost_center_chm, config.sub_cost_center_chm, revenue.chm_cost)]


def build_split_plan_cost_entries(
    revenue: RevenueData,
    config: SapConfig,
    options: CostOptions,
) -> list[PlanCostEntry]:
    # D2/D3 这类拆分物料会把 CS、CHM、PHY 拆成独立行。
    entries: list[PlanCostEntry] = []
    row = 0
    if options.include_cs:
        cs_total = round(float(revenue.chm_cs_cost) + float(revenue.phy_cs_cost), 0)
        if cs_total > 0:
            entries.append(PlanCostEntry(row, config.sub_cost_center_cs, "T01AST", cs_total))
            row += 1
    if options.include_chm:
        chm_lab = round(float(revenue.chm_lab_cost), 0)
        if chm_lab > 0:
            entries.append(PlanCostEntry(row, config.sub_cost_center_chm, "T01AST", chm_lab))
            row += 1
    if options.include_phy:
        phy_lab = round(float(revenue.phy_lab_cost), 0)
        if phy_lab > 0:
            entries.append(PlanCostEntry(row, config.sub_cost_center_phy, "T01AST", phy_lab))
            row += 1
    return entries


def build_single_plan_cost_entries(
    order: OrderData,
    revenue: RevenueData,
    config: SapConfig,
    options: CostOptions,
) -> list[PlanCostEntry]:
    # 普通物料只有一套计划成本视图，lab 成本中心取决于物料归属。
    entries: list[PlanCostEntry] = []
    row = 0

    if options.include_cs:
        cs_cost = round(float(revenue.cs_cost), 0)
        if cs_cost > 0:
            entries.append(PlanCostEntry(row, config.sub_cost_center_cs, "T01AST", cs_cost))
            row += 1

    if options.include_chm or options.include_phy:
        lab_cost = round(float(revenue.lab_cost), 0)
        if lab_cost > 0:
            if is_t75_material(order.material_code):
                if options.include_chm:
                    entries.append(PlanCostEntry(row, config.sub_cost_center_chm, "T01AST", lab_cost))
                    row += 1
            elif options.include_phy:
                entries.append(PlanCostEntry(row, config.sub_cost_center_phy, "T01AST", lab_cost))
                row += 1

    return entries


def build_fremdl_entry(row: int, order: OrderData, config: SapConfig) -> PlanCostEntry | None:
    """为外包成本生成 FREMDL 行；没有外包成本时返回 None。"""
    if order.cost <= 0:
        return None
    return PlanCostEntry(row, config.sub_cost_center_cs, "FREMDL", order.cost)
