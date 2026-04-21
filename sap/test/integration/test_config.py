"""真实 SAP 集成测试参数。"""

from sap import HourData, OrderData, RevenueData, SapConfig


ALLOW_SAVE = False

SAP_CONFIG = SapConfig(
    order_type="ZOR",
    sales_organization="",
    distribution_channels="",
    sales_office="",
    cost_center="",
    sub_cost_center_cs="",
    sub_cost_center_chm="",
    sub_cost_center_phy="",
    cs_code="",
    sales_code="",
    data_ae1=[],
    data_az2=[],
)

ORDER = OrderData(
    sap_no="",
    project_no="",
    material_code="",
    currency_type="CNY",
    exchange_rate=1.0,
    cost=0.0,
    short_text="",
    long_text="",
    global_partner_code="",
    sales_name="",
    ecd="",
)

REVENUE = RevenueData(
    revenue=0.0,
    revenue_cny=0.0,
    chm_cost=0.0,
    phy_cost=0.0,
    chm_revenue=0.0,
    phy_revenue=0.0,
    chm_cs_cost=0.0,
    chm_lab_cost=0.0,
    phy_cs_cost=0.0,
    phy_lab_cost=0.0,
    cs_cost=0.0,
    lab_cost=0.0,
)

HOUR = HourData(
    staff_id="",
    week="",
    allocated_day="",
    order_no="",
    item="",
    material_code="",
    allocated_hours=0.0,
    office_time=0.0,
)
