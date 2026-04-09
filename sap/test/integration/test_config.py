"""
SAP 集成测试参数配置

使用方法:
    1. 填写下方测试参数（未填写的测试会自动跳过）
    2. 确保 SAP GUI 已登录
    3. 运行: pytest sap/test/integration/test_xxx.py -v -s
"""

from sap.models import SapConfig, OrderData, RevenueData, OperationFlags, HourData


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  安全控制
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

ALLOW_SAVE = False      # ⚠️ True = 允许在生产环境执行保存操作，请谨慎


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  SAP 固定配置（对应 config_sap.csv）
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

SAP_CONFIG = SapConfig(
    order_type='ZOR',
    sales_organization='',          # ← 填写
    distribution_channels='',       # ← 填写
    sales_office='',                # ← 填写
    cost_center='',                 # ← 填写
    sub_cost_center_cs='',          # ← 填写
    sub_cost_center_chm='',         # ← 填写
    sub_cost_center_phy='',         # ← 填写
    cs_code='',                     # ← 填写
    sales_code='',                  # ← 填写
    data_ae1=[],
    data_az2=[],
)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  操作控制开关
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

FLAGS = OperationFlags(
    va01=True,
    va02=True,
    vf01=False,
    vf03=False,
    save=False,
    lab_cost=False,
    plan_cost=False,
    cs=True,
    chm=True,
    phy=True,
    every=False,
    contact=True,
)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  订单数据（VA01 创建订单时使用）
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

ORDER = OrderData(
    sap_no='',                      # ← 必填: 客户 SAP 编号
    project_no='',                  # ← 必填: 项目编号
    material_code='',               # ← 必填: 物料代码，如 'T75-405-00'
    currency_type='CNY',
    exchange_rate=1.0,
    amount_vat=0.0,
    cost=0.0,
    short_text='',                  # ← 填写: 短文本
    long_text='',
    global_partner_code='',         # ← 填写: 全球合作伙伴代码
    sales_name='',
    ecd='',                         # ← 填写: 预计完成日期，如 '2026.12.31'
)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  营收数据
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

REVENUE = RevenueData(
    revenue=0.0,                    # ← 填写: 未税收入
    revenue_cny=0.0,                # ← 填写: 人民币收入
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


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  已有订单号（VA02/open/unlock/VF01/VF03 使用）
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

EXISTING_ORDER_NO = ''              # ← 填写已有订单号，如 '60001234'


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
#  工时数据
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

HOUR = HourData(
    staff_id='',                    # ← 填写: 员工 ID
    week='',                        # ← 填写: 周次
    allocated_day='',               # ← 填写: 分配日期
    order_no='',                    # ← 填写: 订单号
    item='',                        # ← 填写: Item 号
    material_code='',               # ← 填写: 物料代码
    allocated_hours=0.0,            # ← 填写: 分配工时
    office_time=0.0,                # ← 填写: 办公时间
)
