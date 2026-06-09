"""SAP 模块的数据模型。"""

from dataclasses import dataclass, field


@dataclass(slots=True)
class SapConfig:
    """固定配置，通常从配置文件加载。"""

    # 销售组织相关固定参数。
    order_type: str
    sales_organization: str
    distribution_channels: str
    sales_office: str
    sales_group: str

    # 成本中心相关固定参数。

    sub_cost_center_cs: str
    sub_cost_center_chm: str
    sub_cost_center_phy: str

    # 合作伙伴代码。
    cs_code: str
    sales_code: str

    # DATA A 判定用的客户清单。
    data_ae1: list[str] = field(default_factory=list)
    data_az2: list[str] = field(default_factory=list)

    # DATA B TUV IC 订单判定用的客户清单（命中则 VA01 写 IC_TRANSAKTION=O1）。
    data_b_tuv: list[str] = field(default_factory=list)

    # 阈值控制。
    revenue_threshold: float = 35000.0
    plan_cost_min_threshold: float = 1000.0

    # A2 物料拆分映射，key 为子码，value 为 (item1000, item2000)。
    a2_material_mapping: dict[str, tuple[str, str]] = field(
        default_factory=lambda: {
            "405": ("T75-405-00", "T20-405-00"),
            "430": ("T20-430-00", "T75-430-00"),
            "441": ("T75-441-00", "T20-441-00"),
        }
    )


@dataclass(slots=True)
class OrderItemData:
    """Order item row data for VA02."""

    item: str = ""
    material_code: str = ""
    long_text: str = ""
    revenue: float = 0.0
    quantity: str = "1"
    unit: str = "pu"


@dataclass(slots=True)
class OrderData:
    """订单业务数据。"""

    # SAP 客户号。
    sap_no: str
    # 客户项目号。
    project_no: str
    # 币种。
    currency_type: str
    # 汇率，仅非 CNY 场景使用。
    exchange_rate: float
    # 订单头短文本。
    short_text: str
    # 产品子类别，用于少数 VA01 头部字段的条件写入。
    product_sub_category: str = ""
    # 全球合作伙伴代码。
    global_partner_code: str = ""
    # 销售名称，当前主要用来判断是否需要补销售伙伴。
    sales_name: str = ""
    # 预计完成日期。
    ecd: str = ""
    # order_center
    order_cost_center: str = ""
    # Sales Group
    sales_group: str = ""

    # 将item集成在这里
    items: list[OrderItemData] = field(default_factory=list)

    # 含税金额，当前保留给上层业务使用。
    amount_vat: float = 0.0

@dataclass(slots=True)
class RevenueData:
    """营收和成本分配数据。"""

    # 未税收入。
    revenue: float
    # 人民币收入。
    revenue_cny: float

    # Data B 使用的部门成本。
    chm_cost: float = 0.0
    phy_cost: float = 0.0

    # A2 双 item 使用的部门收入。
    chm_revenue: float = 0.0
    phy_revenue: float = 0.0

    # 拆分物料计划成本数据。
    chm_cs_cost: float = 0.0
    chm_lab_cost: float = 0.0
    phy_cs_cost: float = 0.0
    phy_lab_cost: float = 0.0

    # 普通物料计划成本数据。
    cs_cost: float = 0.0
    lab_cost: float = 0.0


@dataclass(slots=True, frozen=True)
class PartnerOptions:
    """VA01 合作伙伴写入选项。"""

    # 是否补联系人。
    add_contact: bool = True
    # 是否补销售伙伴。
    add_sales_partner: bool = True


@dataclass(slots=True, frozen=True)
class CostOptions:
    """计划成本写入选项。"""

    # 是否写入 CS 计划成本。
    include_cs: bool = True
    # 是否写入 CHM 计划成本。
    include_chm: bool = True
    # 是否写入 PHY 计划成本。
    include_phy: bool = True


@dataclass(slots=True)
class DataBEntry:
    """Data B 单行明细：写入 VA01/VA02 抬头 Data B 页签的一行成本记录。

    - performer_cost_center: 执行部门成本中心（写入 TABL-KOSTL）。
    - rate_cost_center: 费率成本中心（写入 TABD-KOSTL）；为空时取 performer_cost_center。
    - amount: 固定价格（写入 TABD-FESTPREIS）。
    - item: 关联 SAP item 号（写入 TABD-POSNR）；为空时由 SAP 默认行为兜底；
            多 item（如 "1000;3000"）由下游裁剪取第一个。
    """

    performer_cost_center: str
    rate_cost_center: str
    amount: float
    item: str = ""


@dataclass(slots=True)
class PlanCostEntry:
    """Plan Cost 单条明细：写入计划成本编辑器表格的一行。

    - cost_center: 成本中心（写入 RK70L-HERK2）。
    - category: SAP 成本类别（写入 RK70L-HERK3），'FREMDL' 分包费用 / 'T01AST' 工时。
    - amount: FREMDL 写金额、T01AST 写工时（写入 RK70L-MENGE）。
    """

    cost_center: str
    category: str
    amount: float


@dataclass(slots=True)
class HourData:
    """工时记录数据。"""

    # 员工工号。
    staff_id: str
    # 周次。
    week: str
    # 分配日期。
    allocated_day: str
    # 订单号。
    order_no: str
    # Item 编号。
    item: str
    # 活动/物料代码。
    material_code: str
    # 分配工时。
    allocated_hours: float
    # 办公时间。
    office_time: float


@dataclass(slots=True)
class SapResult:
    """统一的 SAP 操作结果。"""

    # 是否成功。
    success: bool = True
    # 错误或补充消息。
    message: str = ""
    # 订单号。
    order_no: str = ""
    # 形式发票号。
    proforma_no: str = ""
    # SAP 实际金额文本。
    sap_amount_vat: str = ""
    # 当前结果对应的步骤标识。
    step: str = ""

    @staticmethod
    def fail(msg: str, *, step: str = "") -> "SapResult":
        """创建失败结果，统一失败出口。"""
        return SapResult(success=False, message=msg, step=step)

    def append_message(self, msg: str) -> None:
        """在原消息后追加信息，便于保留多个警告/失败原因。"""
        if self.message:
            self.message = f"{self.message};{msg}"
        else:
            self.message = msg
