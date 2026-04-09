"""
SAP 操作数据模型定义

使用方法:
    from sap import SapConfig, OrderData, RevenueData, OperationFlags, HourData, SapResult
    # 或
    from sap.models import SapConfig, OrderData, RevenueData, OperationFlags, HourData, SapResult
"""

from dataclasses import dataclass, field


@dataclass
class SapConfig:
    """固定配置 -- 从 config.csv 加载, 订单间不变"""

    # 销售组织
    order_type: str
    sales_organization: str
    distribution_channels: str
    sales_office: str

    # 成本中心
    cost_center: str             # 原 salesGroup -> 订单成本中心
    sub_cost_center_cs: str      # 原 csCostCenter -> CS分包成本中心
    sub_cost_center_chm: str     # 原 chmCostCenter -> CHM分包成本中心
    sub_cost_center_phy: str     # 原 phyCostCenter -> PHY分包成本中心

    # 人员代码
    cs_code: str
    sales_code: str

    # 数据分类列表
    data_ae1: list[str] = field(default_factory=list)
    data_az2: list[str] = field(default_factory=list)

    # 阈值 (原硬编码)
    revenue_threshold: float = 35000.0
    plan_cost_min_threshold: float = 1000.0

    # A2 物料映射 (原 if/elif 硬编码)
    # key: 物料子码, value: (Item1000物料号, Item2000物料号)
    a2_material_mapping: dict[str, tuple[str, str]] = field(
        default_factory=lambda: {
            '405': ('T75-405-00', 'T20-405-00'),
            '430': ('T20-430-00', 'T75-430-00'),
            '441': ('T75-441-00', 'T20-441-00'),
        }
    )


@dataclass
class OrderData:
    """每单变动数据 -- 从 Excel 行读取"""

    sap_no: str
    project_no: str
    material_code: str
    currency_type: str
    exchange_rate: float
    amount_vat: float
    cost: float                        # 外币成本 (不含税)
    short_text: str
    long_text: str = ''
    global_partner_code: str = ''
    sales_name: str = ''               # 销售名称 (va01 联系人用)
    ecd: str = ''                      # 预计完成日期 (原 self.oneWeekday)


@dataclass
class RevenueData:
    """成本分配数据 -- 从订单数据直接读取 (不再内部计算)"""

    revenue: float                     # 未税收入
    revenue_cny: float                 # 人民币收入 (原 revenueForCny)

    # 部门成本
    chm_cost: float = 0.0
    phy_cost: float = 0.0

    # 部门收入
    chm_revenue: float = 0.0          # 原 chmRe
    phy_revenue: float = 0.0          # 原 phyRe

    # 拆分物料 (D2/D3/A2) 成本核算
    chm_cs_cost: float = 0.0          # 原 chmCsCostAccounting
    chm_lab_cost: float = 0.0         # 原 chmLabCostAccounting
    phy_cs_cost: float = 0.0          # 原 phyCsCostAccounting
    phy_lab_cost: float = 0.0         # 原 phyLabCostAccounting

    # 非拆分物料成本核算
    cs_cost: float = 0.0              # 原 csCostAccounting
    lab_cost: float = 0.0             # 原 labCostAccounting


@dataclass
class OperationFlags:
    """流程控制开关 -- 从 GUI CheckBox 读取"""

    va01: bool = True
    va02: bool = True
    vf01: bool = False
    vf03: bool = False
    save: bool = True
    lab_cost: bool = False
    plan_cost: bool = False
    cs: bool = True
    chm: bool = True
    phy: bool = True
    every: bool = False                # 每次创建新 SAP 对象
    contact: bool = True               # 是否添加联系人


@dataclass
class HourData:
    """工时记录数据 -- 独立模块"""

    staff_id: str
    week: str
    allocated_day: str
    order_no: str
    item: str
    material_code: str
    allocated_hours: float
    office_time: float


@dataclass
class SapResult:
    """SAP 操作统一返回结果

    替代原 res = {'flag': 1, 'msg': ''} 的 dict 模式,
    提供类型安全和 IDE 自动补全。

    使用方法:
        result = SapResult()                    # 默认成功
        result = SapResult.fail('错误信息')     # 快速创建失败结果
        result.append_message('追加信息')       # 追加消息

        if result.success:
            print(result.order_no)
    """

    success: bool = True
    message: str = ''
    order_no: str = ''
    proforma_no: str = ''
    sap_amount_vat: str = ''

    @staticmethod
    def fail(msg: str) -> 'SapResult':
        """快速创建失败结果"""
        return SapResult(success=False, message=msg)

    def append_message(self, msg: str) -> None:
        """追加消息, 用分号分隔"""
        if self.message:
            self.message = f"{self.message};{msg}"
        else:
            self.message = msg
