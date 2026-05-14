"""SAP 业务规则与纯逻辑计算。"""

from __future__ import annotations

from sap.models import OrderData, RevenueData, SapConfig


def resolve_data_a_key(order: OrderData, config: SapConfig) -> str:
    """根据 SAP 客户号判定 DATA A 标签；不再依赖物料代码。"""
    if order.sap_no in config.data_ae1:
        return "E1"
    if order.sap_no in config.data_az2:
        return "Z2"
    return "00"


def should_fill_auftragswert(revenue: RevenueData, config: SapConfig) -> bool:
    """判断是否需要在订单头写入订单价值。"""
    return revenue.revenue_cny >= config.revenue_threshold
