"""解耦后的 SAP 模块公开入口。

对外只暴露：
1. 数据模型
2. 会话对象
3. 按业务域划分的服务对象
"""

from sap.models import (
    CostOptions,
    HourData,
    OrderData,
    OrderItemData,
    PartnerOptions,
    RevenueData,
    SapConfig,
    SapResult,
)
from sap.services.hour_service import HourService
from sap.services.invoice_service import InvoiceService
from sap.services.order_service import OrderService
from sap.session import SapSession

__all__ = [
    "CostOptions",
    "HourData",
    "HourService",
    "InvoiceService",
    "OrderData",
    "OrderItemData",
    "OrderService",
    "PartnerOptions",
    "RevenueData",
    "SapConfig",
    "SapResult",
    "SapSession",
]
