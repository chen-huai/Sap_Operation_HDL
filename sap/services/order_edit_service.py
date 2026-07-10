"""订单编辑服务（VA02 字段对比更新）。

与 OrderService(创建) 业务分割，提供编辑域的显式调用入口；
底层委托 OrderEditTransaction 完成"读-对比-改"。
"""

from sap.models import (
    DataBEntry,
    OrderData,
    PlanCostEntry,
    SapConfig,
    SapResult,
    SubEditEntry,
)
from sap.session import SapSession
from sap.transactions.order_edit import OrderEditTransaction


class OrderEditService:
    """订单编辑域服务。"""

    def __init__(self, session: SapSession, config: SapConfig):
        """基于共享会话初始化编辑事务服务。"""
        self.transaction = OrderEditTransaction(session, config)

    def open_order(self, order_no: str) -> SapResult:
        """打开已有订单（VA02）。"""
        return self.transaction.open(order_no)

    def edit_header(self, order: OrderData, diffs: list[str]) -> SapResult:
        """对比并更新订单抬头字段。"""
        return self.transaction.edit_header(order, diffs)

    def edit_items(self, order: OrderData, diffs: list[str]) -> SapResult:
        """对比并更新 item 行。"""
        return self.transaction.edit_items(order, diffs)

    def edit_data_b(
        self,
        entries: list[DataBEntry],
        sub_edit_entries: list[SubEditEntry],
        order: OrderData,
        diffs: list[str],
    ) -> SapResult:
        """对比并更新 Data B 行。"""
        return self.transaction.edit_data_b(entries, sub_edit_entries, order, diffs)

    def edit_order_value(self, order: OrderData, diffs: list[str]) -> SapResult:
        """对比并更新订单价值(AUFTRAGSWERT)：Σ SAP item 未税净值 × 汇率。"""
        return self.transaction.edit_order_value(order, diffs)

    def edit_plan_cost(
        self,
        entries: list[PlanCostEntry],
        diffs: list[str],
        *,
        target_item: str,
    ) -> SapResult:
        """对比并更新指定 SAP item 的计划成本（按 item 号定位，SAP 无此 item 则跳过）。"""
        return self.transaction.edit_plan_cost(entries, diffs, target_item=target_item)

    def save(self, info: str) -> SapResult:
        """保存当前订单页面。"""
        return self.transaction.save(info)
