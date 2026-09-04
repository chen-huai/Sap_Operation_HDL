"""订单服务。"""

from sap.models import (
    DataBEntry,
    OrderData,
    PartnerOptions,
    PlanCostEntry,
    RevenueData,
    SapConfig,
    SapResult,
)
from sap.session import SapSession
from sap.transactions.order import OrderTransaction


class OrderService:
    """订单域服务，提供清晰的显式调用入口。"""

    def __init__(self, session: SapSession, config: SapConfig):
        """基于共享会话初始化订单事务服务。"""
        self.transaction = OrderTransaction(session, config)

    def create_order(
        self,
        order: OrderData,
        revenue: RevenueData,
        *,
        partner_options: PartnerOptions | None = None,
    ) -> SapResult:
        """创建订单头。"""
        return self.transaction.create(order, revenue, partner_options or PartnerOptions())

    def open_order(self, order_no: str) -> SapResult:
        """打开已有订单。"""
        return self.transaction.open(order_no)

    def add_items(self, order: OrderData, revenue: RevenueData) -> SapResult:
        """为当前订单添加 item。"""
        return self.transaction.add_items(order, revenue)

    def update_items(self, order: OrderData, revenue: RevenueData) -> SapResult:
        """更新当前订单 item。"""
        return self.transaction.update_items(order, revenue)

    def fill_lab_cost_entries(
        self,
        entries: list[DataBEntry],
        order: OrderData,
    ) -> SapResult:
        """按已计算好的 Data B 明细写入人工成本。

        Args:
            entries: Data B 明细列表（DataBEntry）。
            order: 订单数据，用于读取 sales_group 决定是否写入 item 号。

        Returns:
            SapResult: SAP 写入结果。
        """
        return self.transaction.fill_lab_cost_entries(entries, order)

    def fill_order_value(self, order: OrderData) -> SapResult:
        """回填订单价值(AUFTRAGSWERT)：Σ SAP item 未税净值 × 汇率，达阈值才写。

        须在 item 全部录入后调用（读 SAP 概览净值，与创建/编辑同口径）。
        """
        return self.transaction.fill_order_value(order)

    def apply_plan_cost_entries(
        self,
        entries: list[PlanCostEntry],
        *,
        focus_row: int = 0,
        target_item: str | None = None,
    ) -> SapResult:
        """按已计算好的计划成本明细写入计划成本。

        Args:
            entries: 计划成本明细列表（PlanCostEntry）。
            target_item: 目标 item 号；传入时按号在 SAP 概览页实时定位物理行（推荐）。
            focus_row: target_item 为空时的兜底行号。

        Returns:
            SapResult: SAP 写入结果。
        """
        return self.transaction.apply_plan_cost_entries(
            entries, focus_row=focus_row, target_item=target_item
        )

    def save(self, info: str) -> SapResult:
        """保存当前订单页面。"""
        return self.transaction.save(info)

    def lock(self, order_no: str) -> SapResult:
        """打开订单并执行锁定。"""
        open_result = self.open_order(order_no)
        if not open_result.success:
            return open_result
        return self.transaction.set_lock_state(unlocked=False)

    def unlock(self, order_no: str) -> SapResult:
        """打开订单并执行解锁。"""
        open_result = self.open_order(order_no)
        if not open_result.success:
            return open_result
        return self.transaction.set_lock_state(unlocked=True)
