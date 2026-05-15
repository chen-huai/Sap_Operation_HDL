"""订单服务。"""

from sap.models import OrderData, PartnerOptions, RevenueData, SapConfig, SapResult
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

    def fill_lab_cost_entries(self, entries, *, auftragswert_cny: float = 0.0) -> SapResult:
        """
        按已计算好的 Data B 明细写入人工成本。

        Args:
            entries: Data B 明细列表，每项包含 performer_cost_center、rate_cost_center、amount。
            auftragswert_cny: 所有 item 加和金额（CNY）。≥ 阈值时回填订单价值字段。

        Returns:
            SapResult: SAP 写入结果。
        """
        return self.transaction.fill_lab_cost_entries(entries, auftragswert_cny=auftragswert_cny)

    def apply_plan_cost_entries(self, entries, *, focus_row: int = 0) -> SapResult:
        """
        按已计算好的计划成本明细写入计划成本。

        Args:
            entries: 计划成本明细列表，每项包含 cost_center、category、amount。
            focus_row: SAP item 表格中需要进入计划成本界面的行号。

        Returns:
            SapResult: SAP 写入结果。
        """
        return self.transaction.apply_plan_cost_entries(entries, focus_row=focus_row)

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
