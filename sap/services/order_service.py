"""订单服务。"""

from sap.models import CostOptions, OrderData, PartnerOptions, RevenueData, SapConfig, SapResult
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

    def fill_lab_cost(self, order: OrderData, revenue: RevenueData) -> SapResult:
        """填写 Data B 的人工成本。"""
        return self.transaction.fill_lab_cost(order, revenue)

    def apply_plan_cost(
        self,
        order: OrderData,
        revenue: RevenueData,
        *,
        cost_options: CostOptions | None = None,
    ) -> SapResult:
        """写入计划成本。"""
        return self.transaction.apply_plan_cost(order, revenue, cost_options or CostOptions())

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
