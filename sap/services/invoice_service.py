"""发票服务。"""

from sap.session import SapSession
from sap.transactions.invoice import InvoiceTransaction


class InvoiceService:
    """发票域服务。"""

    def __init__(self, session: SapSession):
        """基于共享会话初始化发票事务服务。"""
        self.transaction = InvoiceTransaction(session)

    def create_proforma(self):
        """创建形式发票。"""
        return self.transaction.create_proforma()

    def display_proforma(self):
        """查看形式发票。"""
        return self.transaction.display_proforma()
