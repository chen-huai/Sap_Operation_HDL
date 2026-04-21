"""工时服务。"""

from sap.models import HourData
from sap.session import SapSession
from sap.transactions.hours import HourTransaction


class HourService:
    """工时域服务。"""

    def __init__(self, session: SapSession):
        """基于共享会话初始化工时事务服务。"""
        self.transaction = HourTransaction(session)

    def login(self, hour: HourData):
        """登录工时系统。"""
        return self.transaction.login(hour)

    def record(self, hour: HourData, *, row_num: int = 0, max_rows: int = 20):
        """录入一条工时。"""
        return self.transaction.record(hour, row_num=row_num, max_rows=max_rows)

    def save(self, *, max_retries: int = 14):
        """保存工时。"""
        return self.transaction.save(max_retries=max_retries)
