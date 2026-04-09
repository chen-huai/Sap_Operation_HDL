"""
测试工厂函数和辅助工具

使用方法:
    from sap.test.helpers import make_config, make_order, make_revenue, make_flags, make_hour
    from sap.test.helpers import setup_session_mock
"""

import sys
from unittest.mock import MagicMock, PropertyMock

# Mock win32com — 允许在非 Windows 平台导入 sap 模块
if 'win32com' not in sys.modules:
    sys.modules['win32com'] = MagicMock()
    sys.modules['win32com.client'] = MagicMock()

from sap.models import (
    SapConfig, OrderData, RevenueData, OperationFlags, HourData, SapResult,
)
from sap.function import Sap


def make_config(**overrides) -> SapConfig:
    """创建测试用 SapConfig"""
    defaults = dict(
        order_type='ZOR',
        sales_organization='3002',
        distribution_channels='10',
        sales_office='1000',
        cost_center='1100',
        sub_cost_center_cs='1101',
        sub_cost_center_chm='1102',
        sub_cost_center_phy='1103',
        cs_code='CS001',
        sales_code='SA001',
        data_ae1=['100001'],
        data_az2=['200001'],
    )
    defaults.update(overrides)
    return SapConfig(**defaults)


def make_order(**overrides) -> OrderData:
    """创建测试用 OrderData"""
    defaults = dict(
        sap_no='123456',
        project_no='PRJ-001',
        material_code='T75-405-00',
        currency_type='CNY',
        exchange_rate=1.0,
        amount_vat=10000.0,
        cost=5000.0,
        short_text='Test Short Text',
        long_text='Test Long Text',
        global_partner_code='GP001',
        sales_name='Zhang San',
        ecd='2026.12.31',
    )
    defaults.update(overrides)
    return OrderData(**defaults)


def make_revenue(**overrides) -> RevenueData:
    """创建测试用 RevenueData"""
    defaults = dict(
        revenue=10000.0,
        revenue_cny=72500.0,
        chm_cost=3000.0,
        phy_cost=2000.0,
        chm_revenue=5000.0,
        phy_revenue=3000.0,
        chm_cs_cost=1000.0,
        chm_lab_cost=2000.0,
        phy_cs_cost=800.0,
        phy_lab_cost=1500.0,
        cs_cost=1800.0,
        lab_cost=3500.0,
    )
    defaults.update(overrides)
    return RevenueData(**defaults)


def make_flags(**overrides) -> OperationFlags:
    """创建测试用 OperationFlags"""
    defaults = dict(
        va01=True, va02=True, vf01=False, vf03=False,
        save=True, lab_cost=False, plan_cost=False,
        cs=True, chm=True, phy=True,
        every=False, contact=True,
    )
    defaults.update(overrides)
    return OperationFlags(**defaults)


def make_hour(**overrides) -> HourData:
    """创建测试用 HourData"""
    defaults = dict(
        staff_id='EMP001',
        week='15',
        allocated_day='2026.04.07',
        order_no='ORD-001',
        item='10',
        material_code='T75-405-00',
        allocated_hours=8.0,
        office_time=8.0,
    )
    defaults.update(overrides)
    return HourData(**defaults)


def create_sap_instance(mock_session=None, config=None, flags=None) -> Sap:
    """创建 Sap 实例，跳过 __init__ 的 COM 连接"""
    instance = object.__new__(Sap)
    instance.config = config or make_config()
    instance.flags = flags or make_flags()
    instance.result = SapResult()
    instance.today = '2026.04.09'
    instance.session = mock_session or MagicMock()
    instance.connection = MagicMock()
    instance.application = MagicMock()
    instance.SapGuiAuto = MagicMock()
    return instance


def setup_session_mock(session, text_returns=None):
    """配置 session.findById mock

    Args:
        session: MagicMock session
        text_returns: dict {element_id: text_value} — 指定 .text 读取返回值
                      对于需要读取 .text 的元素, 使用 PropertyMock 确保写入不覆盖读取值

    Returns:
        dict: element_id -> mock_element 的缓存，可用于后续断言
    """
    text_returns = text_returns or {}
    _cache = {}

    def _find_by_id(element_id):
        if element_id not in _cache:
            el = MagicMock()
            if element_id in text_returns:
                type(el).text = PropertyMock(return_value=text_returns[element_id])
            _cache[element_id] = el
        return _cache[element_id]

    session.findById = MagicMock(side_effect=_find_by_id)
    return _cache
