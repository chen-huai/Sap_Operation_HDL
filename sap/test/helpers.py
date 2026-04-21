"""sap/test 的公共工厂和 mock 工具。"""

from __future__ import annotations

import sys
from unittest.mock import MagicMock

# 允许在没有真实 pywin32 的环境里导入 sap 模块。
if "win32com" not in sys.modules:
    sys.modules["win32com"] = MagicMock()
    sys.modules["win32com.client"] = MagicMock()

from sap import (
    CostOptions,
    HourData,
    HourService,
    InvoiceService,
    OrderData,
    OrderService,
    PartnerOptions,
    RevenueData,
    SapConfig,
    SapSession,
)


def make_config(**overrides) -> SapConfig:
    defaults = dict(
        order_type="ZOR",
        sales_organization="3002",
        distribution_channels="10",
        sales_office="1000",
        cost_center="1100",
        sub_cost_center_cs="1101",
        sub_cost_center_chm="1102",
        sub_cost_center_phy="1103",
        cs_code="CS001",
        sales_code="SA001",
        data_ae1=["100001"],
        data_az2=["200001"],
    )
    defaults.update(overrides)
    return SapConfig(**defaults)


def make_order(**overrides) -> OrderData:
    defaults = dict(
        sap_no="123456",
        project_no="PRJ-001",
        material_code="T75-405-00",
        currency_type="CNY",
        exchange_rate=1.0,
        cost=5000.0,
        short_text="Test Short Text",
        amount_vat=10000.0,
        long_text="Test Long Text",
        global_partner_code="GP001",
        sales_name="Zhang San",
        ecd="2026.12.31",
    )
    defaults.update(overrides)
    return OrderData(**defaults)


def make_revenue(**overrides) -> RevenueData:
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


def make_hour(**overrides) -> HourData:
    defaults = dict(
        staff_id="EMP001",
        week="15",
        allocated_day="2026.04.07",
        order_no="ORD-001",
        item="10",
        material_code="T75-405-00",
        allocated_hours=8.0,
        office_time=8.0,
    )
    defaults.update(overrides)
    return HourData(**defaults)


def make_partner_options(**overrides) -> PartnerOptions:
    defaults = dict(add_contact=True, add_sales_partner=True)
    defaults.update(overrides)
    return PartnerOptions(**defaults)


def make_cost_options(**overrides) -> CostOptions:
    defaults = dict(include_cs=True, include_chm=True, include_phy=True)
    defaults.update(overrides)
    return CostOptions(**defaults)


def create_raw_session(text_returns: dict[str, str] | None = None):
    """创建原始 COM session mock，并缓存各元素。"""
    text_returns = text_returns or {}
    raw = MagicMock()
    cache: dict[str, MagicMock] = {}

    def _find_by_id(element_id: str):
        if element_id not in cache:
            element = MagicMock()
            element.text = text_returns.get(element_id, "")
            cache[element_id] = element
        return cache[element_id]

    raw.findById = MagicMock(side_effect=_find_by_id)
    raw._cache = cache
    return raw


def create_sap_session(raw_session=None) -> SapSession:
    """基于 mock raw session 创建 SapSession。"""
    raw_session = raw_session or create_raw_session()
    return SapSession(MagicMock(), MagicMock(), MagicMock(), raw_session)


def create_order_service(raw_session=None, config: SapConfig | None = None) -> OrderService:
    return OrderService(create_sap_session(raw_session), config or make_config())


def create_invoice_service(raw_session=None) -> InvoiceService:
    return InvoiceService(create_sap_session(raw_session))


def create_hour_service(raw_session=None) -> HourService:
    return HourService(create_sap_session(raw_session))
