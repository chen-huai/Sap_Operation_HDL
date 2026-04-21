"""真实 SAP GUI 集成测试 fixtures。"""

from __future__ import annotations

import pytest

from sap import HourService, InvoiceService, OrderService, SapSession
from sap.test.integration.test_config import HOUR, ORDER, REVENUE, SAP_CONFIG


@pytest.fixture(scope="session")
def session_live():
    session = SapSession.connect()
    yield session
    session.close()


@pytest.fixture(scope="session")
def order_service_live(session_live):
    return OrderService(session_live, SAP_CONFIG)


@pytest.fixture(scope="session")
def invoice_service_live(session_live):
    return InvoiceService(session_live)


@pytest.fixture(scope="session")
def hour_service_live(session_live):
    return HourService(session_live)


@pytest.fixture
def order():
    if not ORDER.sap_no:
        pytest.skip("test_config.py: ORDER.sap_no 未配置")
    return ORDER


@pytest.fixture
def revenue():
    return REVENUE


@pytest.fixture
def hour():
    if not HOUR.staff_id:
        pytest.skip("test_config.py: HOUR.staff_id 未配置")
    return HOUR
