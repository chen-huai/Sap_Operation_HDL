"""sap/test 的共享 pytest 配置。"""

from __future__ import annotations

from pathlib import Path

import pytest

from sap.test.helpers import (
    create_hour_service,
    create_invoice_service,
    create_order_service,
    create_raw_session,
    create_sap_session,
)


def pytest_addoption(parser):
    parser.addoption(
        "--run-integration",
        action="store_true",
        default=False,
        help="运行需要真实 SAP GUI 的集成测试",
    )


def pytest_configure(config):
    config.addinivalue_line("markers", "integration: 需要真实 SAP GUI 的集成测试")


def pytest_collection_modifyitems(config, items):
    if config.getoption("--run-integration"):
        return

    skip_integration = pytest.mark.skip(reason="默认跳过集成测试，使用 --run-integration 显式开启")
    for item in items:
        if "integration" in Path(str(item.fspath)).parts:
            item.add_marker(skip_integration)


@pytest.fixture
def raw_session():
    return create_raw_session()


@pytest.fixture
def sap_session(raw_session):
    return create_sap_session(raw_session)


@pytest.fixture
def order_service(raw_session):
    return create_order_service(raw_session)


@pytest.fixture
def invoice_service(raw_session):
    return create_invoice_service(raw_session)


@pytest.fixture
def hour_service(raw_session):
    return create_hour_service(raw_session)
