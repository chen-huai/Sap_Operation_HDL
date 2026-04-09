"""
SAP 集成测试 fixtures

自动行为:
    - SAP 未启动 → 测试 FAIL 并显示连接错误
    - test_config.py 参数未填写 → 对应测试自动 SKIP
    - ALLOW_SAVE = False → 保存类测试自动 SKIP
"""

import pytest
from sap.function import Sap
from sap.test.integration.test_config import (
    SAP_CONFIG, FLAGS, ORDER, REVENUE, EXISTING_ORDER_NO, HOUR, ALLOW_SAVE,
)


@pytest.fixture(scope='session')
def sap_live():
    """真实 SAP 连接（整个测试会话共享）"""
    sap = Sap(config=SAP_CONFIG, flags=FLAGS)
    if not sap.result.success:
        pytest.fail(f'SAP 连接失败: {sap.result.message}')
    yield sap
    sap.end_sap()


@pytest.fixture
def order():
    """订单数据 — sap_no 未配置则 skip"""
    if not ORDER.sap_no:
        pytest.skip('test_config.py: ORDER.sap_no 未配置')
    return ORDER


@pytest.fixture
def revenue():
    """营收数据"""
    return REVENUE


@pytest.fixture
def existing_order_no():
    """已有订单号 — 未配置则 skip"""
    if not EXISTING_ORDER_NO:
        pytest.skip('test_config.py: EXISTING_ORDER_NO 未配置')
    return EXISTING_ORDER_NO


@pytest.fixture
def hour():
    """工时数据 — staff_id 未配置则 skip"""
    if not HOUR.staff_id:
        pytest.skip('test_config.py: HOUR.staff_id 未配置')
    return HOUR


@pytest.fixture
def require_save():
    """保存权限门控 — ALLOW_SAVE=False 则 skip"""
    if not ALLOW_SAVE:
        pytest.skip('test_config.py: ALLOW_SAVE = False, 保存操作已禁用')
