"""订单服务集成冒烟。"""

import pytest


pytestmark = pytest.mark.integration


def test_create_order_smoke(order_service_live, order, revenue):
    result = order_service_live.create_order(order, revenue)
    assert result.success
