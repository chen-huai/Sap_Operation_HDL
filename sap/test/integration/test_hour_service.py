"""工时服务集成冒烟。"""

import pytest


pytestmark = pytest.mark.integration


def test_login_hour_smoke(hour_service_live, hour):
    result = hour_service_live.login(hour)
    assert result.success
