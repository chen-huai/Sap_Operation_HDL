"""SAP 连接冒烟测试。"""

import pytest


pytestmark = pytest.mark.integration


def test_sap_connected(session_live):
    assert session_live.raw is not None
