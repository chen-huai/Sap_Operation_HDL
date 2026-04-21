"""发票服务集成冒烟。"""

import pytest


pytestmark = pytest.mark.integration


def test_create_proforma_smoke(invoice_service_live):
    result = invoice_service_live.create_proforma()
    assert result.success
