"""发票服务测试。"""

from unittest.mock import MagicMock

from sap.test.helpers import create_invoice_service, create_raw_session


class TestInvoiceService:
    def test_create_proforma_success(self):
        service = create_invoice_service(create_raw_session())
        result = service.create_proforma()
        assert result.success

    def test_display_proforma_success(self):
        raw = create_raw_session({"wnd[0]/usr/ctxtVBRK-VBELN": "90000001"})
        service = create_invoice_service(raw)

        result = service.display_proforma()

        assert result.success
        assert result.proforma_no == "90000001"

    def test_create_proforma_failure(self):
        raw = MagicMock()
        raw.findById.side_effect = Exception("VF01 failed")
        service = create_invoice_service(raw)

        result = service.create_proforma()

        assert not result.success
        assert result.step == "vf01"
