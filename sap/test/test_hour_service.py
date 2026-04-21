"""工时服务测试。"""

from unittest.mock import MagicMock

from sap.test.helpers import create_hour_service, create_raw_session, make_hour


class TestHourServiceLogin:
    def test_login_invalid_staff_id(self):
        raw = create_raw_session({"wnd[0]/sbar/pane[0]": "员工不存在"})
        service = create_hour_service(raw)

        result = service.login(make_hour())

        assert not result.success
        assert result.step == "hour_login"

    def test_login_success(self):
        raw = create_raw_session({"wnd[0]/sbar/pane[0]": "Ready"})
        service = create_hour_service(raw)

        result = service.login(make_hour())

        assert result.success


class TestHourServiceRecord:
    def test_record_finds_next_empty_row(self):
        raw = create_raw_session(
            {
                "wnd[0]/usr/tblZIIZRUECKMELD00DYNPRO200/txtZRUCKDS-DATUMK[2,0]": "2026.04.01",
                "wnd[0]/usr/tblZIIZRUECKMELD00DYNPRO200/txtZRUCKDS-DATUMK[2,1]": "",
            }
        )
        service = create_hour_service(raw)

        result = service.record(make_hour())

        assert result.success
        assert raw._cache["wnd[0]/usr/tblZIIZRUECKMELD00DYNPRO200/txtZRUCKDS-DATUMK[2,1]"].text == "2026.04.07"

    def test_record_returns_failure_when_no_empty_row(self):
        raw = create_raw_session(
            {
                "wnd[0]/usr/tblZIIZRUECKMELD00DYNPRO200/txtZRUCKDS-DATUMK[2,0]": "x",
                "wnd[0]/usr/tblZIIZRUECKMELD00DYNPRO200/txtZRUCKDS-DATUMK[2,1]": "x",
            }
        )
        service = create_hour_service(raw)

        result = service.record(make_hour(), row_num=0, max_rows=2)

        assert not result.success
        assert result.step == "hour_record"


class TestHourServiceSave:
    def test_save_retry_until_data_saved(self):
        raw = MagicMock()
        call_count = 0

        def _find_by_id(element_id):
            nonlocal call_count
            call_count += 1
            element = MagicMock()
            if call_count == 2:
                raise Exception("OPTION1 not found")
            if element_id == "wnd[0]/sbar/pane[0]":
                element.text = "Data was saved"
            return element

        raw.findById.side_effect = _find_by_id
        service = create_hour_service(raw)

        result = service.save()

        assert result.success
        assert result.message == "录Hour成功"

    def test_save_exceeds_max_retries(self):
        raw = MagicMock()
        call_count = 0

        def _find_by_id(_):
            nonlocal call_count
            call_count += 1
            if call_count == 2:
                raise Exception("initial")
            raise Exception("retry")

        raw.findById.side_effect = _find_by_id
        service = create_hour_service(raw)

        result = service.save(max_retries=3)

        assert not result.success
        assert "已重试3次" in result.message
