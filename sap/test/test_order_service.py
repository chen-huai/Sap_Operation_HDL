"""订单服务测试。"""

from unittest.mock import MagicMock

from sap.test.helpers import (
    create_order_service,
    create_raw_session,
    make_config,
    make_cost_options,
    make_order,
    make_partner_options,
    make_revenue,
)


PARTNER_TEXTS = {
    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/"
    "subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/"
    "cmbGVS_TC_DATA-REC-PARVW[0,4]": "负责雇员",
    "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4352/"
    "subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/"
    "cmbGVS_TC_DATA-REC-PARVW[0,5]": "送达方",
    "wnd[0]/usr/ctxtVBAK-VBELN": "60001234",
    "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPLV69A:6201/"
    "tblSAPLV69ATCTRL_KONDITIONEN/txtKOMV-KBETR[3,5]": "10,000.00",
    "wnd[0]/sbar/pane[0]": "文档 60001234 已保存",
}


class TestOrderServiceCreate:
    def test_create_order_success(self):
        raw = create_raw_session(PARTNER_TEXTS)
        service = create_order_service(raw)

        result = service.create_order(
            make_order(currency_type="USD", exchange_rate=7.25),
            make_revenue(revenue_cny=50000.0),
            partner_options=make_partner_options(),
        )

        assert result.success
        assert raw._cache["wnd[0]/usr/ctxtVBAK-AUART"].text == "ZOR"
        assert raw._cache[
            "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/"
            "ssubSUBSCREEN_BODY:SAPMV45A:4312/txtZAUFTD-AUFTRAGSWERT"
        ].text == "50000.00"

    def test_create_order_failure(self):
        raw = MagicMock()
        raw.findById.side_effect = Exception("SAP 页面异常")
        service = create_order_service(raw)

        result = service.create_order(make_order(), make_revenue())

        assert not result.success
        assert result.step == "va01"


class TestOrderServiceItemAndCost:
    def test_add_items_non_a2_material(self):
        raw = create_raw_session(PARTNER_TEXTS)
        service = create_order_service(raw)

        result = service.add_items(make_order(material_code="T75-405-00"), make_revenue(revenue=10000.0))

        assert result.success
        assert result.order_no == "60001234"
        assert result.sap_amount_vat == "10000.00"

    def test_apply_plan_cost_below_threshold_no_editor_open(self):
        raw = create_raw_session(PARTNER_TEXTS)
        service = create_order_service(raw)

        result = service.apply_plan_cost(
            make_order(material_code="T75-405-00"),
            make_revenue(revenue_cny=500.0),
            cost_options=make_cost_options(),
        )

        assert result.success
        assert "wnd[1]/usr/btnSPOP-VAROPTION1" not in raw._cache

    def test_fill_lab_cost_success(self):
        raw = create_raw_session(PARTNER_TEXTS)
        service = create_order_service(raw)

        result = service.fill_lab_cost(make_order(material_code="T20-441-00"), make_revenue(phy_cost=888.0))

        assert result.success
        assert raw._cache[
            "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4312/"
            "tblSAPMV45AKOSTENSAETZE/txtTABD-FESTPREIS[5,0]"
        ].text == 888.0


class TestOrderServiceSaveAndLock:
    def test_save_success(self):
        raw = create_raw_session(PARTNER_TEXTS)
        service = create_order_service(raw)

        result = service.save("订单")

        assert result.success

    def test_save_failure_when_status_not_saved(self):
        raw = create_raw_session({"wnd[0]/sbar/pane[0]": "发生错误"})
        service = create_order_service(raw)

        result = service.save("订单")

        assert not result.success
        assert "订单保存失败" in result.message

    def test_open_order_failure(self):
        raw = MagicMock()
        raw.findById.side_effect = Exception("页面打不开")
        service = create_order_service(raw)

        result = service.open_order("60001234")

        assert not result.success
        assert result.step == "open_va02"

    def test_unlock_calls_open_and_sets_success_message(self):
        raw = create_raw_session(PARTNER_TEXTS)
        service = create_order_service(raw, config=make_config())

        result = service.unlock("60001234")

        assert result.success
        assert "Unlock 成功" == result.message
