"""Legacy SAP adapter backed by the new `sap` service interfaces."""

from __future__ import annotations

import re
import time
from typing import Any

from sap import (
    CostOptions,
    HourData,
    HourService,
    InvoiceService,
    OrderData,
    OrderService,
    PartnerOptions,
    RevenueData,
    SapConfig,
    SapSession,
    OrderItemData,
)


class Sap:
    """Compatibility wrapper for the old GUI code."""

    def __init__(self):
        self._sap_session: SapSession | None = None
        self.session = None
        self._invoice_service: InvoiceService | None = None
        self._hour_service: HourService | None = None
        self.current_order_no = ""
        self.res = {"flag": 1, "msg": ""}

        try:
            self._sap_session = SapSession.connect()
            self.session = self._sap_session.raw
        except Exception as exc:
            self.res = {"flag": 0, "msg": str(exc)}

    def _fail(self, message: str, **extra) -> dict[str, Any]:
        legacy = {
            "flag": 0,
            "msg": message,
            "orderNo": "",
            "Proforma No.": "",
            "sapAmountVat": "",
        }
        legacy.update(extra)
        self.res = {"flag": 0, "msg": message}
        return legacy

    def _legacy_result(self, result, **extra) -> dict[str, Any]:
        legacy = {
            "flag": 1 if result.success else 0,
            "msg": result.message,
            "orderNo": getattr(result, "order_no", "") or "",
            "Proforma No.": getattr(result, "proforma_no", "") or "",
            "sapAmountVat": getattr(result, "sap_amount_vat", "") or "",
        }
        legacy.update(extra)
        self.res = {"flag": legacy["flag"], "msg": legacy["msg"]}
        return legacy

    @staticmethod
    def _as_float(value: Any, default: float = 0.0) -> float:
        try:
            if value in ("", None):
                return default
            return float(value)
        except (TypeError, ValueError):
            return default

    @staticmethod
    def _as_str(value: Any) -> str:
        if value is None:
            return ""
        return str(value).strip()

    @staticmethod
    def _as_list(value: Any) -> list[str]:
        if isinstance(value, list):
            return [str(item).strip() for item in value if str(item).strip()]
        if value in ("", None):
            return []
        return [item.strip() for item in str(value).split(";") if item.strip()]

    def _empty_config(self) -> SapConfig:
        return SapConfig(
            order_type="",
            sales_organization="",
            distribution_channels="",
            sales_office="",
            cost_center="",
            sub_cost_center_cs="",
            sub_cost_center_chm="",
            sub_cost_center_phy="",
            cs_code="",
            sales_code="",
            data_ae1=[],
            data_az2=[],
        )

    def _build_config(self, gui_data: dict[str, Any] | None = None) -> SapConfig:
        gui_data = gui_data or {}
        return SapConfig(
            order_type=self._as_str(gui_data.get("orderType")),
            sales_organization=self._as_str(gui_data.get("salesOrganization")),
            distribution_channels=self._as_str(gui_data.get("distributionChannels")),
            sales_office=self._as_str(gui_data.get("salesOffice")),
            cost_center=self._as_str(gui_data.get("salesGroup")),
            sub_cost_center_cs=self._as_str(gui_data.get("csCostCenter")),
            sub_cost_center_chm=self._as_str(gui_data.get("chmCostCenter")),
            sub_cost_center_phy=self._as_str(gui_data.get("phyCostCenter")),
            cs_code=self._as_str(gui_data.get("csCode")),
            sales_code=self._as_str(gui_data.get("salesCode")),
            data_ae1=self._as_list(gui_data.get("dataAE1")),
            data_az2=self._as_list(gui_data.get("dataAZ2")),
        )

    def _build_order(self, gui_data: dict[str, Any]) -> OrderData:
        items = self._build_order_items(gui_data)
        return OrderData(
            sap_no=self._as_str(gui_data.get("sapNo")),
            project_no=self._as_str(gui_data.get("projectNo")),
            material_code=self._as_str(gui_data.get("materialCode")),
            currency_type=self._as_str(gui_data.get("currencyType")),
            exchange_rate=self._as_float(gui_data.get("exchangeRate"), 1.0),
            cost=self._as_float(gui_data.get("cost")),
            short_text=self._as_str(gui_data.get("shortText")),
            amount_vat=self._as_float(gui_data.get("amountVat")),
            long_text=self._as_str(gui_data.get("longText")),
            global_partner_code=self._as_str(gui_data.get("globalPartnerCode")),
            sales_name=self._as_str(gui_data.get("salesName")),
            ecd=self._as_str(gui_data.get("ecd")) or time.strftime("%Y.%m.%d"),
            items=items,
        )

    def _build_order_items(self, gui_data: dict[str, Any]) -> list[OrderItemData]:
        raw_items = gui_data.get("items") or []
        items: list[OrderItemData] = []
        if isinstance(raw_items, list):
            for raw_item in raw_items:
                if not isinstance(raw_item, dict):
                    continue
                material_code = self._as_str(raw_item.get("material_code", raw_item.get("materialCode")))
                if not material_code:
                    continue
                items.append(
                    OrderItemData(
                        item=self._as_str(raw_item.get("item")),
                        material_code=material_code,
                        revenue=self._as_float(raw_item.get("revenue", raw_item.get("amount"))),
                        quantity=self._as_str(raw_item.get("quantity")) or "1",
                        unit=self._as_str(raw_item.get("unit")) or "pu",
                    )
                )
        return items

    def _build_revenue(self, gui_data: dict[str, Any], revenue_data: dict[str, Any]) -> RevenueData:
        return RevenueData(
            revenue=self._as_float(revenue_data.get("revenue"), self._as_float(gui_data.get("amount"))),
            revenue_cny=self._as_float(
                revenue_data.get("revenueForCny"),
                self._as_float(gui_data.get("amount")) * self._as_float(gui_data.get("exchangeRate"), 1.0),
            ),
            chm_cost=self._as_float(revenue_data.get("chmCost")),
            phy_cost=self._as_float(revenue_data.get("phyCost")),
            chm_revenue=self._as_float(revenue_data.get("chmRe")),
            phy_revenue=self._as_float(revenue_data.get("phyRe")),
            chm_cs_cost=self._as_float(revenue_data.get("chmCsCostAccounting")),
            chm_lab_cost=self._as_float(revenue_data.get("chmLabCostAccounting")),
            phy_cs_cost=self._as_float(revenue_data.get("phyCsCostAccounting")),
            phy_lab_cost=self._as_float(revenue_data.get("phyLabCostAccounting")),
            cs_cost=self._as_float(revenue_data.get("csCostAccounting")),
            lab_cost=self._as_float(revenue_data.get("labCostAccounting")),
        )

    def _build_hour(self, hour_data: Any) -> HourData:
        getter = hour_data.get if hasattr(hour_data, "get") else lambda key, default="": default
        return HourData(
            staff_id=self._as_str(getter("staff_id", "")),
            week=self._as_str(getter("week", "")),
            allocated_day=self._as_str(getter("allocated_day", "")),
            order_no=self._as_str(getter("order_no", "")),
            item=self._as_str(getter("item", "")),
            material_code=self._as_str(getter("material_code", "")),
            allocated_hours=self._as_float(getter("allocated_hours", 0)),
            office_time=self._as_float(getter("office_time", 0)),
        )

    def _order_service(self, gui_data: dict[str, Any] | None = None) -> OrderService:
        if not self._sap_session:
            raise RuntimeError(self.res.get("msg") or "SAP session unavailable")
        return OrderService(self._sap_session, self._build_config(gui_data))

    def _invoice(self) -> InvoiceService:
        if not self._sap_session:
            raise RuntimeError(self.res.get("msg") or "SAP session unavailable")
        if self._invoice_service is None:
            self._invoice_service = InvoiceService(self._sap_session)
        return self._invoice_service

    def _hours(self) -> HourService:
        if not self._sap_session:
            raise RuntimeError(self.res.get("msg") or "SAP session unavailable")
        if self._hour_service is None:
            self._hour_service = HourService(self._sap_session)
        return self._hour_service

    def _extract_order_no(self) -> str:
        if not self._sap_session:
            return ""
        try:
            order_no = self._as_str(self._sap_session.read_text("wnd[0]/usr/ctxtVBAK-VBELN"))
            if order_no:
                return order_no
        except Exception:
            pass

        try:
            status_text = self._sap_session.read_status()
        except Exception:
            return ""

        match = re.search(r"(\d{6,})", status_text)
        return match.group(1) if match else ""

    def va01_operate(self, gui_data: dict[str, Any], revenue_data: dict[str, Any]) -> dict[str, Any]:
        if not self._sap_session:
            return self._fail(self.res.get("msg", "SAP session unavailable"))

        try:
            service = self._order_service(gui_data)
            order = self._build_order(gui_data)
            revenue = self._build_revenue(gui_data, revenue_data)
            result = service.create_order(
                order,
                revenue,
                partner_options=PartnerOptions(
                    add_contact=bool(gui_data.get("contactCheck", True)),
                    add_sales_partner=bool(order.sales_name),
                ),
            )
            return self._legacy_result(result)
        except Exception as exc:
            return self._fail(f"Order No未创建成功，{exc}")

    def lab_cost(self, gui_data: dict[str, Any], revenue_data: dict[str, Any]) -> dict[str, Any]:
        if not self._sap_session:
            return self._fail(self.res.get("msg", "SAP session unavailable"))

        try:
            service = self._order_service(gui_data)
            result = service.fill_lab_cost(
                self._build_order(gui_data),
                self._build_revenue(gui_data, revenue_data),
            )
            return self._legacy_result(result)
        except Exception as exc:
            return self._fail(f"Data B未填写，{exc}")

    def save_sap(self, info: str) -> dict[str, Any]:
        if not self._sap_session:
            return self._fail(self.res.get("msg", "SAP session unavailable"))

        try:
            result = self._order_service().save(info)
            legacy = self._legacy_result(result)
            if legacy["flag"]:
                self.current_order_no = self._extract_order_no() or self.current_order_no
                legacy["orderNo"] = self.current_order_no
            return legacy
        except Exception as exc:
            return self._fail(f"{info}保存失败，{exc}")

    def va02_operate(self, gui_data: dict[str, Any], revenue_data: dict[str, Any]) -> dict[str, Any]:
        if not self._sap_session:
            return self._fail(self.res.get("msg", "SAP session unavailable"))

        try:
            service = self._order_service(gui_data)
            order = self._build_order(gui_data)
            revenue = self._build_revenue(gui_data, revenue_data)

            order_no = self.current_order_no or self._extract_order_no()
            if not order_no:
                return self._fail("未找到可用于VA02的Order No.")

            open_result = service.open_order(order_no)
            if not open_result.success:
                return self._legacy_result(open_result, orderNo=order_no)

            item_result = service.add_items(order, revenue)
            legacy = self._legacy_result(item_result, orderNo=item_result.order_no or order_no)
            self.current_order_no = legacy["orderNo"]
            if not item_result.success:
                return legacy

            if gui_data.get("planCostCheck"):
                cost_result = service.apply_plan_cost(
                    order,
                    revenue,
                    cost_options=CostOptions(
                        include_cs=bool(gui_data.get("csCheck", True)),
                        include_chm=bool(gui_data.get("chmCheck", True)),
                        include_phy=bool(gui_data.get("phyCheck", True)),
                    ),
                )
                if not cost_result.success:
                    return self._legacy_result(
                        cost_result,
                        orderNo=self.current_order_no,
                        sapAmountVat=legacy["sapAmountVat"],
                    )
                if cost_result.message:
                    legacy["msg"] = cost_result.message
                    self.res = {"flag": 1, "msg": legacy["msg"]}

            return legacy
        except Exception as exc:
            return self._fail(f"Order添加Item失败，{exc}")

    def vf01_operate(self) -> dict[str, Any]:
        if not self._sap_session:
            return self._fail(self.res.get("msg", "SAP session unavailable"))

        try:
            return self._legacy_result(self._invoice().create_proforma())
        except Exception as exc:
            return self._fail(f"形式发票添加失败，{exc}")

    def vf03_operate(self) -> dict[str, Any]:
        if not self._sap_session:
            return self._fail(self.res.get("msg", "SAP session unavailable"))

        try:
            return self._legacy_result(self._invoice().display_proforma())
        except Exception as exc:
            return self._fail(f"形式发票查看失败，{exc}")

    def open_va02(self, order_no: str) -> dict[str, Any]:
        if not self._sap_session:
            return self._fail(self.res.get("msg", "SAP session unavailable"))

        try:
            self.current_order_no = self._as_str(order_no)
            return self._legacy_result(self._order_service().open_order(self.current_order_no), orderNo=self.current_order_no)
        except Exception as exc:
            return self._fail(f"该Order No {order_no} 未开启，{exc}")

    def unlock_or_lock_order(self, flag: str) -> dict[str, Any]:
        if not self._sap_session:
            return self._fail(self.res.get("msg", "SAP session unavailable"))

        try:
            unlocked = flag == "Unlock"
            result = self._order_service().transaction.set_lock_state(unlocked=unlocked)
            return self._legacy_result(result, orderNo=self.current_order_no)
        except Exception as exc:
            return self._fail(f"{flag} 未成功，{exc}", orderNo=self.current_order_no)

    def end_sap(self) -> None:
        if self._sap_session:
            self._sap_session.close()
        self._sap_session = None
        self.session = None
        self._invoice_service = None
        self._hour_service = None

    def login_hour_gui(self, hour_data: Any) -> dict[str, Any]:
        if not self._sap_session:
            return self._fail(self.res.get("msg", "SAP session unavailable"))

        try:
            return self._legacy_result(self._hours().login(self._build_hour(hour_data)))
        except Exception as exc:
            return self._fail(f"登录工时系统失败，{exc}")

    def recording_hours(self, hour_data: Any, row_num: int = 0) -> dict[str, Any]:
        if not self._sap_session:
            return self._fail(self.res.get("msg", "SAP session unavailable"))

        try:
            return self._legacy_result(self._hours().record(self._build_hour(hour_data), row_num=row_num))
        except Exception as exc:
            return self._fail(f"录Hour失败，{exc}")

    def save_hours(self) -> dict[str, Any]:
        if not self._sap_session:
            return self._fail(self.res.get("msg", "SAP session unavailable"))

        try:
            return self._legacy_result(self._hours().save(max_retries=14))
        except Exception as exc:
            return self._fail(f"保存工时失败，{exc}")


__all__ = ["Sap"]
