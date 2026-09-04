from __future__ import annotations

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

import sap.test.helpers  # noqa: E402,F401
from sap.models import DataBEntry, OrderData  # noqa: E402
from sap.transactions.order_edit import OrderEditTransaction  # noqa: E402


def _order(sales_group: str = "200") -> OrderData:
    return OrderData(
        sap_no="123456",
        project_no="PRJ-001",
        currency_type="CNY",
        exchange_rate=1.0,
        short_text="Test",
        sales_group=sales_group,
    )


class DataBDiffMessageTest(unittest.TestCase):
    def test_missing_performer_and_cost_rows_are_business_readable(self):
        diff = OrderEditTransaction._data_b_diff(
            zul=[("0048601293", "")],
            kos=[],
            entries=[
                DataBEntry("48601294", "48601294", 100.0, item="2000"),
                DataBEntry("48601293", "48601293", 30.0, item="3000"),
            ],
            order=_order(),
        )

        self.assertEqual(
            diff,
            "Data B 有差异：新增执行部门 48601294；"
            "新增成本行，item 2000，费率中心 48601294，金额 100.00；"
            "新增成本行，item 3000，费率中心 48601293，金额 30.00",
        )
        self.assertNotIn("SAP=[", diff)
        self.assertNotIn("Excel=[", diff)

    def test_same_rate_and_item_amount_difference_is_modified_row(self):
        diff = OrderEditTransaction._data_b_diff(
            zul=[("0048601293", "")],
            kos=[("0048601293", "003000", "20")],
            entries=[DataBEntry("48601293", "48601293", 30.0, item="3000")],
            order=_order(),
        )

        self.assertEqual(
            diff,
            "Data B 有差异：修改成本行，item 3000，费率中心 48601293，金额 20.00→30.00",
        )


if __name__ == "__main__":
    unittest.main()
