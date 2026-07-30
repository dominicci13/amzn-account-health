"""Row ordering and fallbacks for the Premium Shipping block (metrics rows 17-23)."""
from __future__ import annotations

import pytest

# Titles exactly as the PSO widget renders them, verified live 2026-07-30.
WIDGET = {
    "On-Time Delivery Rate": "98.54%",
    "Valid Tracking Rate": "99.67%",
    "Cancellation Rate": "0.06%",
}


def test_maps_widget_titles_to_workbook_row_order(premium_shipping_values, count_unavailable):
    assert premium_shipping_values("Eligible", WIDGET) == [
        "Eligible",
        "98.54%",
        count_unavailable,
        "0.06%",
        count_unavailable,
        "99.67%",
        count_unavailable,
    ]


def test_widget_order_differs_from_workbook_order(premium_shipping_values):
    """The card lists cancellation last; the workbook wants it before valid tracking."""
    values = premium_shipping_values("Eligible", WIDGET)
    assert values[3] == "0.06%"
    assert values[5] == "99.67%"


def test_count_cells_are_always_unavailable(premium_shipping_values, count_unavailable):
    values = premium_shipping_values("Eligible", WIDGET)
    assert [values[2], values[4], values[6]] == [count_unavailable] * 3


@pytest.mark.parametrize("missing", sorted(WIDGET))
def test_metric_the_widget_did_not_render_falls_back(
    premium_shipping_values, count_unavailable, missing
):
    partial = {k: v for k, v in WIDGET.items() if k != missing}
    values = premium_shipping_values("Eligible", partial)
    assert values.count(count_unavailable) == 4
    assert values[0] == "Eligible"


def test_empty_status_falls_back(premium_shipping_values, count_unavailable):
    assert premium_shipping_values("", WIDGET)[0] == count_unavailable


def test_not_eligible_status_is_preserved(premium_shipping_values):
    assert premium_shipping_values("Not Eligible", WIDGET)[0] == "Not Eligible"


def test_returns_exactly_seven_values(premium_shipping_values):
    assert len(premium_shipping_values("Eligible", {})) == 7
