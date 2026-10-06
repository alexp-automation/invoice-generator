from datetime import date
from decimal import Decimal

import pytest

from invoices import build_invoices

CLIENTS = {
    "c1": {"name": "Bravo LLC", "terms_days": "15", "tax_rate": "7"},
    "c2": {"name": "Alpha Inc", "terms_days": "", "tax_rate": ""},
}


def row(day, cid, qty, price, desc="Work"):
    return {"date": day, "client_id": cid, "description": desc, "qty": qty, "unit_price": price}


def test_groups_by_client_filters_month_and_numbers_alphabetically():
    rows = [
        row("2026-09-03", "c1", "2", "100"),
        row("2026-09-20", "c2", "1.5", "80"),
        row("2026-08-31", "c1", "10", "100"),  # previous month, ignored
        row("2026-09-01", "c1", "1", "33.335"),
    ]
    invs = build_invoices(rows, CLIENTS, "2026-09", 1001, date(2026, 10, 1))
    assert [(i.number, i.client["name"]) for i in invs] == [("INV-1001", "Alpha Inc"),
                                                          ("INV-1002", "Bravo LLC")]
    bravo = invs[1]
    assert [l.day.day for l in bravo.lines] == [1, 3]  # sorted by date
    assert bravo.subtotal == Decimal("233.34")         # 33.335 rounds half up
    assert bravo.tax == Decimal("16.33")
    assert bravo.total == Decimal("249.67")
    assert bravo.due == date(2026, 10, 16)
    assert invs[0].tax == 0 and invs[0].due == date(2026, 10, 31)  # defaults


def test_unknown_client_is_an_error():
    with pytest.raises(ValueError, match="Unknown client_id"):
        build_invoices([row("2026-09-01", "zz", "1", "1")], CLIENTS, "2026-09", 1, date.today())
