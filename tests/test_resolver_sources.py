from __future__ import annotations

import unittest
import time
from decimal import Decimal

from rrfq_app.models import RrfqItem, SupplierQuote
from rrfq_app.resolver import resolve_items_from_sources
from rrfq_app.supplier_sources import SearchContext, SourceSearchOutcome


class ResolverSourcesTest(unittest.TestCase):
    def test_selects_lowest_normalized_direct_offer(self):
        item = _item()
        sources = [
            _FakeSource(
                "usd",
                "USD Source",
                [
                    SupplierQuote(
                        row=item.row,
                        value=item.value,
                        offer_pn=item.value,
                        offer_mfg="TI",
                        supplier="USD Source",
                        currency="USD",
                        unit_price=Decimal("2.00"),
                        confidence=Decimal("0.95"),
                    )
                ],
            ),
            _FakeSource(
                "rub",
                "RUB Source",
                [
                    SupplierQuote(
                        row=item.row,
                        value=item.value,
                        offer_pn=item.value,
                        offer_mfg="TI",
                        supplier="RUB Source",
                        currency="RUB",
                        unit_price=Decimal("120"),
                        confidence=Decimal("0.95"),
                    )
                ],
            ),
        ]

        result = resolve_items_from_sources([item], sources, SearchContext(usd_rub=Decimal("100")))

        self.assertEqual(result.decisions[0].status, "found")
        self.assertEqual(result.decisions[0].quote.quote.supplier, "RUB Source")
        self.assertEqual(result.decisions[0].quote.unit_price_usd, Decimal("120") / Decimal("100") / Decimal("1.22"))
        self.assertEqual(result.source_statuses[0].status, "ready")
        self.assertEqual(result.source_statuses[1].quote_count, 1)

    def test_estimates_when_source_needs_manual_action(self):
        item = _item()
        sources = [_BlockedSource()]

        result = resolve_items_from_sources([item], sources, SearchContext())

        self.assertEqual(result.decisions[0].status, "estimated")
        self.assertTrue(result.decisions[0].quote.quote.is_estimated)
        self.assertEqual(result.source_statuses[0].status, "needs_manual_action")
        self.assertNotIn("No LCSC API credentials", result.decisions[0].reason)

    def test_searches_sources_in_parallel_per_item(self):
        item = _item()
        sources = [_SlowNoDataSource("a"), _SlowNoDataSource("b"), _SlowNoDataSource("c")]

        started = time.monotonic()
        result = resolve_items_from_sources([item], sources, SearchContext(), use_estimates=False)
        elapsed = time.monotonic() - started

        self.assertEqual(result.decisions[0].status, "not_enough_data")
        self.assertLess(elapsed, 0.45)


class _FakeSource:
    def __init__(self, key, label, quotes):
        self.key = key
        self.label = label
        self.quotes = tuple(quotes)

    def search(self, item, context):
        return SourceSearchOutcome(self.key, self.label, "ready", "ok", self.quotes)


class _BlockedSource:
    key = "blocked"
    label = "Blocked Source"

    def search(self, item, context):
        return SourceSearchOutcome(self.key, self.label, "needs_manual_action", "captcha")


class _SlowNoDataSource:
    def __init__(self, key):
        self.key = key
        self.label = f"Slow {key}"

    def search(self, item, context):
        time.sleep(0.2)
        return SourceSearchOutcome(self.key, self.label, "no_data", "nothing")


def _item():
    return RrfqItem(
        sheet_name="BOM",
        row=49,
        item_no=1,
        value="TMS320F28069PZ",
        requested_mfg="TI",
        current_pn=None,
        current_offer_mfg=None,
        qty_to_buy=Decimal("150"),
        category="IC",
        body="LQFP-100",
    )


if __name__ == "__main__":
    unittest.main()
