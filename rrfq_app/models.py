from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal


@dataclass(frozen=True)
class RrfqColumns:
    item_no: int
    value: int
    pn: int
    mfg_russia: int
    mfg_china: int
    pack_qty: int
    qty_to_buy: int
    unit_price: int
    total_price: int
    lead_time: int
    process: int | None = None
    category: int | None = None
    body: int | None = None
    qty_in_order: int | None = None


@dataclass(frozen=True)
class RrfqItem:
    sheet_name: str
    row: int
    item_no: int | None
    value: str
    requested_mfg: str | None
    current_pn: str | None
    current_offer_mfg: str | None
    qty_to_buy: Decimal | None
    process: str | None = None
    category: str | None = None
    body: str | None = None
    qty_in_order: Decimal | None = None


@dataclass(frozen=True)
class SupplierQuote:
    row: int | None
    value: str
    offer_pn: str | None
    offer_mfg: str | None
    supplier: str
    currency: str
    unit_price: Decimal
    pack_qty: int | None = None
    stock_qty: int | None = None
    confidence: Decimal | None = None
    is_estimated: bool = False
    reason: str | None = None
    source_url: str | None = None


@dataclass(frozen=True)
class NormalizedQuote:
    quote: SupplierQuote
    unit_price_usd: Decimal


@dataclass(frozen=True)
class PriceDecision:
    item: RrfqItem
    quote: NormalizedQuote | None
    status: str
    reason: str
    alternatives_count: int = 0
    offers: tuple[NormalizedQuote, ...] = ()


@dataclass(frozen=True)
class SourceStatus:
    key: str
    label: str
    status: str
    reason: str
    searched_items: int = 0
    quote_count: int = 0


@dataclass(frozen=True)
class ResolutionResult:
    decisions: list[PriceDecision]
    source_statuses: list[SourceStatus]
