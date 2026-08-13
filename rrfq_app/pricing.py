from __future__ import annotations

from decimal import Decimal, InvalidOperation

from .models import NormalizedQuote, RrfqItem, SupplierQuote


class PriceNormalizationError(ValueError):
    pass


def normalize_price_to_usd(
    quote: SupplierQuote,
    *,
    usd_rub: Decimal | None = None,
    usd_cny: Decimal | None = None,
) -> NormalizedQuote:
    currency = quote.currency.upper().strip()
    price = quote.unit_price

    if price <= 0:
        raise PriceNormalizationError("Unit price must be positive.")

    if currency == "USD":
        return NormalizedQuote(quote=quote, unit_price_usd=price)

    if currency in {"RUB", "RUR"}:
        if not usd_rub or usd_rub <= 0:
            raise PriceNormalizationError("USD/RUB rate is required for RUB quotes.")
        return NormalizedQuote(quote=quote, unit_price_usd=price / usd_rub / Decimal("1.22"))

    if currency == "CNY":
        if not usd_cny or usd_cny <= 0:
            raise PriceNormalizationError("USD/CNY rate is required for CNY quotes.")
        return NormalizedQuote(quote=quote, unit_price_usd=price / usd_cny)

    raise PriceNormalizationError(f"Unsupported currency: {quote.currency}")


def choose_best_quote(
    item: RrfqItem,
    quotes: list[SupplierQuote],
    *,
    usd_rub: Decimal | None = None,
    usd_cny: Decimal | None = None,
) -> NormalizedQuote | None:
    candidates: list[NormalizedQuote] = []
    item_value = normalize_part_key(item.value)

    for quote in quotes:
        if quote.row is not None and quote.row != item.row:
            continue
        if quote.row is None and normalize_part_key(quote.value) != item_value:
            continue
        candidates.append(normalize_price_to_usd(quote, usd_rub=usd_rub, usd_cny=usd_cny))

    if not candidates:
        return None

    return min(candidates, key=lambda candidate: candidate.unit_price_usd)


def decimal_from_cell(value: object) -> Decimal | None:
    if value in (None, ""):
        return None
    try:
        return Decimal(str(value))
    except (InvalidOperation, ValueError):
        return None


def normalize_part_key(value: object) -> str:
    return str(value or "").strip().upper().replace(" ", "")

