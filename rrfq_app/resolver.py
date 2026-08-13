from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
from decimal import Decimal

from .estimating import estimate_quote
from .lcsc import LcscClient
from .models import NormalizedQuote, PriceDecision, ResolutionResult, RrfqItem, SourceStatus, SupplierQuote
from .pricing import PriceNormalizationError, normalize_price_to_usd
from .supplier_sources import (
    LcscApiSource,
    SearchContext,
    SourceSearchOutcome,
    SourceStatusAccumulator,
    SupplierSource,
)


def resolve_items(
    items: list[RrfqItem],
    client: LcscClient,
    *,
    use_estimates: bool = True,
) -> list[PriceDecision]:
    result = resolve_items_from_sources(items, [LcscApiSource(client)], SearchContext(), use_estimates=use_estimates)
    return result.decisions


def resolve_items_from_sources(
    items: list[RrfqItem],
    sources: list[SupplierSource],
    context: SearchContext,
    *,
    use_estimates: bool = True,
) -> ResolutionResult:
    decisions: list[PriceDecision] = []
    accumulators = {source.key: SourceStatusAccumulator(source.key, source.label) for source in sources}
    max_workers = min(len(sources), 6) or 1

    for item in items:
        direct_offers: list[NormalizedQuote] = []
        related_quotes: list[SupplierQuote] = []
        source_notes: list[str] = []

        for source, outcome in _search_sources(sources, item, context, max_workers=max_workers):
            accumulators[source.key].add(outcome)
            if outcome.status not in {"ready", "no_data"} and outcome.reason not in source_notes:
                source_notes.append(f"{outcome.label}: {outcome.reason}")

            for quote in outcome.quotes:
                normalized = _normalize_or_none(quote, context)
                if not normalized:
                    source_notes.append(f"{quote.supplier}: cannot normalize {quote.currency} price.")
                    continue
                if _is_direct_offer(quote):
                    direct_offers.append(normalized)
                else:
                    related_quotes.append(quote)

        if direct_offers:
            direct_offers.sort(key=lambda offer: offer.unit_price_usd)
            best = direct_offers[0]
            decisions.append(
                PriceDecision(
                    item=item,
                    quote=best,
                    status="found",
                    reason=best.quote.reason or f"Best direct offer from {best.quote.supplier}.",
                    alternatives_count=max(len(direct_offers) - 1, 0),
                    offers=tuple(direct_offers),
                )
            )
            continue

        if use_estimates:
            estimate = estimate_quote(
                item,
                related_quotes,
                api_configured=_has_configured_api_source(sources),
                api_error="; ".join(source_notes[:3]) or None,
                source_labels=[source.label for source in sources],
            )
            if estimate:
                normalized_estimate = normalize_price_to_usd(estimate, usd_rub=context.usd_rub, usd_cny=context.usd_cny)
                decisions.append(
                    PriceDecision(
                        item=item,
                        quote=normalized_estimate,
                        status="estimated",
                        reason=estimate.reason or "Estimated price.",
                        alternatives_count=len(related_quotes),
                        offers=(normalized_estimate,),
                    )
                )
                continue

        decisions.append(
            PriceDecision(
                item=item,
                quote=None,
                status="not_enough_data",
                reason="; ".join(source_notes[:3]) or "No direct offer and not enough data for estimate.",
                alternatives_count=len(related_quotes),
            )
        )

    statuses = [accumulator.status() for accumulator in accumulators.values()]
    if not statuses:
        statuses = [SourceStatus("none", "Нет источников", "disabled", "Ни один источник не выбран.")]
    return ResolutionResult(decisions=decisions, source_statuses=statuses)


def quotes_from_decisions(decisions: list[PriceDecision]) -> list[SupplierQuote]:
    return [decision.quote.quote for decision in decisions if decision.quote is not None]


def _safe_search(source: SupplierSource, item: RrfqItem, context: SearchContext) -> SourceSearchOutcome:
    try:
        return source.search(item, context)
    except Exception as exc:
        return SourceSearchOutcome(source.key, source.label, "error", f"{type(exc).__name__}: {exc}")


def _search_sources(
    sources: list[SupplierSource],
    item: RrfqItem,
    context: SearchContext,
    *,
    max_workers: int,
) -> list[tuple[SupplierSource, SourceSearchOutcome]]:
    if len(sources) <= 1:
        return [(source, _safe_search(source, item, context)) for source in sources]

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [(source, executor.submit(_safe_search, source, item, context)) for source in sources]
        return [(source, future.result()) for source, future in futures]


def _has_configured_api_source(sources: list[SupplierSource]) -> bool:
    for source in sources:
        client = getattr(source, "client", None)
        if source.key.endswith("_api") and getattr(client, "is_configured", False):
            return True
    return False


def _normalize_or_none(quote: SupplierQuote, context: SearchContext) -> NormalizedQuote | None:
    try:
        return normalize_price_to_usd(quote, usd_rub=context.usd_rub, usd_cny=context.usd_cny)
    except PriceNormalizationError:
        return None


def _is_direct_offer(quote: SupplierQuote) -> bool:
    if quote.is_estimated:
        return False
    confidence = quote.confidence or Decimal("0.70")
    return confidence >= Decimal("0.70")
