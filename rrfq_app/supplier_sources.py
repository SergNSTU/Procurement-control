from __future__ import annotations

import html
import json
import re
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from typing import Any, Protocol

from .lcsc import JsonCache, LcscClient, LcscClientError
from .models import RrfqItem, SourceStatus, SupplierQuote
from .quotes_csv import load_quotes_csv_text


class SupplierSource(Protocol):
    key: str
    label: str

    def search(self, item: RrfqItem, context: "SearchContext") -> "SourceSearchOutcome":
        ...


@dataclass(frozen=True)
class SearchContext:
    usd_rub: Decimal | None = None
    usd_cny: Decimal | None = None
    manual_quotes: tuple[SupplierQuote, ...] = ()


@dataclass(frozen=True)
class SourceSearchOutcome:
    key: str
    label: str
    status: str
    reason: str
    quotes: tuple[SupplierQuote, ...] = ()


@dataclass(frozen=True)
class SourceDefinition:
    key: str
    label: str
    mode: str
    default_enabled: bool
    note: str


SOURCE_DEFINITIONS = [
    SourceDefinition("lcsc_web", "LCSC Web", "public_web", True, "Публичный поиск без API-ключа."),
    SourceDefinition("electronshik", "Electronshik", "public_web", True, "Публичный поиск по сайту."),
    SourceDefinition("promelec", "Промэлектроника", "public_web", True, "Публичный поиск по сайту."),
    SourceDefinition("kompel_sds", "Компэл СДС", "cabinet", False, "Нужен ваш доступ к СДС/кабинету."),
    SourceDefinition("mouser_web", "Mouser Web", "public_web", False, "Публичный поиск, API предпочтительнее."),
    SourceDefinition("digikey_web", "DigiKey Web", "public_web", False, "Публичный поиск, API предпочтительнее."),
    SourceDefinition("lcsc_api", "LCSC API", "api", False, "Опционально, если появится API-ключ."),
    SourceDefinition("manual_csv", "CSV/прайс", "file", False, "Ручная выгрузка поставщика."),
]


class SourceStatusAccumulator:
    def __init__(self, key: str, label: str) -> None:
        self.key = key
        self.label = label
        self.statuses: list[str] = []
        self.reasons: list[str] = []
        self.searched_items = 0
        self.quote_count = 0

    def add(self, outcome: SourceSearchOutcome) -> None:
        self.statuses.append(outcome.status)
        if outcome.reason and outcome.reason not in self.reasons:
            self.reasons.append(outcome.reason)
        if outcome.status not in {"disabled", "needs_api", "needs_login", "needs_file"}:
            self.searched_items += 1
        self.quote_count += len(outcome.quotes)

    def status(self) -> SourceStatus:
        status = _aggregate_status(self.statuses, self.quote_count)
        reason = self.reasons[0] if self.reasons else "No activity."
        if self.quote_count:
            reason = f"Получено предложений: {self.quote_count}."
        return SourceStatus(
            key=self.key,
            label=self.label,
            status=status,
            reason=reason,
            searched_items=self.searched_items,
            quote_count=self.quote_count,
        )


class ManualCsvSource:
    key = "manual_csv"
    label = "CSV/прайс"

    def search(self, item: RrfqItem, context: SearchContext) -> SourceSearchOutcome:
        if not context.manual_quotes:
            return SourceSearchOutcome(self.key, self.label, "needs_file", "CSV/прайс не загружен.")

        quotes = [
            quote
            for quote in context.manual_quotes
            if (quote.row is not None and quote.row == item.row)
            or (quote.row is None and _part_key(quote.value) == _part_key(item.value))
        ]
        if not quotes:
            return SourceSearchOutcome(self.key, self.label, "no_data", "В CSV нет предложения по этой строке.")
        normalized = tuple(
            SupplierQuote(
                row=item.row,
                value=item.value,
                offer_pn=quote.offer_pn or item.value,
                offer_mfg=quote.offer_mfg,
                supplier=quote.supplier or self.label,
                currency=quote.currency,
                unit_price=quote.unit_price,
                pack_qty=quote.pack_qty,
                stock_qty=quote.stock_qty,
                confidence=quote.confidence or Decimal("0.95"),
                reason=quote.reason or "Предложение из загруженного CSV/прайса.",
                source_url=quote.source_url,
            )
            for quote in quotes
        )
        return SourceSearchOutcome(self.key, self.label, "ready", "CSV matched.", normalized)


class CabinetOnlySource:
    def __init__(self, key: str, label: str, reason: str) -> None:
        self.key = key
        self.label = label
        self.reason = reason

    def search(self, item: RrfqItem, context: SearchContext) -> SourceSearchOutcome:
        return SourceSearchOutcome(self.key, self.label, "needs_login", self.reason)


class LcscApiSource:
    key = "lcsc_api"
    label = "LCSC API"

    def __init__(self, client: LcscClient) -> None:
        self.client = client

    def search(self, item: RrfqItem, context: SearchContext) -> SourceSearchOutcome:
        if not self.client.is_configured:
            return SourceSearchOutcome(self.key, self.label, "needs_api", "Нет LCSC API key/secret.")
        try:
            products = self.client.search_products(item.value, exact=True, is_available=False)
            quotes = tuple(_lcsc_products_to_quotes(item, products, exact=True))
            if quotes:
                return SourceSearchOutcome(self.key, self.label, "ready", "LCSC API exact search.", quotes)
            fuzzy = self.client.search_products(item.value, exact=False, is_available=False)
            fuzzy_quotes = tuple(
                quote for quote in _lcsc_products_to_quotes(item, fuzzy, exact=False)
                if (quote.confidence or Decimal("0")) >= Decimal("0.82")
            )
            if fuzzy_quotes:
                return SourceSearchOutcome(self.key, self.label, "ready", "LCSC API fuzzy search.", fuzzy_quotes)
            return SourceSearchOutcome(self.key, self.label, "no_data", "LCSC API не нашёл цену.")
        except LcscClientError as exc:
            return SourceSearchOutcome(self.key, self.label, "error", str(exc))


class PublicWebSource:
    def __init__(
        self,
        *,
        key: str,
        label: str,
        url_templates: tuple[str, ...],
        currency: str,
        cache: JsonCache | None,
        min_interval_seconds: float = 0.2,
        timeout_seconds: float = 8,
        query_limit: int = 2,
    ) -> None:
        self.key = key
        self.label = label
        self.url_templates = url_templates
        self.currency = currency
        self.cache = cache
        self.min_interval_seconds = min_interval_seconds
        self.timeout_seconds = timeout_seconds
        self.query_limit = query_limit
        self._last_request_at = 0.0

    def search(self, item: RrfqItem, context: SearchContext) -> SourceSearchOutcome:
        variants = part_variants(item.value)
        attempted = 0
        blocked_reason: str | None = None
        collected: list[SupplierQuote] = []

        for query in variants[: self.query_limit]:
            for template in self.url_templates[:1]:
                url = template.format(query=urllib.parse.quote(query))
                attempted += 1
                page = self._fetch(url)
                if page.status == "blocked":
                    blocked_reason = page.reason
                    break
                if page.status != "ready":
                    continue
                quotes = _quotes_from_html(
                    item,
                    html_text=page.text,
                    source_label=self.label,
                    source_url=url,
                    currency=self.currency,
                    query=query,
                )
                collected.extend(quotes)
                if any((quote.confidence or Decimal("0")) >= Decimal("0.82") for quote in quotes):
                    break
            if collected or blocked_reason:
                break

        if collected:
            return SourceSearchOutcome(
                self.key,
                self.label,
                "ready",
                f"Публичный поиск: проверено запросов {attempted}.",
                tuple(_dedupe_quotes(collected)),
            )
        if blocked_reason:
            return SourceSearchOutcome(self.key, self.label, "needs_manual_action", blocked_reason)
        return SourceSearchOutcome(self.key, self.label, "no_data", f"Публичный поиск не нашёл цену ({attempted} запросов).")

    def _fetch(self, url: str) -> "_FetchedPage":
        cache_key = f"public-web:{self.key}:{url}"
        if self.cache:
            cached = self.cache.get(cache_key)
            if cached is not None:
                return _FetchedPage("ready", cached, "cache")

        elapsed = time.monotonic() - self._last_request_at
        if elapsed < self.min_interval_seconds:
            time.sleep(self.min_interval_seconds - elapsed)
        self._last_request_at = time.monotonic()

        request = urllib.request.Request(
            url,
            headers={
                "Accept": "text/html,application/xhtml+xml",
                "User-Agent": "RRFQ-Pricer/0.1 (+local procurement tool)",
            },
        )
        try:
            with urllib.request.urlopen(request, timeout=self.timeout_seconds) as response:
                body = response.read()
                charset = response.headers.get_content_charset() or "utf-8"
                text = body.decode(charset, errors="replace")
        except urllib.error.HTTPError as exc:
            if exc.code in {401, 403, 429}:
                return _FetchedPage("blocked", "", f"{self.label}: сайт ограничил доступ HTTP {exc.code}.")
            return _FetchedPage("error", "", f"{self.label}: HTTP {exc.code}.")
        except urllib.error.URLError as exc:
            return _FetchedPage("error", "", f"{self.label}: ошибка соединения {exc.reason}.")

        if _looks_blocked(text):
            return _FetchedPage("blocked", "", f"{self.label}: требуется ручное действие, логин или CAPTCHA.")
        if self.cache:
            self.cache.set(cache_key, text)
        return _FetchedPage("ready", text, "ok")


@dataclass(frozen=True)
class _FetchedPage:
    status: str
    text: str
    reason: str


def build_sources(
    enabled_keys: set[str],
    *,
    cache: JsonCache,
    lcsc_client: LcscClient,
) -> list[SupplierSource]:
    sources: list[SupplierSource] = []
    for definition in SOURCE_DEFINITIONS:
        if definition.key not in enabled_keys:
            continue
        source = _source_for_key(definition.key, cache=cache, lcsc_client=lcsc_client)
        if source is not None:
            sources.append(source)
    return sources


def default_source_keys() -> set[str]:
    return {definition.key for definition in SOURCE_DEFINITIONS if definition.default_enabled}


def source_definitions_json() -> list[dict[str, Any]]:
    return [definition.__dict__ for definition in SOURCE_DEFINITIONS]


def parse_manual_quotes(text: str | None) -> tuple[SupplierQuote, ...]:
    if not text or not text.strip():
        return ()
    return tuple(load_quotes_csv_text(text))


def _source_for_key(key: str, *, cache: JsonCache, lcsc_client: LcscClient) -> SupplierSource | None:
    if key == "lcsc_web":
        return PublicWebSource(
            key="lcsc_web",
            label="LCSC Web",
            currency="USD",
            cache=cache,
            url_templates=(
                "https://www.lcsc.com/search?q={query}",
                "https://www.lcsc.com/products?keywords={query}",
            ),
        )
    if key == "electronshik":
        return PublicWebSource(
            key="electronshik",
            label="Electronshik",
            currency="RUB",
            cache=cache,
            url_templates=(
                "https://www.electronshik.ru/search?search={query}",
                "https://www.electronshik.ru/?search={query}",
            ),
        )
    if key == "promelec":
        return PublicWebSource(
            key="promelec",
            label="Промэлектроника",
            currency="RUB",
            cache=cache,
            url_templates=(
                "https://www.promelec.ru/search/?q={query}",
                "https://www.promelec.ru/?q={query}",
            ),
        )
    if key == "mouser_web":
        return PublicWebSource(
            key="mouser_web",
            label="Mouser Web",
            currency="USD",
            cache=cache,
            url_templates=(
                "https://www.mouser.com/s/?keyword={query}",
                "https://www.mouser.com/c/?q={query}",
            ),
        )
    if key == "digikey_web":
        return PublicWebSource(
            key="digikey_web",
            label="DigiKey Web",
            currency="USD",
            cache=cache,
            url_templates=("https://www.digikey.com/en/products/result?keywords={query}",),
        )
    if key == "kompel_sds":
        return CabinetOnlySource("kompel_sds", "Компэл СДС", "Нужен ваш логин/доступ к СДС или разрешённая выгрузка.")
    if key == "lcsc_api":
        return LcscApiSource(lcsc_client)
    if key == "manual_csv":
        return ManualCsvSource()
    return None


def part_variants(value: str) -> list[str]:
    raw = value.strip()
    if not raw:
        return []
    base = re.sub(r"\[[^\]]+\]", "", raw).strip()
    variants = [raw, base]
    suffixes = ["-TR", "-T", "-R", "R", "T", "TR", "TRPBF", "PBF"]
    for suffix in suffixes:
        if not base.upper().endswith(suffix):
            variants.append(f"{base}{suffix}")
    simplified = re.sub(r"[^A-Za-z0-9]", "", base)
    if simplified:
        variants.append(simplified)
    return _dedupe_strings([variant for variant in variants if variant])


def _quotes_from_html(
    item: RrfqItem,
    *,
    html_text: str,
    source_label: str,
    source_url: str,
    currency: str,
    query: str,
) -> list[SupplierQuote]:
    text = _visible_text(html_text)
    if not _has_part_match(text, item.value, query):
        return []

    prices = _extract_prices(text, currency)
    if not prices:
        return []

    price = _choose_price(prices)
    if price is None:
        return []

    confidence = _confidence_from_text(text, item, query)
    if confidence < Decimal("0.62"):
        return []

    return [
        SupplierQuote(
            row=item.row,
            value=item.value,
            offer_pn=_best_offer_pn(text, item.value, query),
            offer_mfg=_best_mfg(text, item.requested_mfg),
            supplier=source_label,
            currency=currency,
            unit_price=price,
            stock_qty=_extract_stock(text),
            confidence=confidence,
            reason=f"{source_label}: цена найдена на публичной странице, confidence={confidence}.",
            source_url=source_url,
        )
    ]


def _lcsc_products_to_quotes(item: RrfqItem, products: list[dict[str, Any]], *, exact: bool) -> list[SupplierQuote]:
    quotes: list[SupplierQuote] = []
    qty = int(item.qty_to_buy or item.qty_in_order or 1)
    for product in products:
        price = _price_for_quantity(product, qty)
        if price is None:
            continue
        offer_pn = _first_text(product, "productModel", "product_model", "mpn", "mfrPartNumber", "mfr_part_number")
        lcsc_part = _first_text(product, "productCode", "productNo", "productNumber", "lcscPartNumber", "lcsc_part_number")
        offer_mfg = _first_text(product, "brandName", "brand", "manufacturer", "manufacturerName", "mfr")
        confidence = _match_score(item, offer_pn, lcsc_part, offer_mfg)
        if exact and confidence < Decimal("0.70"):
            continue
        quotes.append(
            SupplierQuote(
                row=item.row,
                value=item.value,
                offer_pn=offer_pn or lcsc_part or item.value,
                offer_mfg=offer_mfg,
                supplier="LCSC API",
                currency="USD",
                unit_price=price,
                pack_qty=_first_int(product, "minPacketUnit", "minOrderQty", "minimumOrderQuantity", "packQty"),
                stock_qty=_first_int(product, "stockNumber", "stock", "quantity", "inventoryQuantity", "availableQuantity"),
                confidence=confidence,
                reason="LCSC API priced catalog match.",
            )
        )
    return quotes


def _price_for_quantity(product: dict[str, Any], qty: int) -> Decimal | None:
    tiers: list[tuple[int, Decimal]] = []

    def walk(node: Any) -> None:
        if isinstance(node, list):
            for item in node:
                walk(item)
            return
        if not isinstance(node, dict):
            return
        min_qty = _first_int(node, "startQty", "beginQty", "minQty", "qty", "quantity", "from")
        price = _first_decimal(node, "price", "unitPrice", "unit_price", "usdPrice", "discountPrice")
        if min_qty is not None and price is not None and price > 0:
            tiers.append((min_qty, price))
        for value in node.values():
            if isinstance(value, (dict, list)):
                walk(value)

    walk(product)
    if tiers:
        tiers.sort(key=lambda item: item[0])
        selected = tiers[0][1]
        for min_qty, price in tiers:
            if qty >= min_qty:
                selected = price
        return selected
    return _first_decimal(product, "unitPrice", "unit_price", "price", "usdPrice", "productPrice")


def _first_text(node: dict[str, Any], *keys: str) -> str | None:
    lowered = {str(key).lower(): value for key, value in node.items()}
    for key in keys:
        value = lowered.get(key.lower())
        if value not in (None, ""):
            return str(value).strip()
    return None


def _first_int(node: dict[str, Any], *keys: str) -> int | None:
    for key in keys:
        value = _value_by_key(node, key)
        parsed = _parse_int(value)
        if parsed is not None:
            return parsed
    return None


def _first_decimal(node: dict[str, Any], *keys: str) -> Decimal | None:
    for key in keys:
        value = _value_by_key(node, key)
        parsed = _parse_decimal(value)
        if parsed is not None:
            return parsed
    return None


def _value_by_key(node: dict[str, Any], key: str) -> Any:
    lowered_key = key.lower()
    for candidate_key, value in node.items():
        if str(candidate_key).lower() == lowered_key:
            return value
    return None


def _extract_prices(text: str, currency: str) -> list[Decimal]:
    if currency == "USD":
        pattern = r"\$\s*([0-9][0-9\s,.]*)"
    elif currency == "RUB":
        pattern = r"([0-9][0-9\s,.]*)\s*(?:₽|руб\.?|р\.)"
    else:
        pattern = r"([0-9][0-9\s,.]*)"

    prices: list[Decimal] = []
    for match in re.finditer(pattern, text, flags=re.IGNORECASE):
        parsed = _parse_decimal(match.group(1))
        if parsed is not None and Decimal("0.0001") <= parsed <= Decimal("1000000"):
            prices.append(parsed)
    return prices[:40]


def _choose_price(prices: list[Decimal]) -> Decimal | None:
    valid = [price for price in prices if price > 0]
    if not valid:
        return None
    return min(valid)


def _extract_stock(text: str) -> int | None:
    patterns = [
        r"(?:In-Stock|In Stock|Наличие)[^\d]{0,20}([0-9][0-9\s,]*)",
        r"([0-9][0-9\s,]*)\s*(?:шт|pcs)\b",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            return _parse_int(match.group(1))
    return None


def _has_part_match(text: str, value: str, query: str) -> bool:
    text_key = _part_key(text)
    return _part_key(value) in text_key or _part_key(query) in text_key


def _confidence_from_text(text: str, item: RrfqItem, query: str) -> Decimal:
    item_key = _part_key(item.value)
    query_key = _part_key(query)
    text_key = _part_key(text)
    score = Decimal("0.62")
    if item_key and item_key in text_key:
        score = Decimal("0.86")
    elif query_key and query_key in text_key:
        score = Decimal("0.74")
    if item.requested_mfg and _brand_key(item.requested_mfg) in _brand_key(text):
        score += Decimal("0.06")
    if item.body and _part_key(item.body) in text_key:
        score += Decimal("0.03")
    return min(score, Decimal("0.95"))


def _best_offer_pn(text: str, value: str, query: str) -> str:
    for candidate in [value, query]:
        if _part_key(candidate) in _part_key(text):
            return candidate
    return value


def _best_mfg(text: str, requested_mfg: str | None) -> str | None:
    if requested_mfg and _brand_key(requested_mfg) in _brand_key(text):
        return requested_mfg
    return requested_mfg


def _match_score(item: RrfqItem, offer_pn: str | None, lcsc_part: str | None, offer_mfg: str | None) -> Decimal:
    item_key = _part_key(item.value)
    offer_key = _part_key(offer_pn)
    lcsc_key = _part_key(lcsc_part)
    score = Decimal("0.40")
    if offer_key == item_key:
        score = Decimal("0.92")
    elif lcsc_key == item_key:
        score = Decimal("0.84")
    elif offer_key and (offer_key.startswith(item_key) or item_key.startswith(offer_key)):
        score = Decimal("0.82")
    if offer_mfg and item.requested_mfg and _brand_key(item.requested_mfg) in _brand_key(offer_mfg):
        score += Decimal("0.05")
    return min(score, Decimal("0.98"))


def _visible_text(html_text: str) -> str:
    without_scripts = re.sub(r"<(script|style)\b.*?</\1>", " ", html_text, flags=re.IGNORECASE | re.DOTALL)
    without_tags = re.sub(r"<[^>]+>", " ", without_scripts)
    return html.unescape(re.sub(r"\s+", " ", without_tags))


def _looks_blocked(text: str) -> bool:
    lowered = text.lower()
    markers = ["captcha", "cloudflare", "access denied", "too many requests", "robot check", "проверка безопасности"]
    return any(marker in lowered for marker in markers)


def _dedupe_quotes(quotes: list[SupplierQuote]) -> list[SupplierQuote]:
    deduped: list[SupplierQuote] = []
    seen: set[str] = set()
    for quote in quotes:
        key = json.dumps(
            [quote.supplier, quote.offer_pn, quote.offer_mfg, str(quote.unit_price), quote.currency],
            ensure_ascii=False,
        )
        if key in seen:
            continue
        seen.add(key)
        deduped.append(quote)
    return deduped


def _dedupe_strings(values: list[str]) -> list[str]:
    deduped: list[str] = []
    seen: set[str] = set()
    for value in values:
        marker = value.strip().upper()
        if not marker or marker in seen:
            continue
        seen.add(marker)
        deduped.append(value)
    return deduped


def _parse_int(value: Any) -> int | None:
    if value in (None, ""):
        return None
    match = re.search(r"\d+", str(value).replace(",", "").replace(" ", ""))
    if not match:
        return None
    return int(match.group(0))


def _parse_decimal(value: Any) -> Decimal | None:
    if value in (None, ""):
        return None
    cleaned = str(value).strip().replace("\xa0", " ").replace(" ", "")
    cleaned = cleaned.replace(",", ".")
    match = re.search(r"-?\d+(?:\.\d+)?", cleaned)
    if not match:
        return None
    try:
        return Decimal(match.group(0))
    except InvalidOperation:
        return None


def _part_key(value: str | None) -> str:
    return re.sub(r"[^A-Z0-9]", "", (value or "").upper())


def _brand_key(value: str | None) -> str:
    return re.sub(r"[^A-Z0-9]", "", (value or "").upper())


def _aggregate_status(statuses: list[str], quote_count: int) -> str:
    if quote_count:
        return "ready"
    priority = [
        "needs_manual_action",
        "blocked",
        "error",
        "needs_login",
        "needs_api",
        "needs_file",
        "no_data",
        "disabled",
    ]
    for status in priority:
        if status in statuses:
            return status
    return "no_data"
