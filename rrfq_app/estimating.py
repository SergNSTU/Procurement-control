from __future__ import annotations

from decimal import Decimal
from statistics import median

from .models import RrfqItem, SupplierQuote


def estimate_quote(
    item: RrfqItem,
    related_quotes: list[SupplierQuote],
    *,
    api_configured: bool,
    api_error: str | None = None,
    source_labels: list[str] | None = None,
) -> SupplierQuote | None:
    if not item.value.strip():
        return None

    related_prices = [quote.unit_price for quote in related_quotes if quote.unit_price > 0]
    if related_prices:
        base = Decimal(str(median([float(price) for price in related_prices])))
        confidence = Decimal("0.55")
        source_hint = "по похожим ценовым предложениям"
    else:
        base = _heuristic_base_price(item)
        confidence = Decimal("0.25") if api_configured else Decimal("0.18")
        source_hint = "по типу компонента, корпусу и количеству"

    adjusted = base * _package_multiplier(item) * _manufacturer_multiplier(item) * _quantity_multiplier(item)
    adjusted = _round_price(adjusted)

    if adjusted <= 0:
        return None

    source_text = ", ".join(source_labels or [])
    if source_text:
        reason = f"Прямая цена не найдена в выбранных источниках ({source_text}). Прогноз {source_hint}."
    else:
        reason = f"Прямая цена не найдена. Прогноз {source_hint}."
    if api_error:
        reason = f"{reason} Примечания: {api_error}"

    return SupplierQuote(
        row=item.row,
        value=item.value,
        offer_pn=item.value,
        offer_mfg=item.requested_mfg,
        supplier="Estimate",
        currency="USD",
        unit_price=adjusted,
        confidence=confidence,
        is_estimated=True,
        reason=reason,
    )


def _heuristic_base_price(item: RrfqItem) -> Decimal:
    text = " ".join(
        part for part in [item.value, item.category or "", item.body or "", item.requested_mfg or ""] if part
    ).upper()

    if any(token in text for token in ["TMS320", "STM32", "MSP430", "PIC", "LQFP", "BGA", "MCU", "DSP"]):
        return Decimal("9.50")
    if any(token in text for token in ["FRAM", "FLASH", "EEPROM", "SRAM", "MEMORY"]):
        return Decimal("1.80")
    if any(token in text for token in ["ISOL", "OPTO", "CPC", "MOC", "PS270", "CA-IS"]):
        return Decimal("0.85")
    if any(token in text for token in ["MOSFET", "IRF", "CEB", "TRANSISTOR"]):
        return Decimal("0.45")
    if any(token in text for token in ["DIODE", "SMBJ", "BAV", "BZX", "TVS", "10CTQ", "US1"]):
        return Decimal("0.12")
    if any(token in text for token in ["OPAMP", "COMPARATOR", "LM", "TL", "MC", "REGULATOR", "AMS1117"]):
        return Decimal("0.18")
    if any(token in text for token in ["RESISTOR", "CAPACITOR", "INDUCTOR", "0805", "1206", "0603", "2512"]):
        return Decimal("0.018")
    return Decimal("0.35")


def _package_multiplier(item: RrfqItem) -> Decimal:
    body = (item.body or "").upper()
    if any(token in body for token in ["BGA", "QFN", "LQFP", "TQFP"]):
        return Decimal("1.25")
    if any(token in body for token in ["TO-263", "D2PAK", "DPAK", "TO-252"]):
        return Decimal("1.15")
    if any(token in body for token in ["0805", "0603", "0402", "1206"]):
        return Decimal("0.85")
    return Decimal("1.00")


def _manufacturer_multiplier(item: RrfqItem) -> Decimal:
    mfg = (item.requested_mfg or "").upper()
    premium = ["TI", "TEXAS", "INFINEON", "VISHAY", "RENESAS", "ANALOG", "ADI", "ST"]
    value = item.value.upper()
    if any(token in mfg for token in premium) or any(token in value for token in ["IRFR", "TMS320"]):
        return Decimal("1.20")
    return Decimal("1.00")


def _quantity_multiplier(item: RrfqItem) -> Decimal:
    qty = item.qty_to_buy or item.qty_in_order
    if qty is None:
        return Decimal("1.00")
    if qty >= 1000:
        return Decimal("0.78")
    if qty >= 500:
        return Decimal("0.86")
    if qty >= 100:
        return Decimal("0.95")
    return Decimal("1.10")


def _round_price(value: Decimal) -> Decimal:
    if value >= Decimal("10"):
        return value.quantize(Decimal("0.01"))
    if value >= Decimal("1"):
        return value.quantize(Decimal("0.001"))
    return value.quantize(Decimal("0.0001"))
