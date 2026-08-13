from __future__ import annotations

import csv
import io
from decimal import Decimal
from pathlib import Path

from .models import SupplierQuote


def load_quotes_csv(path: str | Path) -> list[SupplierQuote]:
    with Path(path).open("r", encoding="utf-8-sig", newline="") as handle:
        return _load_quotes(handle)


def load_quotes_csv_text(text: str) -> list[SupplierQuote]:
    return _load_quotes(io.StringIO(text))


def _load_quotes(handle) -> list[SupplierQuote]:
    quotes: list[SupplierQuote] = []

    reader = csv.DictReader(handle)
    for line_no, row in enumerate(reader, start=2):
        unit_price = _required_decimal(row, "unit_price", line_no)
        quotes.append(
            SupplierQuote(
                row=_optional_int(row.get("row")),
                value=(row.get("value") or "").strip(),
                offer_pn=_optional_str(row.get("offer_pn")),
                offer_mfg=_optional_str(row.get("offer_mfg")),
                supplier=(row.get("supplier") or "Manual").strip(),
                currency=(row.get("currency") or "USD").strip().upper(),
                unit_price=unit_price,
                pack_qty=_optional_int(row.get("pack_qty")),
                stock_qty=_optional_int(row.get("stock_qty")),
                confidence=_optional_decimal(row.get("confidence")),
                source_url=_optional_str(row.get("source_url")),
            )
        )

    return quotes


def _optional_str(value: str | None) -> str | None:
    value = (value or "").strip()
    return value or None


def _optional_int(value: str | None) -> int | None:
    value = (value or "").strip()
    if not value:
        return None
    return int(float(value))


def _optional_decimal(value: str | None) -> Decimal | None:
    value = (value or "").strip()
    if not value:
        return None
    return Decimal(value)


def _required_decimal(row: dict[str, str], field: str, line_no: int) -> Decimal:
    value = (row.get(field) or "").strip()
    if not value:
        raise ValueError(f"Missing {field!r} at CSV line {line_no}.")
    return Decimal(value)
