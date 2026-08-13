from __future__ import annotations

import argparse
import sys
from decimal import Decimal
from pathlib import Path

from .quotes_csv import load_quotes_csv
from .rrfq_excel import fill_rrfq, load_visible_items


def main() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser(prog="rrfq")
    subparsers = parser.add_subparsers(dest="command", required=True)

    inspect_parser = subparsers.add_parser("inspect", help="Show visible RRFQ rows.")
    inspect_parser.add_argument("workbook")

    fill_parser = subparsers.add_parser("fill", help="Fill RRFQ from a quote CSV.")
    fill_parser.add_argument("workbook")
    fill_parser.add_argument("--quotes", required=True)
    fill_parser.add_argument("--output", required=True)
    fill_parser.add_argument("--usd-rub", type=Decimal)
    fill_parser.add_argument("--usd-cny", type=Decimal)

    args = parser.parse_args()

    if args.command == "inspect":
        _inspect(Path(args.workbook))
        return

    if args.command == "fill":
        quotes = load_quotes_csv(args.quotes)
        applied = fill_rrfq(
            args.workbook,
            args.output,
            quotes,
            usd_rub=args.usd_rub,
            usd_cny=args.usd_cny,
        )
        print(f"Applied quotes: {len(applied)}")
        for normalized in applied:
            quote = normalized.quote
            row = quote.row if quote.row is not None else "-"
            print(f"row={row} value={quote.value} supplier={quote.supplier} usd={normalized.unit_price_usd:.6f}")


def _inspect(path: Path) -> None:
    items = load_visible_items(path)
    print(f"Visible BOM rows: {len(items)}")
    for item in items:
        qty = item.qty_to_buy if item.qty_to_buy is not None else ""
        mfg = item.requested_mfg or ""
        print(f"row={item.row} #={item.item_no} value={item.value} mfg={mfg} qty_to_buy={qty}")


if __name__ == "__main__":
    main()
