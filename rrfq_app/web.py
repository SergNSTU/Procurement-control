from __future__ import annotations

import argparse
import json
import mimetypes
import os
import re
import sys
import uuid
from decimal import Decimal
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse

from .compel_parser import compel_row_json, compel_summary_json, parse_compel_html, write_compel_xlsx
from .lcsc import JsonCache, LcscClient, LcscCredentials
from .resolver import quotes_from_decisions, resolve_items_from_sources
from .rrfq_excel import fill_rrfq, load_visible_items
from .supplier_sources import (
    SearchContext,
    build_sources,
    default_source_keys,
    parse_manual_quotes,
    source_definitions_json,
)


APP_ROOT = Path(__file__).resolve().parent
STATIC_ROOT = APP_ROOT / "static"
DATA_ROOT = Path.cwd() / "app_data"


class AppHandler(BaseHTTPRequestHandler):
    server_version = "RRFQPricer/0.1"

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self._send_static(STATIC_ROOT / "index.html")
            return
        if parsed.path == "/compel":
            self._send_static(STATIC_ROOT / "compel.html")
            return
        if parsed.path == "/api/health":
            self._send_json({"ok": True})
            return
        if parsed.path == "/api/config":
            self._send_json(
                {
                    "has_lcsc_credentials": LcscCredentials.from_env() is not None,
                    "sources": source_definitions_json(),
                    "default_sources": sorted(default_source_keys()),
                }
            )
            return
        if parsed.path.startswith("/download/"):
            self._download(parsed.path.removeprefix("/download/"))
            return
        if parsed.path.startswith("/download-compel/"):
            self._download_compel(parsed.path.removeprefix("/download-compel/"))
            return
        if parsed.path.startswith("/static/"):
            rel_path = parsed.path.removeprefix("/static/")
            self._send_static((STATIC_ROOT / rel_path).resolve())
            return
        self._send_error(HTTPStatus.NOT_FOUND, "Not found")

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/api/inspect":
            self._inspect_upload(parsed)
            return
        if parsed.path == "/api/price":
            self._price_job()
            return
        if parsed.path == "/api/compel/inspect":
            self._inspect_compel_upload(parsed)
            return
        if parsed.path == "/api/compel/export":
            self._export_compel_job()
            return
        self._send_error(HTTPStatus.NOT_FOUND, "Not found")

    def log_message(self, fmt: str, *args: object) -> None:
        print(f"{self.address_string()} - {fmt % args}", file=sys.stderr)

    def _inspect_upload(self, parsed) -> None:
        query = parse_qs(parsed.query)
        file_name = _safe_filename(unquote(query.get("filename", ["rrfq.xlsx"])[0]))
        if not file_name.lower().endswith(".xlsx"):
            self._send_error(HTTPStatus.BAD_REQUEST, "Please upload an .xlsx RRFQ file.")
            return

        content_length = int(self.headers.get("Content-Length", "0"))
        if content_length <= 0:
            self._send_error(HTTPStatus.BAD_REQUEST, "Empty upload.")
            return

        job_id = uuid.uuid4().hex
        job_dir = _job_dir(job_id)
        job_dir.mkdir(parents=True, exist_ok=True)
        input_path = job_dir / file_name
        input_path.write_bytes(self.rfile.read(content_length))

        try:
            items = load_visible_items(input_path)
        except Exception as exc:
            self._send_error(HTTPStatus.BAD_REQUEST, f"Could not read RRFQ: {exc}")
            return

        self._send_json(
            {
                "job_id": job_id,
                "file_name": file_name,
                "visible_count": len(items),
                "items": [_item_json(item) for item in items],
            }
        )

    def _price_job(self) -> None:
        payload = self._read_json()
        job_id = str(payload.get("job_id", "")).strip()
        if not _valid_job_id(job_id):
            self._send_error(HTTPStatus.BAD_REQUEST, "Invalid job id.")
            return

        job_dir = _job_dir(job_id)
        input_files = sorted(job_dir.glob("*.xlsx"))
        if not input_files:
            self._send_error(HTTPStatus.NOT_FOUND, "Uploaded file was not found.")
            return
        input_path = input_files[0]
        output_path = job_dir / f"{input_path.stem}_priced.xlsx"

        credentials = _credentials_from_payload(payload) or LcscCredentials.from_env()
        cache = JsonCache(DATA_ROOT / "cache" / "supplier_cache.json")
        client = LcscClient(credentials, cache=cache)
        enabled_sources = _enabled_sources(payload)
        sources = build_sources(enabled_sources, cache=cache, lcsc_client=client)
        context = SearchContext(
            usd_rub=_optional_decimal(payload.get("usd_rub")),
            usd_cny=_optional_decimal(payload.get("usd_cny")),
            manual_quotes=parse_manual_quotes(payload.get("manual_csv")),
        )

        try:
            items = load_visible_items(input_path)
            result = resolve_items_from_sources(
                items,
                sources,
                context,
                use_estimates=bool(payload.get("use_estimates", True)),
            )
            decisions = result.decisions
            fill_rrfq(
                input_path,
                output_path,
                quotes_from_decisions(decisions),
                usd_rub=context.usd_rub,
                usd_cny=context.usd_cny,
            )
        except Exception as exc:
            self._send_error(HTTPStatus.INTERNAL_SERVER_ERROR, f"Pricing failed: {exc}")
            return

        self._send_json(
            {
                "job_id": job_id,
                "download_url": f"/download/{job_id}",
                "summary": _summary(decisions),
                "source_statuses": [_source_status_json(status) for status in result.source_statuses],
                "results": [_decision_json(decision) for decision in decisions],
                "has_lcsc_credentials": client.is_configured,
            }
        )

    def _inspect_compel_upload(self, parsed) -> None:
        query = parse_qs(parsed.query)
        file_name = _safe_filename(unquote(query.get("filename", ["compel.html"])[0]))
        if not file_name.lower().endswith((".html", ".htm")):
            self._send_error(HTTPStatus.BAD_REQUEST, "Please upload a saved .html file.")
            return

        content_length = int(self.headers.get("Content-Length", "0"))
        if content_length <= 0:
            self._send_error(HTTPStatus.BAD_REQUEST, "Empty upload.")
            return

        job_id = uuid.uuid4().hex
        job_dir = _job_dir(job_id)
        job_dir.mkdir(parents=True, exist_ok=True)
        input_path = job_dir / file_name
        input_path.write_bytes(self.rfile.read(content_length))

        try:
            rows, summary = parse_compel_html(input_path)
        except Exception as exc:
            self._send_error(HTTPStatus.BAD_REQUEST, f"Could not parse Compel HTML: {exc}")
            return

        self._send_json(
            {
                "job_id": job_id,
                "file_name": file_name,
                "summary": compel_summary_json(summary),
                "rows": [compel_row_json(row) for row in rows],
            }
        )

    def _export_compel_job(self) -> None:
        payload = self._read_json()
        job_id = str(payload.get("job_id", "")).strip()
        if not _valid_job_id(job_id):
            self._send_error(HTTPStatus.BAD_REQUEST, "Invalid job id.")
            return

        job_dir = _job_dir(job_id)
        input_files = sorted([*job_dir.glob("*.html"), *job_dir.glob("*.htm")])
        if not input_files:
            self._send_error(HTTPStatus.NOT_FOUND, "Uploaded HTML file was not found.")
            return

        input_path = input_files[0]
        output_path = job_dir / f"{input_path.stem}_compel_prices.xlsx"
        try:
            rows, summary = parse_compel_html(input_path)
            write_compel_xlsx(output_path, rows, summary)
        except Exception as exc:
            self._send_error(HTTPStatus.INTERNAL_SERVER_ERROR, f"Excel export failed: {exc}")
            return

        self._send_json(
            {
                "job_id": job_id,
                "download_url": f"/download-compel/{job_id}",
                "summary": compel_summary_json(summary),
            }
        )

    def _download(self, job_id: str) -> None:
        if not _valid_job_id(job_id):
            self._send_error(HTTPStatus.BAD_REQUEST, "Invalid job id.")
            return
        files = sorted(_job_dir(job_id).glob("*_priced.xlsx"))
        if not files:
            self._send_error(HTTPStatus.NOT_FOUND, "Priced file was not found.")
            return
        path = files[-1]
        data = path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Disposition", f'attachment; filename="{path.name}"')
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _download_compel(self, job_id: str) -> None:
        if not _valid_job_id(job_id):
            self._send_error(HTTPStatus.BAD_REQUEST, "Invalid job id.")
            return
        files = sorted(_job_dir(job_id).glob("*_compel_prices.xlsx"))
        if not files:
            self._send_error(HTTPStatus.NOT_FOUND, "Compel Excel file was not found.")
            return
        path = files[-1]
        data = path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Disposition", f'attachment; filename="{path.name}"')
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _send_static(self, path: Path) -> None:
        try:
            resolved = path.resolve()
            if not str(resolved).startswith(str(STATIC_ROOT.resolve())):
                self._send_error(HTTPStatus.FORBIDDEN, "Forbidden")
                return
            data = resolved.read_bytes()
        except FileNotFoundError:
            self._send_error(HTTPStatus.NOT_FOUND, "Not found")
            return

        content_type = mimetypes.guess_type(str(resolved))[0] or "application/octet-stream"
        if resolved.suffix == ".html":
            content_type = "text/html; charset=utf-8"
        elif resolved.suffix == ".css":
            content_type = "text/css; charset=utf-8"
        elif resolved.suffix == ".js":
            content_type = "application/javascript; charset=utf-8"
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _send_json(self, payload: dict, status: HTTPStatus = HTTPStatus.OK) -> None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _send_error(self, status: HTTPStatus, message: str) -> None:
        self._send_json({"error": message}, status)

    def _read_json(self) -> dict:
        content_length = int(self.headers.get("Content-Length", "0"))
        if content_length <= 0:
            return {}
        return json.loads(self.rfile.read(content_length).decode("utf-8"))


def main() -> None:
    parser = argparse.ArgumentParser(prog="rrfq-web")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=int(os.environ.get("RRFQ_PORT", "8765")))
    args = parser.parse_args()

    DATA_ROOT.mkdir(parents=True, exist_ok=True)
    server = ThreadingHTTPServer((args.host, args.port), AppHandler)
    url = f"http://{args.host}:{server.server_port}"
    print(f"RRFQ Pricer is running at {url}", flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


def _job_dir(job_id: str) -> Path:
    return DATA_ROOT / "jobs" / job_id


def _safe_filename(value: str) -> str:
    name = Path(value).name or "rrfq.xlsx"
    return re.sub(r"[^A-Za-z0-9_. -]", "_", name)


def _valid_job_id(value: str) -> bool:
    return bool(re.fullmatch(r"[a-f0-9]{32}", value))


def _credentials_from_payload(payload: dict) -> LcscCredentials | None:
    key = str(payload.get("lcsc_key", "")).strip()
    secret = str(payload.get("lcsc_secret", "")).strip()
    if not key or not secret:
        return None
    return LcscCredentials(key=key, secret=secret)


def _optional_decimal(value: object) -> Decimal | None:
    if value in (None, ""):
        return None
    return Decimal(str(value))


def _enabled_sources(payload: dict) -> set[str]:
    raw = payload.get("enabled_sources")
    if not raw:
        enabled = default_source_keys()
    elif isinstance(raw, list):
        enabled = {str(item) for item in raw}
    else:
        enabled = {str(raw)}
    if str(payload.get("manual_csv", "")).strip():
        enabled.add("manual_csv")
    return enabled


def _item_json(item) -> dict:
    return {
        "row": item.row,
        "item_no": item.item_no,
        "value": item.value,
        "requested_mfg": item.requested_mfg,
        "category": item.category,
        "body": item.body,
        "qty_to_buy": _decimal_json(item.qty_to_buy),
    }


def _decision_json(decision) -> dict:
    quote = decision.quote.quote if decision.quote else None
    return {
        "row": decision.item.row,
        "item_no": decision.item.item_no,
        "value": decision.item.value,
        "requested_mfg": decision.item.requested_mfg,
        "category": decision.item.category,
        "body": decision.item.body,
        "qty_to_buy": _decimal_json(decision.item.qty_to_buy),
        "status": decision.status,
        "reason": decision.reason,
        "alternatives_count": decision.alternatives_count,
        "source": quote.supplier if quote else None,
        "offer_pn": quote.offer_pn if quote else None,
        "offer_mfg": quote.offer_mfg if quote else None,
        "unit_price_usd": _decimal_json(decision.quote.unit_price_usd if decision.quote else None),
        "pack_qty": quote.pack_qty if quote else None,
        "stock_qty": quote.stock_qty if quote else None,
        "confidence": _decimal_json(quote.confidence if quote else None),
        "is_estimated": bool(quote.is_estimated) if quote else False,
        "source_url": quote.source_url if quote else None,
        "offers": [_offer_json(offer) for offer in decision.offers],
    }


def _offer_json(offer) -> dict:
    quote = offer.quote
    return {
        "source": quote.supplier,
        "offer_pn": quote.offer_pn,
        "offer_mfg": quote.offer_mfg,
        "currency": quote.currency,
        "unit_price": _decimal_json(quote.unit_price),
        "unit_price_usd": _decimal_json(offer.unit_price_usd),
        "stock_qty": quote.stock_qty,
        "pack_qty": quote.pack_qty,
        "confidence": _decimal_json(quote.confidence),
        "reason": quote.reason,
        "source_url": quote.source_url,
        "is_estimated": quote.is_estimated,
    }


def _source_status_json(status) -> dict:
    return {
        "key": status.key,
        "label": status.label,
        "status": status.status,
        "reason": status.reason,
        "searched_items": status.searched_items,
        "quote_count": status.quote_count,
    }


def _summary(decisions) -> dict:
    total = len(decisions)
    found = sum(1 for decision in decisions if decision.status == "found")
    estimated = sum(1 for decision in decisions if decision.status == "estimated")
    missing = sum(1 for decision in decisions if decision.status == "not_enough_data")
    return {"total": total, "found": found, "estimated": estimated, "missing": missing}


def _decimal_json(value: Decimal | None) -> str | None:
    if value is None:
        return None
    return format(value, "f")


if __name__ == "__main__":
    main()
