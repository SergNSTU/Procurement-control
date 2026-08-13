from __future__ import annotations

import hashlib
import json
import os
import random
import string
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Any


class LcscClientError(RuntimeError):
    pass


@dataclass(frozen=True)
class LcscCredentials:
    key: str
    secret: str

    @classmethod
    def from_env(cls) -> "LcscCredentials | None":
        key = os.environ.get("LCSC_API_KEY", "").strip()
        secret = os.environ.get("LCSC_API_SECRET", "").strip()
        if not key or not secret:
            return None
        return cls(key=key, secret=secret)


class JsonCache:
    def __init__(self, path: str | Path, ttl_seconds: int = 60 * 60 * 12) -> None:
        self.path = Path(path)
        self.ttl_seconds = ttl_seconds
        self._data: dict[str, dict[str, Any]] | None = None
        self._lock = threading.RLock()

    def get(self, key: str) -> Any | None:
        with self._lock:
            data = self._load()
            entry = data.get(key)
            if not entry:
                return None
            if time.time() - float(entry.get("created_at", 0)) > self.ttl_seconds:
                return None
            return entry.get("value")

    def set(self, key: str, value: Any) -> None:
        with self._lock:
            data = self._load()
            data[key] = {"created_at": time.time(), "value": value}
            self.path.parent.mkdir(parents=True, exist_ok=True)
            self.path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    def _load(self) -> dict[str, dict[str, Any]]:
        if self._data is not None:
            return self._data
        if not self.path.exists():
            self._data = {}
            return self._data
        try:
            self._data = json.loads(self.path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            self._data = {}
        return self._data


class LcscClient:
    base_url = "https://ips.lcsc.com"

    def __init__(
        self,
        credentials: LcscCredentials | None,
        *,
        cache: JsonCache | None = None,
        min_interval_seconds: float = 0.75,
        timeout_seconds: float = 25,
    ) -> None:
        self.credentials = credentials
        self.cache = cache
        self.min_interval_seconds = min_interval_seconds
        self.timeout_seconds = timeout_seconds
        self._last_request_at = 0.0

    @property
    def is_configured(self) -> bool:
        return self.credentials is not None

    def search_products(
        self,
        keyword: str,
        *,
        currency: str = "USD",
        exact: bool = True,
        page_size: int = 30,
        is_available: bool = False,
    ) -> list[dict[str, Any]]:
        if not self.credentials:
            raise LcscClientError("LCSC API key and secret are not configured.")

        params = {
            "keyword": keyword,
            "currency": currency,
            "match_type": "exact" if exact else "fuzzy",
            "page_size": str(page_size),
            "current_page": "1",
            "is_available": "true" if is_available else "false",
            "is_pre_sale": "false",
        }
        response = self._get("/rest/wmsc2agent/search/product", params)
        return _extract_products(response)

    def product_info(self, product_number: str, *, currency: str = "USD") -> dict[str, Any]:
        if not self.credentials:
            raise LcscClientError("LCSC API key and secret are not configured.")

        response = self._get(
            f"/rest/wmsc2agent/product/info/{urllib.parse.quote(product_number)}",
            {"currency": currency},
        )
        result = response.get("result") if isinstance(response, dict) else response
        return result if isinstance(result, dict) else {}

    def _get(self, path: str, params: dict[str, str]) -> dict[str, Any]:
        unsigned_params = dict(params)
        cache_key = _cache_key(path, unsigned_params)
        if self.cache:
            cached = self.cache.get(cache_key)
            if cached is not None:
                return cached

        signed_params = self._signed_params(unsigned_params)
        url = f"{self.base_url}{path}?{urllib.parse.urlencode(signed_params)}"
        self._wait_for_rate_limit()
        request = urllib.request.Request(url, headers={"Accept": "application/json"})

        try:
            with urllib.request.urlopen(request, timeout=self.timeout_seconds) as response:
                body = response.read().decode("utf-8")
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace")
            raise LcscClientError(f"LCSC HTTP {exc.code}: {detail[:300]}") from exc
        except urllib.error.URLError as exc:
            raise LcscClientError(f"LCSC connection error: {exc.reason}") from exc

        data = json.loads(body)
        if isinstance(data, dict) and data.get("success") is False:
            raise LcscClientError(f"LCSC error {data.get('code')}: {data.get('message')}")

        if self.cache:
            self.cache.set(cache_key, data)
        return data

    def _signed_params(self, params: dict[str, str]) -> dict[str, str]:
        assert self.credentials is not None
        nonce = _nonce()
        timestamp = str(int(time.time()))
        signature_payload = (
            f"key={self.credentials.key}&nonce={nonce}"
            f"&secret={self.credentials.secret}&timestamp={timestamp}"
        )
        signed = dict(params)
        signed.update(
            {
                "key": self.credentials.key,
                "nonce": nonce,
                "timestamp": timestamp,
                "signature": hashlib.sha1(signature_payload.encode("utf-8")).hexdigest(),
            }
        )
        return signed

    def _wait_for_rate_limit(self) -> None:
        elapsed = time.monotonic() - self._last_request_at
        if elapsed < self.min_interval_seconds:
            time.sleep(self.min_interval_seconds - elapsed)
        self._last_request_at = time.monotonic()


def _nonce() -> str:
    alphabet = string.ascii_letters + string.digits
    return "".join(random.choice(alphabet) for _ in range(16))


def _cache_key(path: str, params: dict[str, str]) -> str:
    return json.dumps({"path": path, "params": params}, sort_keys=True, ensure_ascii=False)


def _extract_products(response: Any) -> list[dict[str, Any]]:
    result = response.get("result") if isinstance(response, dict) else response
    candidates: list[dict[str, Any]] = []

    def walk(node: Any) -> None:
        if isinstance(node, list):
            for item in node:
                walk(item)
            return
        if not isinstance(node, dict):
            return
        if _looks_like_product(node):
            candidates.append(node)
        for value in node.values():
            if isinstance(value, (list, dict)):
                walk(value)

    walk(result)
    deduped: list[dict[str, Any]] = []
    seen: set[str] = set()
    for product in candidates:
        marker = json.dumps(product, sort_keys=True, ensure_ascii=False)[:1000]
        if marker in seen:
            continue
        seen.add(marker)
        deduped.append(product)
    return deduped


def _looks_like_product(node: dict[str, Any]) -> bool:
    keys = {str(key).lower() for key in node.keys()}
    product_keys = {
        "productmodel",
        "product_model",
        "mpn",
        "productcode",
        "productno",
        "productnumber",
        "lcscpartnumber",
        "lcsc_part_number",
        "brandname",
        "manufacturer",
    }
    return bool(keys & product_keys)
