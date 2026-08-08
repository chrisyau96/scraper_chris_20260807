#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HKTVmall 商品爬蟲 – Streamlit 網頁版
支援關鍵字 / 分類 / 混合 / 全站掃描，自動分割任務突破 API 分頁上限。
"""

from __future__ import annotations

import hashlib
import io
import json
import os
import re
import time
import zipfile
from copy import deepcopy
from datetime import datetime, timezone
from typing import Any, Callable, Iterable
from urllib.parse import quote, unquote

import pandas as pd
import requests
import streamlit as st
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from PIL import Image as PILImage
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# =============================================================================
# 常數
# =============================================================================

APP_VERSION = "1.4-webapp-unlimited"
HKTV_SEARCH_URL = "https://keyword-search-server.hktvmall.com/api/search"
HKTV_API_KEY = "0e6c95ec-4c8b-4f71-8855-11eeafe74966"
HKTV_INDEX_NAME = "hktvProduct"

API_MAX_OFFSET = 10000
DEFAULT_PAGE_SIZE = 60
DEFAULT_PRICE_RANGE_BUCKETS = [
    "0-50",
    "50-100",
    "100-200",
    "200-500",
    "500-1000",
    "1000-3000",
    "3000-10000",
    "10000-999999",
]
MAX_SPLIT_DEPTH = 6
MAX_SUBCATEGORY_DEPTH = 4
MAX_BRAND_SPLITS = 30

CHECKPOINT_DIR = "hktv_checkpoints"
BACKUP_DIR = "hktv_backups"
EXPORT_DIR = "hktv_exports"
IMAGE_CACHE_DIR = "hktv_image_cache"
PREVIEW_PAGE_SIZE_OPTIONS = (100, 500, 1000, 5000, 0)  # 0 = show all rows

WEBSITES: dict[str, dict[str, str]] = {
    "hktv_zh": {
        "label": "HKTVmall 繁體",
        "lang": "zh",
        "product_url_base": "https://www.hktvmall.com/hktv/zh/main/p/",
        "search_url_base": "https://www.hktvmall.com/hktv/zh/main/search",
        "pdp_url_base": "https://www.hktvmall.com/hktv/zh/main/p/",
    },
    "hktv_en": {
        "label": "HKTVmall English",
        "lang": "en",
        "product_url_base": "https://www.hktvmall.com/hktv/en/main/p/",
        "search_url_base": "https://www.hktvmall.com/hktv/en/main/search",
        "pdp_url_base": "https://www.hktvmall.com/hktv/en/main/p/",
    },
}

EXCEL_COLUMNS: list[str] = [
    "商品圖片",
    "商品名稱",
    "品牌",
    "現價 (HK$)",
    "原價 (HK$)",
    "庫存狀態",
    "銷量",
    "評分",
    "評論數",
    "商戶",
    "分類",
    "分類路徑",
    "商品編號",
    "產地",
    "規格",
    "商品連結",
    "抓取時間",
]

# Internal keys kept in session rows for processing / dedup.
INTERNAL_ROW_KEYS = (
    "product_code",
    "name",
    "brand",
    "selling_price",
    "original_price",
    "availability",
    "sales_count",
    "average_rating",
    "number_of_reviews",
    "store",
    "category",
    "category_path",
    "primary_cat_code",
    "country_of_origin",
    "packing_spec",
    "summary",
    "image",
    "image_local",
    "product_url",
    "scraped_at",
    "keyword",
    "task_categories",
    "price_range_filter",
    "brand_filter",
    "source_mode",
)

SORT_OPTIONS: dict[str, tuple[str, bool] | None] = {
    "預設（抓取順序）": None,
    "商品編號 A→Z": ("商品編號", True),
    "商品編號 Z→A": ("商品編號", False),
    "現價 低→高": ("現價 (HK$)", True),
    "現價 高→低": ("現價 (HK$)", False),
    "銷量 少→多": ("銷量", True),
    "銷量 多→少": ("銷量", False),
    "評分 低→高": ("評分", True),
    "評分 高→低": ("評分", False),
    "評論數 少→多": ("評論數", True),
    "評論數 多→少": ("評論數", False),
    "品牌 A→Z": ("品牌", True),
    "名稱 A→Z": ("商品名稱", True),
}

API_SORT_OPTIONS: dict[str, str] = {
    "銷量（高→低）": "salesVolume:desc",
    "價格（低→高）": "price:asc",
    "價格（高→低）": "price:desc",
    "相關性 / 預設": "",
}

SKIP_BRAND_FACETS = frozenset({"OtherBrands", "ShippedfromMainland"})

# =============================================================================
# 通用工具
# =============================================================================


def clean(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    if isinstance(value, bool):
        return "是" if value else "否"
    if isinstance(value, (dict, list)):
        return display_text(value)
    text = str(value).strip()
    return re.sub(r"\s+", " ", text)


def parse_jsonish(value: Any) -> Any:
    if isinstance(value, str):
        raw = value.strip()
        if (raw.startswith("{") and raw.endswith("}")) or (raw.startswith("[") and raw.endswith("]")):
            try:
                return json.loads(raw)
            except json.JSONDecodeError:
                try:
                    return json.loads(unquote(raw))
                except json.JSONDecodeError:
                    return value
    return value


def display_fragments(value: Any) -> list[str]:
    value = parse_jsonish(value)
    if value is None or isinstance(value, bool):
        return []
    if isinstance(value, (int, float)):
        return [clean(value)]
    if isinstance(value, str):
        text = clean(unquote(value))
        return [text] if text else []
    if isinstance(value, (list, tuple, set)):
        out: list[str] = []
        for item in value:
            out.extend(display_fragments(item))
        return unique_strings(out)
    if isinstance(value, dict):
        for key in (
            "nameZh", "nameTc", "name", "nameEn", "nameZhCN",
            "labelZh", "label", "text", "value", "formattedValue", "code",
        ):
            if key in value:
                out = display_fragments(value.get(key))
                if out:
                    return out
        out = []
        for nested in value.values():
            out.extend(display_fragments(nested))
        return unique_strings(out)
    return [clean(value)] if clean(value) else []


def unique_strings(values: Iterable[str]) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for value in values:
        text = clean(value)
        if text and text not in seen:
            seen.add(text)
            out.append(text)
    return out


def display_text(value: Any, separator: str = " / ") -> str:
    return separator.join(unique_strings(display_fragments(value)))


def category_text(value: Any) -> str:
    return display_text(value, separator="; ")


def numeric_value(value: Any, *, allow_zero: bool = False) -> int | float | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        num = float(value)
        if num < 0:
            return None
        if num == 0 and not allow_zero:
            return None
        return int(num) if num.is_integer() else num
    text = clean(value).replace("HK$", "").replace("$", "").replace(",", "").strip()
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    if not match:
        return None
    num = float(match.group(0))
    if num < 0:
        return None
    if num == 0 and not allow_zero:
        return None
    return int(num) if num.is_integer() else num


def price_from_list(rows: Any, wanted_types: set[str] | None = None) -> int | float | None:
    if not isinstance(rows, list):
        rows = [rows] if isinstance(rows, dict) else []
    for row in rows:
        if not isinstance(row, dict):
            continue
        price_type = clean(row.get("priceType")).upper()
        if wanted_types and price_type not in wanted_types:
            continue
        value = numeric_value(first(row.get("value"), row.get("formattedValue")))
        if value is not None:
            return value
    return None


def normalise_availability(value: Any) -> str:
    if isinstance(value, bool):
        return "有貨" if value else "缺貨"
    if isinstance(value, dict):
        status = value.get("stockLevelStatus")
        if isinstance(status, dict):
            code = clean(status.get("code")).casefold()
            if code in {"instock", "lowstock"}:
                return "有貨"
            if code in {"outofstock", "soldout"}:
                return "缺貨"
        if value.get("forceInStock"):
            return "有貨"
    text = display_text(value).casefold()
    if any(x in text for x in ("缺貨", "售罄", "outofstock", "soldout", "無貨")):
        return "缺貨"
    if any(x in text for x in ("有貨", "instock", "現貨", "available")):
        return "有貨"
    return clean(value)


def hktv_category_path(source: dict[str, Any]) -> str:
    raw_items = source.get("categoryStructureDisplay") or []
    if not isinstance(raw_items, list):
        raw_items = [raw_items]
    for raw in raw_items:
        try:
            nodes = json.loads(unquote(clean(raw)))
        except (TypeError, ValueError, json.JSONDecodeError):
            continue
        if not isinstance(nodes, list):
            continue
        names = [
            category_text(first(node.get("nameZh"), node.get("name"), node.get("nameEn"), node.get("code")))
            for node in nodes
            if isinstance(node, dict)
        ]
        path = " > ".join(name for name in names if name)
        if path:
            return path
    levels = [
        category_text(first(source.get("mainCatNameZh"), source.get("mainCatNameEn"))),
        category_text(source.get("subCat1NameZh")),
        category_text(source.get("subCat2NameZh")),
        category_text(source.get("subCat3NameZh")),
        category_text(source.get("subCat4NameZh")),
    ]
    return " > ".join(level for level in levels if level)


def extract_brand(source: dict[str, Any]) -> str:
    for key in ("brandZh", "brandDisplay", "brand", "brandEn"):
        value = source.get(key)
        parsed = parse_jsonish(value)
        text = display_text(parsed) if parsed is not None else clean(value)
        if text and not text.startswith("{"):
            return text
    return ""


def row_to_display_record(row: dict[str, Any]) -> dict[str, Any]:
    current = numeric_value(row.get("selling_price"), allow_zero=True)
    original = numeric_value(row.get("original_price"), allow_zero=True)
    sales = numeric_value(row.get("sales_count"), allow_zero=True)
    rating = numeric_value(row.get("average_rating"), allow_zero=True)
    reviews = numeric_value(row.get("number_of_reviews"), allow_zero=True)
    return {
        "商品圖片": clean(row.get("image")),
        "商品名稱": clean(row.get("name")),
        "品牌": clean(row.get("brand")),
        "現價 (HK$)": current,
        "原價 (HK$)": original,
        "庫存狀態": clean(row.get("availability")),
        "銷量": sales,
        "評分": rating,
        "評論數": reviews,
        "商戶": clean(row.get("store")),
        "分類": clean(row.get("category")),
        "分類路徑": clean(row.get("category_path")),
        "商品編號": clean(row.get("product_code")),
        "產地": clean(row.get("country_of_origin")),
        "規格": clean(row.get("packing_spec")),
        "商品連結": clean(row.get("product_url")),
        "抓取時間": clean(row.get("scraped_at")),
        "_image_local": clean(row.get("image_local")),
    }


def first(*values: Any) -> Any:
    for value in values:
        if value is None:
            continue
        if isinstance(value, str) and not value.strip():
            continue
        if isinstance(value, float) and pd.isna(value):
            continue
        return value
    return None


def unique(seq: list[Any]) -> list[Any]:
    seen: set[Any] = set()
    out: list[Any] = []
    for item in seq:
        if item in seen:
            continue
        seen.add(item)
        out.append(item)
    return out


def now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def ensure_dir(path: str) -> None:
    os.makedirs(path, exist_ok=True)


def safe_filename(name: str, max_len: int = 80) -> str:
    text = re.sub(r"[^\w\-.]+", "_", clean(name), flags=re.UNICODE)
    return (text or "item")[:max_len]


def merge_filters(base: dict | None, extra: dict | None) -> dict:
    merged: dict[str, Any] = {}
    if base:
        merged.update(base)
    if extra:
        for key, value in extra.items():
            if value is None:
                continue
            if isinstance(value, list) and not value:
                continue
            if key in merged and isinstance(merged[key], list) and isinstance(value, list):
                merged[key] = unique(merged[key] + value)
            else:
                merged[key] = value
    return merged


def is_top_level_category_code(code: str) -> bool:
    if not code:
        return False
    if not code.startswith("AA"):
        return True
    return len(code) == 14 and code.endswith("0000000")


def append_activity(message: str, level: str = "info") -> None:
    init_state()
    entry = {"time": now_iso(), "level": level, "message": message}
    st.session_state.activity_log.append(entry)
    st.session_state.activity_log = st.session_state.activity_log[-500:]
    st.session_state.last_message = message


def website_cfg(site_key: str | None = None) -> dict[str, str]:
    key = site_key or st.session_state.get("website_key", "hktv_zh")
    return WEBSITES.get(key, WEBSITES["hktv_zh"])


# =============================================================================
# HTTP Session
# =============================================================================


def build_session() -> requests.Session:
    session = requests.Session()
    retry = Retry(
        total=3,
        connect=3,
        read=3,
        backoff_factor=0.6,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=("GET", "POST"),
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=20, pool_maxsize=20)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session.headers.update(
        {
            "accept": "application/json, text/html, */*",
            "Content-Type": "application/json",
            "User-Agent": "Mozilla/5.0 (compatible; HKTVmallScraper/1.4)",
            "Authorization": f"ApiKey {HKTV_API_KEY}",
        }
    )
    return session


def get_session() -> requests.Session:
    init_state()
    if st.session_state.http_session is None:
        st.session_state.http_session = build_session()
    return st.session_state.http_session


def close_session() -> None:
    session = st.session_state.get("http_session")
    if session is not None:
        try:
            session.close()
        except Exception:
            pass
    st.session_state.http_session = None


# =============================================================================
# 分頁 / 任務輔助（v1.4 新增）
# =============================================================================


def api_max_page_number(page_size: int) -> int:
    if page_size <= 0:
        return 0
    return max(0, (API_MAX_OFFSET // page_size) - 1)


def should_attempt_split_on_end(task: dict, page_number: int, hits: list, page_size: int) -> bool:
    """True when pagination likely hit the API cliff before all products were fetched."""
    if int(task.get("split_depth") or 0) >= MAX_SPLIT_DEPTH:
        return False
    if len(hits) >= page_size:
        return False
    if page_number <= 0 and not hits:
        return False
    # Paginated at least once, or already collected products, but page is now empty/partial.
    if page_number > 0 or int(task.get("products_collected") or 0) > 0:
        return True
    return False


def try_split_task(
    session: requests.Session,
    tasks: list[dict],
    task_idx: int,
    task: dict,
    *,
    reason: str,
    timeout: float,
) -> dict[str, Any] | None:
    """Attempt to split an oversized task; return a status dict if split succeeded."""
    sub_tasks = split_oversized_task(session, task, timeout=timeout)
    if not sub_tasks:
        return None
    insert_split_tasks(tasks, task_idx, sub_tasks)
    st.session_state.stats["tasks_split"] += len(sub_tasks)
    task["done"] = True
    st.session_state.stats["tasks_completed"] += 1
    st.session_state.task_idx = task_idx + 1
    msg = f"已分割任務（{reason}）→ {len(sub_tasks)} 個子任務"
    append_activity(msg, "warning")
    return {"status": "split", "message": msg, "added_tasks": len(sub_tasks)}


def build_task_extra_filter(task: dict) -> dict[str, Any]:
    extra: dict[str, Any] = {}
    if task.get("price_range"):
        extra["priceRange"] = [task["price_range"]]
    if task.get("brand"):
        extra["brand"] = [task["brand"]]
    return extra


def task_page_signature(
    task_idx: int,
    page_number: int,
    categories: list[str],
    keyword: str,
    price_range: str | None,
    brand: str | None,
) -> str:
    payload = json.dumps(
        {
            "task_idx": task_idx,
            "page_number": page_number,
            "categories": categories,
            "keyword": keyword,
            "price_range": price_range,
            "brand": brand,
        },
        sort_keys=True,
        ensure_ascii=False,
    )
    return hashlib.sha1(payload.encode("utf-8")).hexdigest()


def insert_split_tasks(tasks: list[dict], task_idx: int, sub_tasks: list[dict]) -> None:
    if not sub_tasks:
        return
    tasks[task_idx + 1 : task_idx + 1] = sub_tasks


def fetch_brand_facets(
    session: requests.Session,
    *,
    keyword: str = "",
    categories: list[str] | None = None,
    extra_filter: dict | None = None,
    limit: int = MAX_BRAND_SPLITS,
    timeout: float = 30.0,
) -> list[str]:
    filter_obj: dict[str, Any] = {}
    if categories:
        filter_obj["category"] = list(categories)
    result = hktv_search_request(
        session,
        keyword=keyword,
        page_number=0,
        page_size=1,
        filter_obj=filter_obj,
        aggregations=["brand"],
        extra_filter=extra_filter,
        timeout=timeout,
    )
    brand_agg = result.get("aggregations", {}).get("brand") or {}
    if not isinstance(brand_agg, dict):
        return []
    brands = [
        b
        for b, _ in sorted(brand_agg.items(), key=lambda kv: kv[1], reverse=True)
        if b and b not in SKIP_BRAND_FACETS
    ]
    return brands[:limit]


# =============================================================================
# API 層
# =============================================================================


def hktv_search_request(
    session: requests.Session,
    *,
    keyword: str = "",
    page_number: int = 0,
    page_size: int = DEFAULT_PAGE_SIZE,
    filter_obj: dict | None = None,
    aggregations: list[str] | None = None,
    extra_filter: dict | None = None,
    sort_by: str = "salesVolume:desc",
    timeout: float = 30.0,
) -> dict[str, Any]:
    merged_filter = merge_filters(filter_obj, extra_filter)

    query_parts: list[str] = []
    for code in merged_filter.get("category") or []:
        query_parts.append(f":category:{code}")
    if sort_by:
        request_query = f"{''.join(query_parts)}:{sort_by}" if query_parts else f":{sort_by}"
    else:
        request_query = "".join(query_parts)

    request_body: dict[str, Any] = {
        "indexName": HKTV_INDEX_NAME,
        "keyword": keyword or "",
        "highlight": ["*"],
        "page": {"pageNumber": page_number, "pageSize": page_size},
        "filter": merged_filter,
    }
    if sort_by:
        request_body["sort"] = sort_by
    if request_query:
        request_body["query"] = request_query
    if aggregations is not None:
        request_body["aggregations"] = aggregations

    response = session.post(
        HKTV_SEARCH_URL,
        headers={
            "accept": "application/json",
            "accept-language": "zh-HK,zh;q=0.9,en;q=0.7",
            "content-type": "application/json",
            "authorization": f"ApiKey {HKTV_API_KEY}",
            "origin": "https://www.hktvmall.com",
            "referer": "https://www.hktvmall.com/hktv/zh/",
            "user-agent": (
                "Mozilla/5.0 (Linux; Android 10; Mobile) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/146.0.0.0 Mobile Safari/537.36"
            ),
        },
        json={"requests": [request_body]},
        timeout=timeout,
    )
    response.raise_for_status()
    payload = response.json()
    if payload.get("code") not in (None, 200) and not payload.get("results"):
        raise RuntimeError(payload.get("message") or "HKTV search API error")
    results = payload.get("results") or []
    if not results:
        return {"hits": [], "total": 0, "aggregations": {}, "raw": payload}
    block = results[0]
    return {
        "hits": block.get("hits") or [],
        "total": block.get("totalValue") or block.get("total") or 0,
        "aggregations": block.get("aggregations") or {},
        "raw": payload,
    }


def hktv_fetch_page(
    session: requests.Session,
    *,
    keyword: str = "",
    page_number: int = 0,
    page_size: int = DEFAULT_PAGE_SIZE,
    categories: list[str] | None = None,
    aggregations: list[str] | None = None,
    extra_filter: dict | None = None,
    sort_by: str = "salesVolume:desc",
    timeout: float = 30.0,
) -> dict[str, Any]:
    filter_obj: dict[str, Any] = {}
    if categories:
        filter_obj["category"] = list(categories)
    return hktv_search_request(
        session,
        keyword=keyword,
        page_number=page_number,
        page_size=page_size,
        filter_obj=filter_obj,
        aggregations=aggregations,
        extra_filter=extra_filter,
        sort_by=sort_by,
        timeout=timeout,
    )


def fetch_top_level_categories(
    session: requests.Session,
    timeout: float = 30.0,
) -> list[dict[str, Any]]:
    result = hktv_search_request(
        session,
        keyword="",
        page_number=0,
        page_size=1,
        filter_obj={},
        aggregations=["category"],
        timeout=timeout,
    )
    agg = result.get("aggregations", {}).get("category") or {}
    if not isinstance(agg, dict):
        return []
    items: list[dict[str, Any]] = []
    for code, count in sorted(agg.items(), key=lambda kv: kv[1], reverse=True):
        if not is_top_level_category_code(code):
            continue
        items.append({"code": code, "name": code, "count": int(count)})
    return items


def discover_child_category_codes(
    session: requests.Session,
    parent_code: str,
    timeout: float = 30.0,
    *,
    keyword: str = "",
    extra_filter: dict | None = None,
) -> list[dict[str, str]]:
    discovered: dict[str, str] = {}

    def collect_from_hits(hits: list[dict]) -> None:
        for hit in hits:
            src = hit.get("source") or hit
            code = clean(src.get("primaryCatCode"))
            if not code or code == parent_code:
                continue
            cat_display = src.get("categoryStructureDisplay")
            if isinstance(cat_display, list) and cat_display:
                name = clean(cat_display[-1])
            else:
                name = first(
                    src.get("catNameZh"),
                    src.get("mainCatNameZh"),
                    src.get("subCat1NameZh"),
                    code,
                )
            discovered[code] = clean(name) or code

    base_filter = merge_filters({"category": [parent_code]}, extra_filter)

    first_page = hktv_search_request(
        session,
        keyword=keyword,
        page_number=0,
        page_size=DEFAULT_PAGE_SIZE,
        filter_obj=base_filter,
        aggregations=["category", "brand"],
        timeout=timeout,
    )
    collect_from_hits(first_page.get("hits") or [])

    top_brands = fetch_brand_facets(
        session,
        keyword=keyword,
        categories=[parent_code],
        extra_filter=extra_filter,
        limit=8,
        timeout=timeout,
    )

    for bucket in DEFAULT_PRICE_RANGE_BUCKETS:
        bucket_filter = merge_filters(base_filter, {"priceRange": [bucket]})
        sampled = hktv_search_request(
            session,
            keyword=keyword,
            page_number=0,
            page_size=DEFAULT_PAGE_SIZE,
            filter_obj=bucket_filter,
            aggregations=["category", "brand"],
            timeout=timeout,
        )
        collect_from_hits(sampled.get("hits") or [])

    for brand in top_brands:
        brand_filter = merge_filters(base_filter, {"brand": [brand]})
        sampled = hktv_search_request(
            session,
            keyword=keyword,
            page_number=0,
            page_size=DEFAULT_PAGE_SIZE,
            filter_obj=brand_filter,
            aggregations=["category", "brand"],
            timeout=timeout,
        )
        collect_from_hits(sampled.get("hits") or [])

    children = [{"code": code, "name": name} for code, name in discovered.items() if code != parent_code]
    children.sort(key=lambda item: item["code"])
    return children


def fetch_subcategories_for_code(
    session: requests.Session,
    parent_code: str,
    timeout: float = 30.0,
    *,
    keyword: str = "",
    extra_filter: dict | None = None,
) -> list[dict[str, str]]:
    return discover_child_category_codes(
        session,
        parent_code,
        timeout,
        keyword=keyword,
        extra_filter=extra_filter,
    )


# =============================================================================
# 商品解析
# =============================================================================


def hktv_pdp_fallback_fields(
    session: requests.Session,
    product_code: str,
    site_key: str = "hktv_zh",
    timeout: float = 20.0,
) -> dict[str, Any]:
    """從商品詳情頁補充搜尋 API 缺少的欄位。"""
    if not product_code:
        return {}
    cfg = WEBSITES.get(site_key, WEBSITES["hktv_zh"])
    url = cfg["pdp_url_base"] + quote(product_code, safe="")
    try:
        response = session.get(url, timeout=timeout, headers={"accept": "text/html"})
        response.raise_for_status()
        html = response.text
    except Exception:
        return {}

    fields: dict[str, Any] = {}
    match = re.search(r"var\s+productData\s*=\s*(\{.*?\});\s*\n", html, re.S)
    if match:
        try:
            data = json.loads(match.group(1))
            fields.update(
                {
                    "name": first(data.get("name"), data.get("nameZh")),
                    "brand": first(data.get("brandName"), data.get("brand")),
                    "selling_price": first(
                        (data.get("price") or {}).get("value"),
                        data.get("sellingPrice"),
                    ),
                    "saved_price": first(
                        (data.get("price") or {}).get("saved"),
                        data.get("savedPrice"),
                    ),
                    "average_rating": data.get("averageRating"),
                    "number_of_reviews": data.get("numberOfReviews"),
                    "stock": data.get("stock"),
                    "in_stock": data.get("purchasable"),
                    "summary": first(data.get("summary"), data.get("description")),
                    "packing_spec": data.get("packingSpec"),
                    "store": first(data.get("storeName"), data.get("storeDisplay")),
                    "store_code": data.get("storeCode"),
                    "country_of_origin": data.get("countryOfOrigin"),
                    "image": first(
                        (data.get("images") or [None])[0] if isinstance(data.get("images"), list) else None,
                        data.get("imageLink"),
                    ),
                    "product_url": first(data.get("url"), url),
                }
            )
            categories = data.get("categories") or []
            if categories:
                names = [clean(c.get("name")) for c in categories if isinstance(c, dict)]
                fields["category_path"] = display_fragments(names)
                if names:
                    fields["main_category"] = names[0]
                    if len(names) > 1:
                        fields["sub_category"] = names[1]
        except json.JSONDecodeError:
            pass

    if not fields.get("name"):
        for block in re.findall(r'<script type="application/ld\+json">(.*?)</script>', html, re.S):
            try:
                ld = json.loads(block)
            except json.JSONDecodeError:
                continue
            if ld.get("@type") != "Product":
                continue
            fields.setdefault("name", ld.get("name"))
            brand = ld.get("brand")
            if isinstance(brand, dict):
                fields.setdefault("brand", brand.get("name"))
            offers = ld.get("offers") or {}
            if isinstance(offers, dict):
                fields.setdefault("selling_price", offers.get("price"))
            rating = ld.get("aggregateRating") or {}
            if isinstance(rating, dict):
                fields.setdefault("average_rating", rating.get("ratingValue"))
                fields.setdefault("number_of_reviews", rating.get("reviewCount"))
            images = ld.get("image")
            if images and not fields.get("image"):
                fields["image"] = images[0] if isinstance(images, list) else images
            break
    return {k: v for k, v in fields.items() if v not in (None, "", [])}


def parse_hktv_hit(
    hit: dict,
    *,
    task: dict,
    site_key: str = "hktv_zh",
    session: requests.Session | None = None,
    use_pdp_fallback: bool = False,
) -> dict[str, Any]:
    src = hit.get("source") or hit
    cfg = WEBSITES.get(site_key, WEBSITES["hktv_zh"])
    code = clean(src.get("code") or src.get("productSearchCode"))
    name = first(src.get("nameZh"), src.get("nameEn"), src.get("nameZhCN"), src.get("name"))
    brand = extract_brand(src)
    breadcrumb = hktv_category_path(src)
    category = category_text(first(src.get("catNameZh"), breadcrumb.split(" > ")[-1] if breadcrumb else ""))
    current_price = numeric_value(
        first(src.get("sellingPrice"), src.get("currentPrice"), src.get("discountPrice"))
    )
    original_price = numeric_value(
        first(src.get("basePrice"), src.get("listPrice"), src.get("originalPrice"), src.get("regularPrice"))
    )
    if original_price is None:
        original_price = price_from_list(src.get("savedPrice"), {"BUY", "BASE", "LIST", "ORIGINAL", "REGULAR"})
    if original_price is None:
        original_price = price_from_list(src.get("priceList"), {"BUY", "BASE", "LIST", "ORIGINAL", "REGULAR"})
    image = first(
        src.get("imageLink"),
        src.get("imageUrl"),
        src.get("thumbnailUrl"),
        (src.get("images") or [None])[0] if isinstance(src.get("images"), list) else None,
        (src.get("gallery") or [None])[0] if isinstance(src.get("gallery"), list) else None,
    )
    if isinstance(image, dict):
        image = first(image.get("url"), image.get("imageUrl"))
    product_url = first(src.get("urlZh"), src.get("urlEn"), src.get("url"))
    if product_url:
        product_url = clean(product_url)
        if product_url.startswith("//"):
            product_url = "https:" + product_url
        elif product_url.startswith("/"):
            product_url = "https://www.hktvmall.com" + product_url
        elif product_url.startswith("main/"):
            product_url = f"https://www.hktvmall.com/hktv/zh/{product_url}"
    if not product_url and code:
        product_url = cfg["product_url_base"] + code

    row: dict[str, Any] = {
        "product_code": code,
        "name": clean(name),
        "brand": brand,
        "selling_price": current_price,
        "original_price": original_price,
        "availability": normalise_availability(first(src.get("hasStock"), src.get("stock"))),
        "sales_count": numeric_value(first(src.get("salesVolume"), src.get("soldCount")), allow_zero=True),
        "average_rating": numeric_value(first(src.get("averageRating"), src.get("rating")), allow_zero=True),
        "number_of_reviews": numeric_value(first(src.get("numberOfReviews"), src.get("reviewCount")), allow_zero=True),
        "store": clean(first(src.get("storeNameZh"), src.get("storeDisplay"), src.get("storeName"), src.get("store"))),
        "category": category,
        "category_path": breadcrumb or category,
        "primary_cat_code": clean(src.get("primaryCatCode")),
        "country_of_origin": display_text(
            first(src.get("countryOfOriginDisplay"), src.get("countryOfOriginZh"), src.get("countryOfOrigin"))
        ),
        "packing_spec": clean(first(src.get("packingSpecZh"), src.get("packingSpecEn"))),
        "summary": clean(first(src.get("summaryZh"), src.get("summaryEn"))),
        "image": clean(image),
        "image_local": "",
        "product_url": clean(product_url),
        "keyword": clean(task.get("keyword")),
        "task_categories": ",".join(task.get("categories") or []),
        "price_range_filter": clean(task.get("price_range")),
        "brand_filter": clean(task.get("brand")),
        "source_mode": clean(task.get("mode") or task.get("method")),
        "scraped_at": now_iso(),
    }

    if use_pdp_fallback and session and code:
        missing = [k for k in ("name", "brand", "selling_price", "image") if not row.get(k)]
        if missing:
            fallback = hktv_pdp_fallback_fields(session, code, site_key=site_key)
            for key, value in fallback.items():
                mapped = {
                    "name": "name",
                    "brand": "brand",
                    "selling_price": "selling_price",
                    "average_rating": "average_rating",
                    "number_of_reviews": "number_of_reviews",
                    "summary": "summary",
                    "image": "image",
                    "category_path": "category_path",
                    "main_category": "category",
                }.get(key, key)
                if not row.get(mapped):
                    if key == "in_stock":
                        row["availability"] = normalise_availability(value)
                    elif mapped in row:
                        row[mapped] = value
    return row


# =============================================================================
# 匯出（含圖片）
# =============================================================================


def download_image_bytes(session: requests.Session, url: str, timeout: float = 20.0) -> bytes | None:
    if not url:
        return None
    try:
        resp = session.get(url, timeout=timeout, headers={"accept": "image/*"})
        resp.raise_for_status()
        return resp.content
    except Exception:
        return None


def cache_product_image(session: requests.Session, row: dict, timeout: float = 20.0) -> str:
    code = clean(row.get("product_code"))
    image_url = clean(row.get("image"))
    if not code or not image_url:
        return ""
    ensure_dir(IMAGE_CACHE_DIR)
    ext = os.path.splitext(image_url.split("?")[0])[1] or ".jpg"
    local_path = os.path.join(IMAGE_CACHE_DIR, safe_filename(code) + ext)
    if os.path.exists(local_path) and os.path.getsize(local_path) > 0:
        return local_path
    data = download_image_bytes(session, image_url, timeout=timeout)
    if not data:
        return ""
    with open(local_path, "wb") as fh:
        fh.write(data)
    return local_path


def rows_to_dataframe(rows: list[dict]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame(columns=EXCEL_COLUMNS)
    records = [row_to_display_record(row) for row in rows]
    df = pd.DataFrame(records)
    for col in EXCEL_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df[EXCEL_COLUMNS]


def sort_dataframe(df: pd.DataFrame, sort_key: str) -> pd.DataFrame:
    spec = SORT_OPTIONS.get(sort_key)
    if not spec or df.empty:
        return df
    col, ascending = spec
    if col not in df.columns:
        return df
    work = df.copy()
    if col in {"現價 (HK$)", "原價 (HK$)", "評分", "評論數", "銷量"}:
        work[col] = pd.to_numeric(work[col], errors="coerce")
    return work.sort_values(by=col, ascending=ascending, kind="mergesort", na_position="last")


def build_excel_with_images(
    df: pd.DataFrame,
    session: requests.Session,
    *,
    embed_images: bool = True,
    image_row_height: int = 80,
    timeout: float = 20.0,
    source_rows: list[dict] | None = None,
) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "products"
    header_font = Font(bold=True)
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=1):
        ws.append(row)
        if r_idx == 1:
            for cell in ws[r_idx]:
                cell.font = header_font
                cell.alignment = Alignment(vertical="center", wrap_text=True)

    image_col_idx = None
    if "_image_local" in df.columns:
        image_col_idx = df.columns.get_loc("_image_local") + 1
    elif "商品圖片" in df.columns:
        image_col_idx = df.columns.get_loc("商品圖片") + 1

    export_rows = source_rows or []
    row_lookup = {clean(r.get("product_code")): r for r in export_rows}

    for idx, record in enumerate(df.to_dict(orient="records"), start=2):
        ws.row_dimensions[idx].height = image_row_height
        for col_idx in range(1, len(df.columns) + 1):
            ws.cell(row=idx, column=col_idx).alignment = Alignment(vertical="top", wrap_text=True)
        if not embed_images or image_col_idx is None:
            continue
        code = clean(record.get("商品編號"))
        source_row = row_lookup.get(code, {})
        local_path = clean(record.get("_image_local") or source_row.get("image_local"))
        if not local_path:
            local_path = cache_product_image(session, source_row or record, timeout=timeout)
        if not local_path or not os.path.exists(local_path):
            continue
        try:
            pil_img = PILImage.open(local_path)
            pil_img.thumbnail((72, 72))
            thumb_path = local_path + ".thumb.png"
            pil_img.save(thumb_path, format="PNG")
            xl_img = XLImage(thumb_path)
            xl_img.anchor = f"{get_column_letter(image_col_idx)}{idx}"
            ws.add_image(xl_img)
        except Exception:
            continue

    for col_idx, col_name in enumerate(df.columns, start=1):
        values = [str(col_name)] + [str(v) for v in df[col_name].head(200).tolist()]
        width = min(max(len(v) for v in values) + 2, 50)
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def build_export_package(
    rows: list[dict],
    session: requests.Session,
    *,
    label: str | None = None,
    include_images: bool = True,
    embed_images_in_excel: bool = True,
    timeout: float = 20.0,
) -> tuple[bytes, str]:
    """建立 ZIP 匯出包：Excel + CSV + images/ + manifest.json"""
    df = rows_to_dataframe(rows)
    stamp = label or datetime.now().strftime("%Y%m%d_%H%M%S")
    ensure_dir(EXPORT_DIR)

    if include_images:
        for record in rows:
            local = cache_product_image(session, record, timeout=timeout)
            record["image_local"] = local
        df = rows_to_dataframe(rows)

    excel_bytes = build_excel_with_images(
        df,
        session,
        embed_images=embed_images_in_excel and include_images,
        timeout=timeout,
        source_rows=rows,
    )
    csv_bytes = df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
    manifest = {
        "app_version": APP_VERSION,
        "exported_at": now_iso(),
        "row_count": len(df),
        "columns": EXCEL_COLUMNS,
        "label": stamp,
    }

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(f"hktv_products_{stamp}.xlsx", excel_bytes)
        zf.writestr(f"hktv_products_{stamp}.csv", csv_bytes)
        zf.writestr("manifest.json", json.dumps(manifest, ensure_ascii=False, indent=2))
        if include_images:
            for record in rows:
                local_path = clean(record.get("image_local"))
                if local_path and os.path.exists(local_path):
                    arcname = os.path.join("images", os.path.basename(local_path))
                    zf.write(local_path, arcname=arcname)
    filename = f"hktv_export_{stamp}.zip"
    return zip_buffer.getvalue(), filename


# =============================================================================
# 檢查點 / 備份
# =============================================================================


def save_checkpoint(state: dict, label: str = "latest") -> str:
    ensure_dir(CHECKPOINT_DIR)
    path = os.path.join(CHECKPOINT_DIR, f"checkpoint_{label}.json")
    serializable = {
        "app_version": APP_VERSION,
        "saved_at": now_iso(),
        "rows": state.get("rows", []),
        "tasks": state.get("tasks", []),
        "task_idx": state.get("task_idx", 0),
        "seen_product_ids": list(state.get("seen_product_ids", set())),
        "seen_page_signatures": list(state.get("seen_page_signatures", set())),
        "stats": state.get("stats", {}),
        "settings": state.get("settings", {}),
        "activity_log": state.get("activity_log", [])[-200:],
        "website_key": state.get("website_key", "hktv_zh"),
    }
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(serializable, fh, ensure_ascii=False, indent=2)
    return path


def load_checkpoint_bytes(data: bytes) -> dict:
    payload = json.loads(data.decode("utf-8"))
    payload["seen_product_ids"] = set(payload.get("seen_product_ids") or [])
    payload["seen_page_signatures"] = set(payload.get("seen_page_signatures") or [])
    return payload


def save_backup(rows: list[dict], label: str | None = None) -> str | None:
    df = rows_to_dataframe(rows)
    if df.empty:
        return None
    ensure_dir(BACKUP_DIR)
    stamp = label or datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_path = os.path.join(BACKUP_DIR, f"hktv_products_{stamp}.csv")
    xlsx_path = os.path.join(BACKUP_DIR, f"hktv_products_{stamp}.xlsx")
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    session = get_session()
    with open(xlsx_path, "wb") as fh:
        fh.write(build_excel_with_images(df, session, embed_images=False, source_rows=rows))
    return csv_path


# =============================================================================
# 任務建立 / 分割
# =============================================================================


def new_task(
    *,
    method: str,
    categories: list[str] | None = None,
    keyword: str = "",
    price_range: str | None = None,
    brand: str | None = None,
    split_depth: int = 0,
    label: str = "",
) -> dict[str, Any]:
    return {
        "method": method,
        "mode": method,
        "categories": list(categories or []),
        "keyword": keyword or "",
        "price_range": price_range,
        "brand": brand,
        "split_depth": split_depth,
        "page_number": 0,
        "done": False,
        "error": None,
        "label": label or method,
        "products_collected": 0,
    }


def clone_task_for_split(task: dict, **overrides: Any) -> dict:
    new_one = deepcopy(task)
    new_one.update(overrides)
    new_one["page_number"] = 0
    new_one["done"] = False
    new_one["error"] = None
    new_one["products_collected"] = 0
    new_one["split_depth"] = int(task.get("split_depth") or 0) + 1
    return new_one


def split_oversized_task(
    session: requests.Session,
    task: dict,
    timeout: float = 30.0,
) -> list[dict]:
    if int(task.get("split_depth") or 0) >= MAX_SPLIT_DEPTH:
        return []

    parent_code = (task.get("categories") or [None])[0]
    keyword = task.get("keyword") or ""
    extra_filter = build_task_extra_filter(task)

    if parent_code:
        children = discover_child_category_codes(
            session,
            parent_code,
            timeout,
            keyword=keyword,
            extra_filter=extra_filter or None,
        )
        if children:
            return [
                clone_task_for_split(
                    task,
                    categories=[child["code"]],
                    label=f"{task.get('label', 'task')} > {child['name']}",
                )
                for child in children
            ]

    if not task.get("price_range"):
        return [
            clone_task_for_split(
                task,
                price_range=bucket,
                label=f"{task.get('label', 'task')} @ {bucket}",
            )
            for bucket in DEFAULT_PRICE_RANGE_BUCKETS
        ]

    if not task.get("brand"):
        brands = fetch_brand_facets(
            session,
            keyword=keyword,
            categories=task.get("categories"),
            extra_filter=extra_filter or None,
            timeout=timeout,
        )
        if brands:
            return [
                clone_task_for_split(
                    task,
                    brand=brand_name,
                    label=f"{task.get('label', 'task')} / {brand_name}",
                )
                for brand_name in brands
            ]
    return []


def expand_subcategory_tasks(
    session: requests.Session,
    tasks: list[dict],
    *,
    depth: int = 0,
    max_depth: int = MAX_SUBCATEGORY_DEPTH,
    timeout: float = 30.0,
) -> list[dict]:
    if depth >= max_depth:
        return tasks
    expanded: list[dict] = []
    for task in tasks:
        expanded.append(task)
        parent = (task.get("categories") or [None])[0]
        if not parent:
            continue
        children = discover_child_category_codes(session, parent, timeout)
        if not children:
            continue
        child_tasks = [
            new_task(
                method="category",
                categories=[child["code"]],
                label=f"subcategory:{child['name']}",
            )
            for child in children
        ]
        expanded.extend(
            expand_subcategory_tasks(
                session,
                child_tasks,
                depth=depth + 1,
                max_depth=max_depth,
                timeout=timeout,
            )
        )
    return expanded


def build_tasks(cfg: dict, session: requests.Session | None = None) -> list[dict]:
    method = cfg.get("method", "keyword")
    session = session or build_session()
    timeout = float(cfg.get("timeout", 30))
    tasks: list[dict] = []

    keywords = [clean(k) for k in cfg.get("keywords", []) if clean(k)]
    category_codes = [clean(c) for c in cfg.get("category_codes", []) if clean(c)]
    expand_subcats = bool(cfg.get("expand_subcategories", False))

    if method == "keyword":
        for kw in keywords:
            tasks.append(new_task(method="keyword", keyword=kw, label=f"關鍵字:{kw}"))

    elif method == "category":
        for code in category_codes:
            base = new_task(method="category", categories=[code], label=f"分類:{code}")
            if expand_subcats:
                tasks.extend(expand_subcategory_tasks(session, [base], timeout=timeout))
            else:
                tasks.append(base)

    elif method == "both":
        if not keywords and not category_codes:
            return []
        if keywords and category_codes:
            for code in category_codes:
                for kw in keywords:
                    tasks.append(
                        new_task(
                            method="both",
                            categories=[code],
                            keyword=kw,
                            label=f"分類:{code}+關鍵字:{kw}",
                        )
                    )
        elif keywords:
            for kw in keywords:
                tasks.append(new_task(method="keyword", keyword=kw, label=f"關鍵字:{kw}"))
        else:
            for code in category_codes:
                tasks.append(new_task(method="category", categories=[code], label=f"分類:{code}"))

    elif method == "all":
        top_cats = fetch_top_level_categories(session, timeout=timeout)
        if not top_cats:
            tasks.append(new_task(method="all", label="全站商品"))
        else:
            for item in top_cats:
                tasks.append(
                    new_task(
                        method="all",
                        categories=[item["code"]],
                        label=f"全站:{item['code']}",
                    )
                )
    return tasks


# =============================================================================
# 狀態管理
# =============================================================================


def init_state() -> None:
    defaults: dict[str, Any] = {
        "app_version": APP_VERSION,
        "initialized": True,
        "http_session": None,
        "website_key": "hktv_zh",
        "running": False,
        "auto_run": False,
        "tasks": [],
        "task_idx": 0,
        "rows": [],
        "seen_product_ids": set(),
        "seen_page_signatures": set(),
        "stats": {
            "pages_fetched": 0,
            "products_added": 0,
            "tasks_completed": 0,
            "tasks_split": 0,
            "errors": 0,
        },
        "settings": {},
        "last_message": "",
        "activity_log": [],
        "preview_sort": "預設（抓取順序）",
        "last_backup_count": 0,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            if isinstance(value, set):
                st.session_state[key] = set()
            elif isinstance(value, dict):
                st.session_state[key] = deepcopy(value)
            elif isinstance(value, list):
                st.session_state[key] = list(value)
            else:
                st.session_state[key] = value


def reset_run_state(keep_settings: bool = True) -> None:
    settings = deepcopy(st.session_state.get("settings", {}))
    website_key = st.session_state.get("website_key", "hktv_zh")
    preview_sort = st.session_state.get("preview_sort", "預設（抓取順序）")
    close_session()
    st.session_state.tasks = []
    st.session_state.task_idx = 0
    st.session_state.rows = []
    st.session_state.seen_product_ids = set()
    st.session_state.seen_page_signatures = set()
    st.session_state.stats = {
        "pages_fetched": 0,
        "products_added": 0,
        "tasks_completed": 0,
        "tasks_split": 0,
        "errors": 0,
    }
    st.session_state.running = False
    st.session_state.auto_run = False
    st.session_state.last_message = ""
    st.session_state.activity_log = []
    st.session_state.last_backup_count = 0
    st.session_state.http_session = None
    st.session_state.website_key = website_key
    st.session_state.preview_sort = preview_sort
    if keep_settings:
        st.session_state.settings = settings


def expand_all_product_category_tasks(cfg: dict) -> int:
    """持續發現頂層分類並追加任務（無一次性 gate）。"""
    session = get_session()
    top_cats = fetch_top_level_categories(session, timeout=float(cfg.get("timeout", 30)))
    existing = {
        tuple(task.get("categories") or [])
        for task in st.session_state.tasks
        if task.get("method") == "all"
    }
    added = 0
    insert_at = len(st.session_state.tasks)
    for item in top_cats:
        key = (item["code"],)
        if key in existing:
            continue
        st.session_state.tasks.insert(
            insert_at,
            new_task(method="all", categories=[item["code"]], label=f"全站:{item['code']}"),
        )
        existing.add(key)
        added += 1
        insert_at += 1
    if added:
        append_activity(f"新增 {added} 個頂層分類任務", "info")
    return added


# =============================================================================
# 執行引擎
# =============================================================================


def run_one_step(cfg: dict) -> dict[str, Any]:
    init_state()
    session = get_session()

    if cfg.get("method") == "all":
        expand_all_product_category_tasks(cfg)

    tasks: list[dict] = st.session_state.tasks
    if not tasks:
        return {"status": "idle", "message": "尚未建立任務。"}

    task_idx = st.session_state.task_idx
    while task_idx < len(tasks) and tasks[task_idx].get("done"):
        task_idx += 1
    st.session_state.task_idx = task_idx
    if task_idx >= len(tasks):
        st.session_state.auto_run = False
        return {"status": "complete", "message": "所有任務已完成。"}

    task = tasks[task_idx]
    page_size = int(cfg.get("page_size") or DEFAULT_PAGE_SIZE)
    page_number = int(task.get("page_number") or 0)
    cfg_max_pages = int(cfg.get("max_pages") or 0)
    product_limit = int(cfg.get("limit") or 0)
    timeout = float(cfg.get("timeout") or 30)
    keyword = task.get("keyword") or ""
    categories = list(task.get("categories") or [])
    price_range = task.get("price_range")
    brand = task.get("brand")
    extra_filter = build_task_extra_filter(task)
    use_pdp_fallback = bool(cfg.get("use_pdp_fallback", False))
    site_key = cfg.get("website_key", st.session_state.get("website_key", "hktv_zh"))

    try:
        result = hktv_fetch_page(
            session,
            keyword=keyword,
            page_number=page_number,
            page_size=page_size,
            categories=categories or None,
            extra_filter=extra_filter or None,
            sort_by=cfg.get("sort_by") or "salesVolume:desc",
            timeout=timeout,
        )
    except Exception as exc:
        task["error"] = str(exc)
        st.session_state.stats["errors"] += 1
        task["done"] = True
        st.session_state.task_idx = task_idx + 1
        append_activity(f"任務錯誤：{exc}", "error")
        return {"status": "error", "message": str(exc)}

    hits = result.get("hits") or []
    total = int(result.get("total") or 0)
    st.session_state.stats["pages_fetched"] += 1

    sig = task_page_signature(task_idx, page_number, categories, keyword, price_range, brand)
    repeated_page = sig in st.session_state.seen_page_signatures
    st.session_state.seen_page_signatures.add(sig)

    max_page = api_max_page_number(page_size)
    if cfg_max_pages > 0:
        max_page = min(max_page, cfg_max_pages - 1)
    at_window_limit = page_number >= max_page
    full_page = len(hits) >= page_size

    if repeated_page or (at_window_limit and full_page):
        split_result = try_split_task(
            session, tasks, task_idx, task,
            reason="重複頁面" if repeated_page else "分頁視窗上限",
            timeout=timeout,
        )
        if split_result:
            return split_result

    added = 0
    for hit in hits:
        product_id = clean(hit.get("id") or (hit.get("source") or {}).get("code"))
        if not product_id:
            continue
        if product_id in st.session_state.seen_product_ids:
            continue
        row = parse_hktv_hit(
            hit,
            task=task,
            site_key=site_key,
            session=session,
            use_pdp_fallback=use_pdp_fallback,
        )
        st.session_state.rows.append(row)
        st.session_state.seen_product_ids.add(product_id)
        added += 1
        task["products_collected"] = int(task.get("products_collected") or 0) + 1

    st.session_state.stats["products_added"] += added

    stop_due_to_limit = product_limit > 0 and len(st.session_state.rows) >= product_limit
    stop_due_to_max_pages = cfg_max_pages > 0 and (page_number + 1) >= cfg_max_pages
    no_more_hits = len(hits) == 0
    natural_end = total > 0 and (page_number + 1) * page_size >= total

    if stop_due_to_limit:
        st.session_state.auto_run = False
        task["done"] = True
        st.session_state.stats["tasks_completed"] += 1
        st.session_state.task_idx = task_idx + 1
        msg = f"已達商品上限 {product_limit}"
        append_activity(msg, "info")
        return {"status": "limit", "message": msg, "added": added}

    if stop_due_to_max_pages or no_more_hits or natural_end:
        if should_attempt_split_on_end(task, page_number, hits, page_size):
            split_result = try_split_task(
                session, tasks, task_idx, task,
                reason="API 分頁中斷（可能仍有未抓取商品）",
                timeout=timeout,
            )
            if split_result:
                return split_result
        task["done"] = True
        st.session_state.stats["tasks_completed"] += 1
        st.session_state.task_idx = task_idx + 1
        msg = f"任務完成：{task.get('label', task_idx)}（本頁 +{added}）"
        append_activity(msg, "info")
        return {"status": "task_done", "message": msg, "added": added}

    if at_window_limit and full_page:
        split_result = try_split_task(
            session, tasks, task_idx, task,
            reason="分頁視窗上限",
            timeout=timeout,
        )
        if split_result:
            return split_result
        task["done"] = True
        st.session_state.stats["tasks_completed"] += 1
        st.session_state.task_idx = task_idx + 1
        msg = "達分頁上限且無法再分割"
        append_activity(msg, "error")
        return {"status": "window_stop", "message": msg}

    task["page_number"] = page_number + 1
    msg = f"第 {page_number} 頁完成 +{added}（累計 {len(st.session_state.rows)}）"
    append_activity(msg, "info")
    return {
        "status": "progress",
        "message": msg,
        "added": added,
        "page_number": page_number,
        "task_label": task.get("label"),
    }


def maybe_auto_backup(cfg: dict) -> None:
    every = int(cfg.get("auto_backup_every") or 0)
    if every <= 0:
        return
    count = len(st.session_state.rows)
    if count <= 0 or count < st.session_state.last_backup_count + every:
        return
    path = save_backup(st.session_state.rows, label=f"auto_{count}")
    if path:
        st.session_state.last_backup_count = count
        append_activity(f"自動備份：{path}", "info")


# =============================================================================
# Streamlit UI（繁體中文）
# =============================================================================


def render_sidebar() -> dict:
    st.sidebar.title("HKTVmall 爬蟲")
    st.sidebar.caption(f"版本 {APP_VERSION}")
    st.sidebar.caption(
        "大型結果集會自動分割任務（子分類 → 價格區間 → 品牌），"
        "突破 API 約 10,000 offset 分頁上限。"
    )

    website_key = st.sidebar.selectbox(
        "網站",
        options=list(WEBSITES.keys()),
        format_func=lambda k: WEBSITES[k]["label"],
        index=0,
    )
    st.session_state.website_key = website_key

    method = st.sidebar.selectbox(
        "抓取方式",
        options=["keyword", "category", "both", "all"],
        format_func=lambda x: {
            "keyword": "關鍵字搜尋",
            "category": "分類代碼",
            "both": "分類 + 關鍵字",
            "all": "全站商品（按頂層分類）",
        }[x],
    )

    keywords: list[str] = []
    category_codes: list[str] = []
    expand_subcategories = False

    if method in {"keyword", "both"}:
        raw_kw = st.sidebar.text_area("關鍵字（每行一個）", value="", height=100)
        keywords = [clean(x) for x in raw_kw.splitlines() if clean(x)]

    if method in {"category", "both", "all"} and method != "all":
        raw_cat = st.sidebar.text_area("分類代碼（每行一個）", value="", height=100)
        category_codes = [clean(x) for x in raw_cat.splitlines() if clean(x)]
        expand_subcategories = st.sidebar.checkbox("展開子分類（最多 4 層）", value=False)

    page_size = st.sidebar.number_input("每頁筆數", min_value=10, max_value=120, value=DEFAULT_PAGE_SIZE, step=10)
    max_pages = st.sidebar.number_input(
        "每任務最多頁數（0 = 無限制）",
        min_value=0,
        max_value=10000,
        value=0,
        step=1,
    )
    limit = st.sidebar.number_input(
        "商品上限（0 = 無限制）",
        min_value=0,
        max_value=10_000_000,
        value=0,
        step=100,
    )
    timeout = st.sidebar.number_input("請求逾時（秒）", min_value=5, max_value=120, value=30, step=5)
    steps_per_loop = st.sidebar.number_input("每次迴圈步數", min_value=1, max_value=100, value=3, step=1)
    loop_delay = st.sidebar.number_input("迴圈間隔（秒）", min_value=0.0, max_value=5.0, value=0.2, step=0.1)
    auto_backup_every = st.sidebar.number_input("每 N 筆自動備份（0=關）", min_value=0, max_value=100000, value=0, step=500)
    use_pdp_fallback = st.sidebar.checkbox("缺少欄位時抓取商品頁補充", value=False)
    include_images = st.sidebar.checkbox("匯出時下載商品圖片", value=True)
    embed_images_in_excel = st.sidebar.checkbox("Excel 內嵌縮圖", value=True)

    sort_label = st.sidebar.selectbox(
        "搜尋排序（API）",
        options=list(API_SORT_OPTIONS.keys()),
        index=0,
        help="影響抓取時 API 回傳順序；預覽表格排序可在下方另行選擇。",
    )

    return {
        "method": method,
        "keywords": keywords,
        "category_codes": category_codes,
        "expand_subcategories": expand_subcategories,
        "page_size": int(page_size),
        "max_pages": int(max_pages),
        "limit": int(limit),
        "timeout": float(timeout),
        "steps_per_loop": int(steps_per_loop),
        "loop_delay": float(loop_delay),
        "auto_backup_every": int(auto_backup_every),
        "use_pdp_fallback": use_pdp_fallback,
        "include_images": include_images,
        "embed_images_in_excel": embed_images_in_excel,
        "website_key": website_key,
        "sort_by": API_SORT_OPTIONS[sort_label],
    }


def render_activity_log() -> None:
    with st.expander("活動紀錄", expanded=False):
        logs = list(reversed(st.session_state.get("activity_log", [])[-80:]))
        if not logs:
            st.write("尚無紀錄。")
            return
        for entry in logs:
            icon = {"info": "ℹ️", "warning": "⚠️", "error": "❌"}.get(entry.get("level", "info"), "•")
            st.text(f"{icon} [{entry.get('time', '')}] {entry.get('message', '')}")


def render_task_queue() -> None:
    with st.expander("任務佇列", expanded=False):
        tasks = st.session_state.tasks
        if not tasks:
            st.write("尚未建立任務。")
            return
        task_df = pd.DataFrame(
            [
                {
                    "#": idx,
                    "標籤": t.get("label"),
                    "方式": t.get("method"),
                    "分類": ",".join(t.get("categories") or []),
                    "關鍵字": t.get("keyword"),
                    "價格區間": t.get("price_range"),
                    "品牌": t.get("brand"),
                    "分割深度": t.get("split_depth"),
                    "頁碼": t.get("page_number"),
                    "完成": t.get("done"),
                    "商品數": t.get("products_collected"),
                    "錯誤": t.get("error"),
                }
                for idx, t in enumerate(tasks)
            ]
        )
        st.dataframe(task_df, use_container_width=True, height=280)


def render_preview(cfg: dict) -> None:
    rows = st.session_state.rows
    if not rows:
        st.info("尚無資料。請建立任務後按「開始」。")
        return

    sort_key = st.selectbox(
        "預覽排序",
        options=list(SORT_OPTIONS.keys()),
        index=list(SORT_OPTIONS.keys()).index(st.session_state.get("preview_sort", "預設（抓取順序）")),
        key="preview_sort_select",
    )
    st.session_state.preview_sort = sort_key
    df = sort_dataframe(rows_to_dataframe(rows), sort_key)
    total = len(df)

    page_size = st.selectbox(
        "預覽每頁筆數（0 = 顯示全部）",
        options=list(PREVIEW_PAGE_SIZE_OPTIONS),
        index=1,
        format_func=lambda n: "全部" if n == 0 else str(n),
        key="preview_page_size_select",
    )
    page_count = 1 if page_size == 0 else max(1, (total + page_size - 1) // page_size)
    page_no = st.number_input(
        "預覽頁碼",
        min_value=1,
        max_value=page_count,
        value=1,
        step=1,
        key="preview_page_no",
    ) if page_size > 0 else 1

    if page_size > 0:
        start = (page_no - 1) * page_size
        view_df = df.iloc[start : start + page_size]
        st.subheader("資料預覽")
        st.caption(f"共 {total:,} 筆｜顯示第 {start + 1:,}–{min(start + page_size, total):,} 筆（下載檔案包含全部 {total:,} 筆）")
    else:
        view_df = df
        st.subheader("資料預覽")
        st.caption(f"共 {total:,} 筆（已全部顯示；下載檔案同樣包含全部資料）")

    st.dataframe(
        view_df,
        use_container_width=True,
        height=min(700, max(300, 35 * min(len(view_df) + 1, 25))),
        column_config={
            "商品圖片": st.column_config.ImageColumn("商品圖片", width="small"),
            "商品連結": st.column_config.LinkColumn("商品連結", display_text="開啟商品"),
            "現價 (HK$)": st.column_config.NumberColumn("現價 (HK$)", format="HK$%d"),
            "原價 (HK$)": st.column_config.NumberColumn("原價 (HK$)", format="HK$%d"),
            "銷量": st.column_config.NumberColumn("銷量", format="%d"),
            "評分": st.column_config.NumberColumn("評分", format="%.1f"),
            "評論數": st.column_config.NumberColumn("評論數", format="%d"),
        },
    )

    session = get_session()
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.download_button(
            f"下載 CSV（全部 {total:,} 筆）",
            data=df.to_csv(index=False).encode("utf-8-sig"),
            file_name=f"hktv_products_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with c2:
        st.download_button(
            f"下載 Excel（全部 {total:,} 筆）",
            data=build_excel_with_images(
                df,
                session,
                embed_images=cfg.get("embed_images_in_excel", True) and cfg.get("include_images", True),
                timeout=cfg.get("timeout", 30),
                source_rows=rows,
            ),
            file_name=f"hktv_products_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with c3:
        zip_bytes, zip_name = build_export_package(
            st.session_state.rows,
            session,
            include_images=cfg.get("include_images", True),
            embed_images_in_excel=cfg.get("embed_images_in_excel", True),
            timeout=cfg.get("timeout", 30),
        )
        st.download_button(
            "下載完整匯出包（ZIP）",
            data=zip_bytes,
            file_name=zip_name,
            mime="application/zip",
            use_container_width=True,
        )
    with c4:
        if st.button("立即備份", use_container_width=True):
            path = save_backup(st.session_state.rows)
            if path:
                st.success(f"已備份：{path}")
            else:
                st.warning("沒有資料可備份。")


def render_main() -> None:
    init_state()
    cfg = render_sidebar()

    st.title("HKTVmall 商品爬蟲")
    st.write(
        "從 HKTVmall 搜尋 API 抓取商品資料，支援檢查點還原、去重、"
        "Excel 圖片匯出，以及自動任務分割以取得完整商品清單。"
    )

    b1, b2, b3, b4, b5 = st.columns(5)
    with b1:
        if st.button("建立任務", use_container_width=True):
            reset_run_state(keep_settings=True)
            st.session_state.settings = deepcopy(cfg)
            st.session_state.tasks = build_tasks(cfg, session=get_session())
            append_activity(f"已建立 {len(st.session_state.tasks)} 個任務", "info")
            st.success(f"已建立 {len(st.session_state.tasks)} 個任務")
    with b2:
        if st.button("開始", use_container_width=True):
            st.session_state.settings = deepcopy(cfg)
            if not st.session_state.tasks:
                st.session_state.tasks = build_tasks(cfg, session=get_session())
            st.session_state.auto_run = True
            st.session_state.running = True
            append_activity("開始自動抓取", "info")
    with b3:
        if st.button("暫停", use_container_width=True):
            st.session_state.auto_run = False
            st.session_state.running = False
            append_activity("已暫停", "warning")
    with b4:
        if st.button("儲存檢查點", use_container_width=True):
            path = save_checkpoint(
                {
                    "rows": st.session_state.rows,
                    "tasks": st.session_state.tasks,
                    "task_idx": st.session_state.task_idx,
                    "seen_product_ids": st.session_state.seen_product_ids,
                    "seen_page_signatures": st.session_state.seen_page_signatures,
                    "stats": st.session_state.stats,
                    "settings": st.session_state.settings,
                    "activity_log": st.session_state.activity_log,
                    "website_key": st.session_state.website_key,
                }
            )
            append_activity(f"檢查點已儲存：{path}", "info")
            st.success(f"檢查點已儲存：{path}")
    with b5:
        if st.button("重設", use_container_width=True):
            reset_run_state(keep_settings=False)
            st.warning("執行狀態已清除")

    uploaded = st.file_uploader("還原檢查點 JSON", type=["json"])
    if uploaded is not None:
        try:
            data = load_checkpoint_bytes(uploaded.getvalue())
            st.session_state.rows = data.get("rows", [])
            st.session_state.tasks = data.get("tasks", [])
            st.session_state.task_idx = data.get("task_idx", 0)
            st.session_state.seen_product_ids = data.get("seen_product_ids", set())
            st.session_state.seen_page_signatures = data.get("seen_page_signatures", set())
            st.session_state.stats = data.get("stats", st.session_state.stats)
            st.session_state.settings = data.get("settings", st.session_state.settings)
            st.session_state.activity_log = data.get("activity_log", [])
            st.session_state.website_key = data.get("website_key", "hktv_zh")
            append_activity("已還原檢查點", "info")
            st.success("檢查點還原成功。")
        except Exception as exc:
            st.error(f"還原失敗：{exc}")

    stats = st.session_state.stats
    m1, m2, m3, m4, m5, m6 = st.columns(6)
    m1.metric("商品數", len(st.session_state.rows))
    m2.metric("已完成任務", stats.get("tasks_completed", 0))
    m3.metric("已抓取頁數", stats.get("pages_fetched", 0))
    m4.metric("分割任務", stats.get("tasks_split", 0))
    m5.metric("錯誤", stats.get("errors", 0))
    m6.metric("目前任務", st.session_state.task_idx)
    if st.session_state.last_message:
        st.caption(st.session_state.last_message)

    if st.session_state.auto_run:
        active_cfg = st.session_state.get("settings") or cfg
        for _ in range(max(1, active_cfg.get("steps_per_loop", 1))):
            step = run_one_step(active_cfg)
            maybe_auto_backup(active_cfg)
            if step.get("status") in {"complete", "idle", "limit"}:
                st.session_state.auto_run = False
                st.session_state.running = False
                break
            if active_cfg.get("limit", 0) > 0 and len(st.session_state.rows) >= active_cfg["limit"]:
                st.session_state.auto_run = False
                st.session_state.running = False
                break
        delay = float(active_cfg.get("loop_delay", 0.2))
        if delay > 0:
            time.sleep(delay)
        st.rerun()

    render_preview(cfg)
    render_activity_log()
    render_task_queue()


def main() -> None:
    st.set_page_config(page_title="HKTVmall 商品爬蟲", page_icon="🛒", layout="wide")
    render_main()


if __name__ == "__main__":
    main()
