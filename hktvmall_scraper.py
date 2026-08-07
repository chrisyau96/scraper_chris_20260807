#!/usr/bin/env python3
"""
HKTVmall product scraper – Streamlit web app.
Unlimited pagination via automatic task splitting (category / price / brand).
"""

from __future__ import annotations

import hashlib
import io
import json
import os
import re
import time
from copy import deepcopy
from datetime import datetime, timezone
from typing import Any

import pandas as pd
import requests
import streamlit as st
from openpyxl.utils import get_column_letter
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

APP_VERSION = "1.4-webapp-unlimited"
HKTV_SEARCH_URL = "https://keyword-search-server.hktvmall.com/api/search"
HKTV_API_KEY = "0e6c95ec-4c8b-4f71-8855-11eeafe74966"
HKTV_INDEX_NAME = "hktvproduct"

API_MAX_OFFSET = 10000
DEFAULT_PAGE_SIZE = 60
DEFAULT_PRICE_RANGE_BUCKETS = [
    "0-50",
    "50-100",
    "100-200",
    "200-500",
    "500-1000",
    "1000-3000",
    "3000-",
]
MAX_SPLIT_DEPTH = 6
MAX_SUBCATEGORY_DEPTH = 4
MAX_BRAND_SPLITS = 20
CHECKPOINT_DIR = "hktv_checkpoints"
BACKUP_DIR = "hktv_backups"

PRODUCT_URL_BASE = "https://www.hktvmall.com/hktv/zh/main/p/"

# ---------------------------------------------------------------------------
# Generic helpers (unchanged utilities)
# ---------------------------------------------------------------------------


def clean(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    text = str(value).strip()
    return re.sub(r"\s+", " ", text)


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


def display_fragments(parts: list[str]) -> str:
    return " / ".join(clean(p) for p in parts if clean(p))


def now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def ensure_dir(path: str) -> None:
    os.makedirs(path, exist_ok=True)


def make_http_session() -> requests.Session:
    session = requests.Session()
    retry = Retry(
        total=3,
        connect=3,
        read=3,
        backoff_factor=0.6,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=("GET", "POST"),
    )
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session.headers.update(
        {
            "accept": "application/json",
            "Content-Type": "application/json",
            "User-Agent": "Mozilla/5.0 (compatible; HKTVmallScraper/1.4)",
            "Authorization": f"ApiKey {HKTV_API_KEY}",
        }
    )
    return session


def get_http_session() -> requests.Session:
    if st.session_state.http_session is None:
        st.session_state.http_session = make_http_session()
    return st.session_state.http_session


def close_http_session() -> None:
    session = st.session_state.get("http_session")
    if session is not None:
        try:
            session.close()
        except Exception:
            pass
    st.session_state.http_session = None


def api_max_page_number(page_size: int) -> int:
    if page_size <= 0:
        return 0
    return max(0, (API_MAX_OFFSET // page_size) - 1)


def is_top_level_category_code(code: str) -> bool:
    if not code:
        return False
    if not code.startswith("AA"):
        return True
    return len(code) == 14 and code.endswith("0000000")


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


def build_filter_from_task(task: dict) -> dict:
    filt: dict[str, Any] = {}
    categories = task.get("categories") or []
    if categories:
        filt["category"] = list(categories)
    if task.get("price_range"):
        filt["priceRange"] = [task["price_range"]]
    if task.get("brand"):
        filt["brand"] = [task["brand"]]
    return filt


def page_signature(
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


# ---------------------------------------------------------------------------
# API layer
# ---------------------------------------------------------------------------


def hktv_search_request(
    session: requests.Session,
    *,
    keyword: str = "",
    page_number: int = 0,
    page_size: int = DEFAULT_PAGE_SIZE,
    filter_obj: dict | None = None,
    aggregations: list[str] | None = None,
    extra_filter: dict | None = None,
    timeout: float = 30.0,
) -> dict[str, Any]:
    """POST to HKTVmall keyword search API and return the first result block."""
    merged_filter = merge_filters(filter_obj, extra_filter)

    # priceRange in filter is converted to the API range field
    price_ranges = merged_filter.pop("priceRange", None)
    range_obj: dict[str, str] = {}
    if price_ranges:
        bucket = price_ranges[0] if isinstance(price_ranges, list) else str(price_ranges)
        if bucket:
            range_obj["sellingPrice"] = bucket

    request_body: dict[str, Any] = {
        "indexName": HKTV_INDEX_NAME,
        "keyword": keyword or "",
        "page": {"pageNumber": page_number, "pageSize": page_size},
        "filter": merged_filter,
        "aggregations": aggregations or ["category", "brand"],
    }
    if range_obj:
        request_body["range"] = range_obj

    response = session.post(
        HKTV_SEARCH_URL,
        json={"requests": [request_body]},
        timeout=timeout,
    )
    response.raise_for_status()
    payload = response.json()
    if payload.get("code") != 200:
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
        timeout=timeout,
    )


def fetch_top_level_categories(
    session: requests.Session,
    timeout: float = 30.0,
) -> list[dict[str, str]]:
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
    items: list[dict[str, str]] = []
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
    """Discover child category codes under parent_code from product hits and samples."""
    discovered: dict[str, str] = {}

    def collect_from_hits(hits: list[dict]) -> None:
        for hit in hits:
            src = hit.get("source") or hit
            code = clean(src.get("primaryCatCode"))
            if not code or code == parent_code:
                continue
            cat_display = src.get("categoryStructureDisplay")
            if isinstance(cat_display, list):
                name = clean(cat_display[-1]) if cat_display else code
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

    brand_agg = first_page.get("aggregations", {}).get("brand") or {}
    top_brands = []
    if isinstance(brand_agg, dict):
        top_brands = [
            b
            for b, _ in sorted(brand_agg.items(), key=lambda kv: kv[1], reverse=True)
            if b and b not in ("OtherBrands", "ShippedfromMainland")
        ][:8]

    for bucket in DEFAULT_PRICE_RANGE_BUCKETS:
        bucket_filter = merge_filters(base_filter, {"priceRange": [bucket]})
        sampled = hktv_search_request(
            session,
            keyword=keyword,
            page_number=0,
            page_size=DEFAULT_PAGE_SIZE,
            filter_obj=bucket_filter,
            aggregations=["category"],
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
            aggregations=["category"],
            timeout=timeout,
        )
        collect_from_hits(sampled.get("hits") or [])

    children: list[dict[str, str]] = []
    for code, name in discovered.items():
        if code == parent_code:
            continue
        children.append({"code": code, "name": name})
    children.sort(key=lambda item: item["code"])
    return children


def fetch_subcategories_for_code(
    session: requests.Session,
    parent_code: str,
    timeout: float = 30.0,
) -> list[dict[str, str]]:
    """Compatibility wrapper – aggregations.category is {code: count}, not facets."""
    return discover_child_category_codes(session, parent_code, timeout)


# ---------------------------------------------------------------------------
# Product parsing & export
# ---------------------------------------------------------------------------


def hit_to_row(hit: dict, *, task: dict) -> dict[str, Any]:
    src = hit.get("source") or hit
    code = clean(src.get("code") or src.get("productSearchCode"))
    name = first(src.get("nameZh"), src.get("nameEn"), src.get("nameZhCN"))
    brand = first(src.get("brandDisplay"), src.get("brandZh"), src.get("brand"), src.get("brandEn"))
    image = first(
        src.get("imageLink"),
        (src.get("images") or [None])[0] if isinstance(src.get("images"), list) else None,
        (src.get("gallery") or [None])[0] if isinstance(src.get("gallery"), list) else None,
    )
    category_path = display_fragments(
        [
            first(src.get("mainCatNameZh"), src.get("mainCatNameEn")),
            src.get("subCat1NameZh"),
            src.get("subCat2NameZh"),
            src.get("subCat3NameZh"),
            src.get("subCat4NameZh"),
        ]
    )
    product_url = first(src.get("urlZh"), src.get("urlEn"))
    if not product_url and code:
        product_url = PRODUCT_URL_BASE + code

    return {
        "product_code": code,
        "base_product": clean(src.get("baseProduct")),
        "name": clean(name),
        "brand": clean(brand),
        "selling_price": src.get("sellingPrice"),
        "selling_price_range": clean(src.get("sellingPriceRange")),
        "saved_price": src.get("savedPrice"),
        "average_rating": src.get("averageRating"),
        "number_of_reviews": src.get("numberOfReviews"),
        "in_stock": src.get("hasStock"),
        "stock": clean(src.get("stock")),
        "loyalty_point": src.get("loyaltyPoint"),
        "number_of_variants": src.get("numberOfVariants"),
        "main_category": clean(first(src.get("mainCatNameZh"), src.get("mainCatNameEn"))),
        "sub_category": clean(first(src.get("subCat1NameZh"), src.get("subCat1NameEn"))),
        "category_path": category_path,
        "primary_cat_code": clean(src.get("primaryCatCode")),
        "image": clean(image),
        "product_url": clean(product_url),
        "store": clean(first(src.get("storeNameZh"), src.get("storeDisplay"), src.get("store"))),
        "keyword": clean(task.get("keyword")),
        "task_categories": ",".join(task.get("categories") or []),
        "price_range_filter": clean(task.get("price_range")),
        "brand_filter": clean(task.get("brand")),
        "source_mode": clean(task.get("mode")),
        "scraped_at": now_iso(),
    }


def rows_to_dataframe(rows: list[dict]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame()
    return pd.DataFrame(rows)


def autosize_excel_columns(writer: pd.ExcelWriter, sheet_name: str, df: pd.DataFrame) -> None:
    worksheet = writer.sheets[sheet_name]
    for idx, col in enumerate(df.columns, start=1):
        max_len = max([len(str(col))] + [len(str(v)) for v in df[col].head(200).tolist()])
        worksheet.column_dimensions[get_column_letter(idx)].width = min(max_len + 2, 60)


def dataframe_to_excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="products")
        autosize_excel_columns(writer, "products", df)
    return buffer.getvalue()


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
    }
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(serializable, fh, ensure_ascii=False, indent=2)
    return path


def save_backup(df: pd.DataFrame, label: str | None = None) -> str | None:
    if df.empty:
        return None
    ensure_dir(BACKUP_DIR)
    stamp = label or datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_path = os.path.join(BACKUP_DIR, f"hktv_products_{stamp}.csv")
    xlsx_path = os.path.join(BACKUP_DIR, f"hktv_products_{stamp}.xlsx")
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    with open(xlsx_path, "wb") as fh:
        fh.write(dataframe_to_excel_bytes(df))
    return csv_path


def load_checkpoint(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as fh:
        data = json.load(fh)
    data["seen_product_ids"] = set(data.get("seen_product_ids") or [])
    data["seen_page_signatures"] = set(data.get("seen_page_signatures") or [])
    return data


# ---------------------------------------------------------------------------
# Task model & splitting
# ---------------------------------------------------------------------------


def new_task(
    *,
    mode: str,
    categories: list[str] | None = None,
    keyword: str = "",
    price_range: str | None = None,
    brand: str | None = None,
    split_depth: int = 0,
    label: str = "",
) -> dict[str, Any]:
    return {
        "mode": mode,
        "categories": list(categories or []),
        "keyword": keyword or "",
        "price_range": price_range,
        "brand": brand,
        "split_depth": split_depth,
        "page_number": 0,
        "done": False,
        "error": None,
        "label": label or mode,
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
    """Split a task that hit the pagination window into finer sub-tasks."""
    if int(task.get("split_depth") or 0) >= MAX_SPLIT_DEPTH:
        return []

    parent_code = (task.get("categories") or [None])[0]
    keyword = task.get("keyword") or ""
    extra_filter: dict[str, Any] = {}
    if task.get("price_range"):
        extra_filter["priceRange"] = [task["price_range"]]
    if task.get("brand"):
        extra_filter["brand"] = [task["brand"]]

    # (a) child category codes from products
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

    # (b) price range buckets
    if not task.get("price_range"):
        return [
            clone_task_for_split(
                task,
                price_range=bucket,
                label=f"{task.get('label', 'task')} @ {bucket}",
            )
            for bucket in DEFAULT_PRICE_RANGE_BUCKETS
        ]

    # (c) brand facets
    if not task.get("brand"):
        probe_filter = merge_filters(build_filter_from_task(task), None)
        probe = hktv_search_request(
            session,
            keyword=keyword,
            page_number=0,
            page_size=1,
            filter_obj=probe_filter,
            aggregations=["brand"],
            timeout=timeout,
        )
        brand_agg = probe.get("aggregations", {}).get("brand") or {}
        brands = []
        if isinstance(brand_agg, dict):
            brands = [
                b
                for b, _ in sorted(brand_agg.items(), key=lambda kv: kv[1], reverse=True)
                if b and b not in ("OtherBrands", "ShippedfromMainland")
            ][:MAX_BRAND_SPLITS]
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
                mode="subcategory",
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


def build_initial_tasks(settings: dict) -> list[dict]:
    mode = settings["mode"]
    if mode == "keyword":
        keywords = settings.get("keywords") or []
        return [new_task(mode="keyword", keyword=kw, label=f"keyword:{kw}") for kw in keywords]
    if mode == "category":
        codes = settings.get("category_codes") or []
        return [
            new_task(mode="category", categories=[code], label=f"category:{code}")
            for code in codes
        ]
    if mode == "subcategory":
        codes = settings.get("category_codes") or []
        base = [
            new_task(mode="subcategory", categories=[code], label=f"subcategory:{code}")
            for code in codes
        ]
        session = get_http_session()
        return expand_subcategory_tasks(session, base, timeout=settings.get("timeout", 30))
    if mode == "all_products":
        session = get_http_session()
        top_cats = fetch_top_level_categories(session, timeout=settings.get("timeout", 30))
        if not top_cats:
            return [new_task(mode="all_products", label="all_products")]
        return [
            new_task(
                mode="all_products",
                categories=[item["code"]],
                label=f"all_products:{item['code']}",
            )
            for item in top_cats
        ]
    return []


# ---------------------------------------------------------------------------
# Scraper step engine
# ---------------------------------------------------------------------------


def init_state() -> None:
    defaults = {
        "app_version": APP_VERSION,
        "initialized": True,
        "http_session": None,
        "running": False,
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
        "expanded_all_product_categories": False,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            if isinstance(value, set):
                st.session_state[key] = set()
            elif isinstance(value, dict):
                st.session_state[key] = deepcopy(value)
            else:
                st.session_state[key] = value


def reset_run_state(keep_settings: bool = True) -> None:
    settings = deepcopy(st.session_state.get("settings", {}))
    close_http_session()
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
    st.session_state.last_message = ""
    st.session_state.expanded_all_product_categories = False
    st.session_state.http_session = None
    if keep_settings:
        st.session_state.settings = settings


def expand_all_product_category_tasks(settings: dict) -> int:
    """Continuously discover top-level categories and append missing tasks."""
    session = get_http_session()
    top_cats = fetch_top_level_categories(session, timeout=settings.get("timeout", 30))
    existing = {
        tuple(task.get("categories") or [])
        for task in st.session_state.tasks
        if task.get("mode") == "all_products"
    }
    added = 0
    insert_at = len(st.session_state.tasks)
    for item in top_cats:
        key = (item["code"],)
        if key in existing:
            continue
        st.session_state.tasks.insert(
            insert_at,
            new_task(
                mode="all_products",
                categories=[item["code"]],
                label=f"all_products:{item['code']}",
            ),
        )
        existing.add(key)
        added += 1
        insert_at += 1
    return added


def run_one_step(settings: dict) -> dict[str, Any]:
    """Fetch a single page for the current task. Returns a status dict."""
    init_state()
    session = get_http_session()

    if settings.get("mode") == "all_products":
        expand_all_product_category_tasks(settings)

    tasks: list[dict] = st.session_state.tasks
    if not tasks:
        return {"status": "idle", "message": "No tasks queued."}

    task_idx = st.session_state.task_idx
    while task_idx < len(tasks) and tasks[task_idx].get("done"):
        task_idx += 1
    st.session_state.task_idx = task_idx
    if task_idx >= len(tasks):
        return {"status": "complete", "message": "All tasks finished."}

    task = tasks[task_idx]
    page_size = int(settings.get("page_size") or DEFAULT_PAGE_SIZE)
    page_number = int(task.get("page_number") or 0)
    max_pages = int(settings.get("max_pages") or 0)
    product_limit = int(settings.get("product_limit") or 0)
    timeout = float(settings.get("timeout") or 30)
    keyword = task.get("keyword") or ""
    categories = list(task.get("categories") or [])
    price_range = task.get("price_range")
    brand = task.get("brand")

    extra_filter: dict[str, Any] = {}
    if price_range:
        extra_filter["priceRange"] = [price_range]
    if brand:
        extra_filter["brand"] = [brand]

    try:
        result = hktv_fetch_page(
            session,
            keyword=keyword,
            page_number=page_number,
            page_size=page_size,
            categories=categories or None,
            extra_filter=extra_filter or None,
            timeout=timeout,
        )
    except Exception as exc:
        task["error"] = str(exc)
        st.session_state.stats["errors"] += 1
        task["done"] = True
        st.session_state.task_idx = task_idx + 1
        return {"status": "error", "message": str(exc)}

    hits = result.get("hits") or []
    total = int(result.get("total") or 0)
    st.session_state.stats["pages_fetched"] += 1

    sig = page_signature(task_idx, page_number, categories, keyword, price_range, brand)
    repeated_page = sig in st.session_state.seen_page_signatures
    st.session_state.seen_page_signatures.add(sig)

    max_page_idx = api_max_page_number(page_size)
    at_window_limit = page_number >= max_page_idx
    full_page = len(hits) >= page_size

    if repeated_page or (at_window_limit and full_page):
        sub_tasks = split_oversized_task(session, task, timeout=timeout)
        if sub_tasks:
            tasks[task_idx + 1 : task_idx + 1] = sub_tasks
            st.session_state.stats["tasks_split"] += len(sub_tasks)
            task["done"] = True
            st.session_state.stats["tasks_completed"] += 1
            st.session_state.task_idx = task_idx + 1
            reason = "repeated page" if repeated_page else "pagination window"
            msg = f"Split task ({reason}) into {len(sub_tasks)} sub-tasks."
            st.session_state.last_message = msg
            return {"status": "split", "message": msg, "added_tasks": len(sub_tasks)}

    added = 0
    for hit in hits:
        product_id = clean(hit.get("id") or (hit.get("source") or {}).get("code"))
        if not product_id:
            continue
        if product_id in st.session_state.seen_product_ids:
            continue
        row = hit_to_row(hit, task=task)
        st.session_state.rows.append(row)
        st.session_state.seen_product_ids.add(product_id)
        added += 1
        task["products_collected"] = int(task.get("products_collected") or 0) + 1

    st.session_state.stats["products_added"] += added

    stop_due_to_limit = product_limit > 0 and len(st.session_state.rows) >= product_limit
    stop_due_to_max_pages = max_pages > 0 and (page_number + 1) >= max_pages
    no_more_hits = len(hits) == 0
    natural_end = (page_number + 1) * page_size >= total

    if stop_due_to_limit or stop_due_to_max_pages or no_more_hits or natural_end:
        task["done"] = True
        st.session_state.stats["tasks_completed"] += 1
        st.session_state.task_idx = task_idx + 1
        msg = f"Task done ({task.get('label', task_idx)}); +{added} products on last page."
        st.session_state.last_message = msg
        return {
            "status": "task_done",
            "message": msg,
            "added": added,
            "total_rows": len(st.session_state.rows),
        }

    if at_window_limit and full_page:
        sub_tasks = split_oversized_task(session, task, timeout=timeout)
        if sub_tasks:
            tasks[task_idx + 1 : task_idx + 1] = sub_tasks
            st.session_state.stats["tasks_split"] += len(sub_tasks)
            task["done"] = True
            st.session_state.stats["tasks_completed"] += 1
            st.session_state.task_idx = task_idx + 1
            msg = f"Reached pagination window; split into {len(sub_tasks)} sub-tasks."
            st.session_state.last_message = msg
            return {"status": "split", "message": msg, "added_tasks": len(sub_tasks)}
        task["done"] = True
        st.session_state.stats["tasks_completed"] += 1
        st.session_state.task_idx = task_idx + 1
        msg = "Reached pagination window and cannot split further."
        st.session_state.last_message = msg
        return {"status": "window_stop", "message": msg}

    task["page_number"] = page_number + 1
    msg = f"Page {page_number} fetched (+{added}); total rows={len(st.session_state.rows)}"
    st.session_state.last_message = msg
    return {
        "status": "progress",
        "message": msg,
        "added": added,
        "page_number": page_number,
        "task_label": task.get("label"),
    }


# ---------------------------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------------------------


def render_sidebar() -> dict:
    st.sidebar.header("HKTVmall Scraper")
    st.sidebar.caption(f"Version {APP_VERSION}")
    st.sidebar.caption(
        "Large result sets are fetched via automatic task splitting "
        "(sub-categories, price buckets, brands) to bypass the ~10,000 offset API cap."
    )

    mode = st.sidebar.selectbox(
        "Scrape method",
        options=["keyword", "category", "subcategory", "all_products"],
        format_func=lambda x: {
            "keyword": "Keyword search",
            "category": "Category code(s)",
            "subcategory": "Expand subcategories",
            "all_products": "All products (by top-level category)",
        }[x],
    )

    keywords: list[str] = []
    category_codes: list[str] = []
    if mode == "keyword":
        raw_keywords = st.sidebar.text_area("Keywords (one per line)", value="")
        keywords = [clean(line) for line in raw_keywords.splitlines() if clean(line)]
    else:
        raw_codes = st.sidebar.text_area("Category code(s), one per line", value="")
        category_codes = [clean(line) for line in raw_codes.splitlines() if clean(line)]

    page_size = st.sidebar.number_input("Page size", min_value=10, max_value=120, value=DEFAULT_PAGE_SIZE, step=10)
    max_pages = st.sidebar.number_input(
        "Max pages per task (0 = unlimited)",
        min_value=0,
        max_value=10000,
        value=0,
        step=1,
    )
    product_limit = st.sidebar.number_input(
        "Product limit (0 = unlimited)",
        min_value=0,
        max_value=10_000_000,
        value=0,
        step=100,
    )
    timeout = st.sidebar.number_input("Request timeout (seconds)", min_value=5, max_value=120, value=30, step=5)
    steps_per_click = st.sidebar.number_input("Steps per Run click", min_value=1, max_value=50, value=1, step=1)
    auto_backup_every = st.sidebar.number_input("Auto-backup every N products (0=off)", min_value=0, max_value=100000, value=0, step=500)

    return {
        "mode": mode,
        "keywords": keywords,
        "category_codes": category_codes,
        "page_size": int(page_size),
        "max_pages": int(max_pages),
        "product_limit": int(product_limit),
        "timeout": float(timeout),
        "steps_per_click": int(steps_per_click),
        "auto_backup_every": int(auto_backup_every),
    }


def render_main() -> None:
    init_state()
    settings = render_sidebar()

    st.title("HKTVmall Product Scraper")
    st.write(
        "Scrape HKTVmall search results with checkpointing, deduplication, and unlimited "
        "coverage via automatic task splitting when the API pagination window is reached."
    )

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        if st.button("Prepare tasks", use_container_width=True):
            reset_run_state(keep_settings=True)
            st.session_state.settings = deepcopy(settings)
            st.session_state.tasks = build_initial_tasks(settings)
            st.session_state.running = False
            st.success(f"Prepared {len(st.session_state.tasks)} task(s).")
    with col2:
        if st.button("Run step(s)", use_container_width=True):
            st.session_state.settings = deepcopy(settings)
            if not st.session_state.tasks:
                st.session_state.tasks = build_initial_tasks(settings)
            st.session_state.running = True
            results = []
            for _ in range(max(1, settings["steps_per_click"])):
                step = run_one_step(st.session_state.settings)
                results.append(step)
                if step.get("status") in {"complete", "idle"}:
                    break
                if settings["product_limit"] > 0 and len(st.session_state.rows) >= settings["product_limit"]:
                    break
            st.session_state.running = False
            st.info(results[-1].get("message", "Step finished."))
    with col3:
        if st.button("Save checkpoint", use_container_width=True):
            path = save_checkpoint(
                {
                    "rows": st.session_state.rows,
                    "tasks": st.session_state.tasks,
                    "task_idx": st.session_state.task_idx,
                    "seen_product_ids": st.session_state.seen_product_ids,
                    "seen_page_signatures": st.session_state.seen_page_signatures,
                    "stats": st.session_state.stats,
                    "settings": st.session_state.settings,
                }
            )
            st.success(f"Checkpoint saved: {path}")
    with col4:
        if st.button("Reset", use_container_width=True):
            reset_run_state(keep_settings=False)
            st.warning("Run state cleared.")

    uploaded = st.file_uploader("Restore checkpoint JSON", type=["json"])
    if uploaded is not None:
        try:
            data = json.loads(uploaded.getvalue().decode("utf-8"))
            st.session_state.rows = data.get("rows", [])
            st.session_state.tasks = data.get("tasks", [])
            st.session_state.task_idx = data.get("task_idx", 0)
            st.session_state.seen_product_ids = set(data.get("seen_product_ids") or [])
            st.session_state.seen_page_signatures = set(data.get("seen_page_signatures") or [])
            st.session_state.stats = data.get("stats", st.session_state.stats)
            st.session_state.settings = data.get("settings", st.session_state.settings)
            st.success("Checkpoint restored.")
        except Exception as exc:
            st.error(f"Failed to restore checkpoint: {exc}")

    stats = st.session_state.stats
    st.subheader("Run status")
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Products", len(st.session_state.rows))
    m2.metric("Tasks done", stats.get("tasks_completed", 0))
    m3.metric("Pages fetched", stats.get("pages_fetched", 0))
    m4.metric("Tasks split", stats.get("tasks_split", 0))
    m5.metric("Errors", stats.get("errors", 0))
    if st.session_state.last_message:
        st.caption(st.session_state.last_message)

    df = rows_to_dataframe(st.session_state.rows)
    if not df.empty:
        st.subheader("Preview")
        st.dataframe(df.head(200), use_container_width=True)

        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                "Download CSV",
                data=df.to_csv(index=False).encode("utf-8-sig"),
                file_name=f"hktv_products_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
                use_container_width=True,
            )
        with c2:
            st.download_button(
                "Download Excel",
                data=dataframe_to_excel_bytes(df),
                file_name=f"hktv_products_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        if settings.get("auto_backup_every", 0) > 0:
            every = settings["auto_backup_every"]
            if len(df) % every == 0:
                path = save_backup(df, label=f"auto_{len(df)}")
                if path:
                    st.caption(f"Auto-backup saved: {path}")

    with st.expander("Task queue", expanded=False):
        if st.session_state.tasks:
            task_df = pd.DataFrame(
                [
                    {
                        "idx": idx,
                        "label": t.get("label"),
                        "mode": t.get("mode"),
                        "categories": ",".join(t.get("categories") or []),
                        "keyword": t.get("keyword"),
                        "price_range": t.get("price_range"),
                        "brand": t.get("brand"),
                        "split_depth": t.get("split_depth"),
                        "page_number": t.get("page_number"),
                        "done": t.get("done"),
                        "products": t.get("products_collected"),
                        "error": t.get("error"),
                    }
                    for idx, t in enumerate(st.session_state.tasks)
                ]
            )
            st.dataframe(task_df, use_container_width=True)
        else:
            st.write("No tasks yet. Click **Prepare tasks**.")


def main() -> None:
    st.set_page_config(page_title="HKTVmall Scraper", layout="wide")
    render_main()


if __name__ == "__main__":
    main()
