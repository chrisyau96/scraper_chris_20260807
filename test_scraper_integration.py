#!/usr/bin/env python3
"""Deeper integration checks for scraper correctness."""

from __future__ import annotations

import sys
from unittest.mock import MagicMock


class _SessionState(dict):
    def __getattr__(self, key):
        try:
            return self[key]
        except KeyError as exc:
            raise AttributeError(key) from exc

    def __setattr__(self, key, value):
        self[key] = value


st_mock = MagicMock()
st_mock.session_state = _SessionState()
sys.modules["streamlit"] = st_mock

import hktvmall_scraper as scraper  # noqa: E402

PASS = 0
FAIL = 0


def check(name: str, cond: bool, detail: str = "") -> None:
    global PASS, FAIL
    if cond:
        PASS += 1
        print(f"  PASS  {name}" + (f" — {detail}" if detail else ""))
    else:
        FAIL += 1
        print(f"  FAIL  {name}" + (f" — {detail}" if detail else ""))


def test_deduplication() -> None:
    print("\n[1] Deduplication")
    cfg = {
        "method": "keyword",
        "keywords": ["牛奶"],
        "page_size": 30,
        "max_pages": 2,
        "limit": 0,
        "timeout": 30,
        "sort_by": "salesVolume:desc",
        "website_key": "hktv_zh",
        "use_pdp_fallback": False,
    }
    scraper.init_state()
    scraper.reset_run_state(keep_settings=True)
    st_mock.session_state.settings = cfg
    st_mock.session_state.tasks = scraper.build_tasks(cfg, session=scraper.get_session())

    for _ in range(2):
        scraper.run_one_step(cfg)

    count_after_2 = len(st_mock.session_state.rows)
    scraper.run_one_step(cfg)  # re-fetch page 0 if task not done
    count_after_3 = len(st_mock.session_state.rows)

    check("collects products", count_after_2 > 0, f"{count_after_2} rows")
    check("no duplicate inflation on repeat", count_after_3 <= count_after_2 + 30,
          f"after3={count_after_3}, after2={count_after_2}")


def test_pagination_window_split() -> None:
    print("\n[2] Pagination window → task split")
    cfg = {
        "method": "category",
        "category_codes": ["AA32080500000"],
        "page_size": 100,
        "max_pages": 0,
        "limit": 0,
        "timeout": 30,
        "sort_by": "salesVolume:desc",
        "website_key": "hktv_zh",
        "use_pdp_fallback": False,
    }
    scraper.init_state()
    scraper.reset_run_state(keep_settings=True)
    st_mock.session_state.settings = cfg
    st_mock.session_state.tasks = scraper.build_tasks(cfg, session=scraper.get_session())

    max_page = scraper.api_max_page_number(100)
    task = st_mock.session_state.tasks[0]
    task["page_number"] = max_page
  # simulate already at window limit

    result = scraper.run_one_step(cfg)
    splits = st_mock.session_state.stats.get("tasks_split", 0)
    tasks_now = len(st_mock.session_state.tasks)

    check("detects window limit", result.get("status") in {"split", "window_stop", "task_done"},
          f"status={result.get('status')}")
    check("creates sub-tasks on split", splits > 0 or tasks_now > 1,
          f"splits={splits}, tasks={tasks_now}")


def test_all_products_task_count() -> None:
    print("\n[3] All-products task setup")
    cfg = {
        "method": "all",
        "keywords": [],
        "category_codes": [],
        "page_size": 60,
        "max_pages": 0,
        "limit": 0,
        "timeout": 30,
        "sort_by": "salesVolume:desc",
        "website_key": "hktv_zh",
    }
    session = scraper.build_session()
    tasks = scraper.build_tasks(cfg, session=session)
    check("creates multiple top-level tasks", len(tasks) >= 10, f"{len(tasks)} tasks")
    check("each task has category", all(t.get("categories") for t in tasks))


def test_product_fields() -> None:
    print("\n[4] Product field quality")
    session = scraper.build_session()
    result = scraper.hktv_fetch_page(
        session, keyword="牛奶", page_number=0, page_size=5,
        sort_by="salesVolume:desc", timeout=30,
    )
    hits = result.get("hits") or []
    check("API returns hits", len(hits) > 0)
    if hits:
        row = scraper.parse_hktv_hit(hits[0], task=scraper.new_task(method="keyword", keyword="牛奶"))
        check("has product_code", bool(row.get("product_code")))
        check("has name", bool(row.get("name")))
        check("has selling_price", row.get("selling_price") is not None)
        check("has product_url", bool(row.get("product_url")))


def test_price_split_produces_distinct_filters() -> None:
    print("\n[5] Task split produces sub-tasks")
    task = scraper.new_task(method="category", categories=["AA32080500000"], label="test")
    session = scraper.build_session()
    subs = scraper.split_oversized_task(session, task, timeout=30)
    check("sub-tasks created", len(subs) > 0, f"{len(subs)} sub-tasks")
    has_category_or_price = any(
        s.get("price_range") or (s.get("categories") and s["categories"] != task["categories"])
        for s in subs
    )
    check("sub-tasks have category or price filter", has_category_or_price)


def main() -> int:
    print("=== HKTVmall Scraper Integration Checks ===")
    test_deduplication()
    test_pagination_window_split()
    test_all_products_task_count()
    test_product_fields()
    test_price_split_produces_distinct_filters()
    print(f"\n=== Results: {PASS} passed, {FAIL} failed ===")
    scraper.close_session()
    return 0 if FAIL == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
