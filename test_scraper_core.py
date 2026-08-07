#!/usr/bin/env python3
"""Non-UI smoke tests for hktvmall_scraper core logic."""

from __future__ import annotations

import sys
from types import SimpleNamespace
from unittest.mock import MagicMock

# Stub streamlit before import
st_mock = MagicMock()
st_mock.session_state = {}


class _SessionState(dict):
    def __getattr__(self, key):
        try:
            return self[key]
        except KeyError as exc:
            raise AttributeError(key) from exc

    def __setattr__(self, key, value):
        self[key] = value


st_mock.session_state = _SessionState()
sys.modules["streamlit"] = st_mock

import hktvmall_scraper as scraper  # noqa: E402


def test_api_fetch_keyword() -> None:
    session = scraper.build_session()
    result = scraper.hktv_fetch_page(
        session,
        keyword="空氣清新機",
        page_number=0,
        page_size=10,
        sort_by="salesVolume:desc",
        timeout=30,
    )
    hits = result.get("hits") or []
    assert len(hits) > 0, "expected keyword hits"
    print(f"OK keyword fetch: {len(hits)} hits")


def test_top_level_categories() -> None:
    session = scraper.build_session()
    cats = scraper.fetch_top_level_categories(session, timeout=30)
    assert len(cats) > 5, f"expected multiple top categories, got {len(cats)}"
    print(f"OK top-level categories: {len(cats)} (sample: {cats[0]['code']})")


def test_discover_child_codes() -> None:
    session = scraper.build_session()
    children = scraper.discover_child_category_codes(session, "AA32080500000", timeout=30)
    assert len(children) > 0, "expected child category codes under supermarket"
    print(f"OK child discovery: {len(children)} codes (sample: {children[0]['code']})")


def test_price_range_filter() -> None:
    session = scraper.build_session()
    result = scraper.hktv_fetch_page(
        session,
        categories=["AA32080500000"],
        page_number=0,
        page_size=10,
        extra_filter={"priceRange": ["0-100"]},
        sort_by="salesVolume:desc",
        timeout=30,
    )
    hits = result.get("hits") or []
    assert len(hits) > 0, "expected hits with price range filter"
    print(f"OK price range filter: {len(hits)} hits")


def test_split_task_produces_subtasks() -> None:
    session = scraper.build_session()
    task = scraper.new_task(method="category", categories=["AA32080500000"], label="test")
    subs = scraper.split_oversized_task(session, task, timeout=30)
    assert len(subs) > 0, "expected split sub-tasks"
    print(f"OK split_oversized_task: {len(subs)} sub-tasks")


def test_api_max_page_number() -> None:
    assert scraper.api_max_page_number(100) == 99
    assert scraper.api_max_page_number(60) == 165
    print("OK api_max_page_number")


def main() -> int:
    tests = [
        test_api_max_page_number,
        test_api_fetch_keyword,
        test_top_level_categories,
        test_discover_child_codes,
        test_price_range_filter,
        test_split_task_produces_subtasks,
    ]
    for test in tests:
        print(f"\n--- {test.__name__} ---")
        test()
    print("\nAll core tests passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
