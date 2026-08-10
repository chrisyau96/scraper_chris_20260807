#!/usr/bin/env python3
"""Verify category URL/code parsing and auto-breakdown for user-provided examples."""

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

CATEGORY_URLS = [
    "https://www.hktvmall.com/hktv/zh/pets",
    "https://www.hktvmall.com/hktv/zh/mothernbaby",
    "https://www.hktvmall.com/hktv/zh/homenfamily",
    "https://www.hktvmall.com/hktv/zh/gadgetsandelectronics",
]

CATEGORY_CODES = [
    "AA11850000000",
    "AA11800000000",
]

EXPECTED_SLUGS = ["pets", "mothernbaby", "homenfamily", "gadgetsandelectronics"]

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


def test_url_slug_extraction() -> None:
    print("\n[1] Category URL slug extraction")
    for url, expected in zip(CATEGORY_URLS, EXPECTED_SLUGS):
        slug = scraper.extract_category_slug_from_url(url)
        check(f"slug from {expected}", slug == expected, f"got={slug}")


def test_aa_code_extraction() -> None:
    print("\n[2] AA category code extraction")
    for code in CATEGORY_CODES:
        parsed = scraper.extract_aa_category_code(code)
        check(f"code {code}", parsed == code, f"got={parsed}")
    check("slug URL is not AA code", scraper.extract_aa_category_code(CATEGORY_URLS[0]) == "")


def test_api_fetch_roots() -> None:
    print("\n[3] API accepts roots and returns products")
    session = scraper.build_session()
    for slug in EXPECTED_SLUGS:
        result = scraper.hktv_fetch_page(
            session, categories=[slug], page_number=0, page_size=3, timeout=30,
        )
        total = int(result.get("total") or 0)
        hits = len(result.get("hits") or [])
        check(f"fetch slug {slug}", total > 0 and hits > 0, f"total={total}, hits={hits}")
    for code in CATEGORY_CODES:
        result = scraper.hktv_fetch_page(
            session, categories=[code], page_number=0, page_size=3, timeout=30,
        )
        total = int(result.get("total") or 0)
        hits = len(result.get("hits") or [])
        check(f"fetch code {code}", total > 0 and hits > 0, f"total={total}, hits={hits}")


def test_auto_breakdown_tasks() -> None:
    print("\n[4] Auto-breakdown creates sub-category tasks")
    session = scraper.build_session()

    url_cfg = {
        "method": "category",
        "category_urls": CATEGORY_URLS,
        "category_codes": [],
        "auto_breakdown_categories": True,
        "timeout": 30,
    }
    url_tasks = scraper.build_tasks(url_cfg, session=session)
    check("4 URL roots produce tasks", len(url_tasks) > len(CATEGORY_URLS), f"{len(url_tasks)} tasks")
    check("URL tasks use AA sub-codes", all(
        (t.get("categories") or [""])[0].startswith("AA")
        for t in url_tasks[:20]
    ))

    code_cfg = {
        "method": "category",
        "category_urls": [],
        "category_codes": CATEGORY_CODES,
        "auto_breakdown_categories": True,
        "timeout": 30,
    }
    code_tasks = scraper.build_tasks(code_cfg, session=session)
    check("2 AA roots produce tasks", len(code_tasks) > len(CATEGORY_CODES), f"{len(code_tasks)} tasks")
    check("code tasks are scrapeable", all(t.get("categories") for t in code_tasks[:10]))


def test_locale_from_category_url() -> None:
    print("\n[5] Product language from category URL")
    cfg = {
        "category_urls": ["https://www.hktvmall.com/hktv/zh/mothernbaby"],
        "category_codes": [],
    }
    check("zh URL → hktv_zh", scraper.resolve_website_key(cfg) == "hktv_zh")
    cfg_en = {
        "category_urls": ["https://www.hktvmall.com/hktv/en/mothernbaby"],
        "category_codes": [],
    }
    check("en URL → hktv_en", scraper.resolve_website_key(cfg_en) == "hktv_en")


def main() -> int:
    print("=== Category URL/Code Breakdown Tests ===")
    test_url_slug_extraction()
    test_aa_code_extraction()
    test_api_fetch_roots()
    test_auto_breakdown_tasks()
    test_locale_from_category_url()
    print(f"\n=== Results: {PASS} passed, {FAIL} failed ===")
    scraper.close_session()
    return 0 if FAIL == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
