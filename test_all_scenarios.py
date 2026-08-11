#!/usr/bin/env python3
"""End-to-end scenario checks before handoff."""

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
st_mock.cache_data = lambda **kwargs: (lambda fn: fn)
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


def scenario_version_and_columns() -> None:
    print("\n[1] Version and export columns")
    check("app version", scraper.APP_VERSION.startswith("1.6."))
    check("no Pack Size column", "Pack Size" not in scraper.EXCEL_COLUMNS)
    check("has Product Image column", "Product Image" in scraper.EXCEL_COLUMNS)


def scenario_category_urls() -> None:
    print("\n[2] Category URL slugs (4 examples)")
    urls = [
        ("https://www.hktvmall.com/hktv/zh/pets", "pets"),
        ("https://www.hktvmall.com/hktv/zh/mothernbaby", "mothernbaby"),
        ("https://www.hktvmall.com/hktv/zh/homenfamily", "homenfamily"),
        ("https://www.hktvmall.com/hktv/zh/gadgetsandelectronics", "gadgetsandelectronics"),
    ]
    session = scraper.build_session()
    for url, slug in urls:
        parsed = scraper.extract_category_slug_from_url(url)
        check(f"slug {slug}", parsed == slug, f"got={parsed}")
        result = scraper.hktv_fetch_page(session, categories=[slug], page_number=0, page_size=3, timeout=30)
        total = int(result.get("total") or 0)
        hits = len(result.get("hits") or [])
        check(f"API fetch {slug}", total > 0 and hits > 0, f"total={total}")


def scenario_category_codes() -> None:
    print("\n[3] Category AA codes")
    session = scraper.build_session()
    codes = ["AA28570000000", "AA11850000000", "AA11800000000"]
    for code in codes:
        parsed = scraper.extract_aa_category_code(code)
        check(f"parse {code}", parsed == code)
        total = int(scraper.hktv_fetch_page(session, categories=[code], page_number=0, page_size=3, timeout=30).get("total") or 0)
        check(f"API fetch {code}", total > 0, f"total={total}")


def scenario_aa_scope_no_overscrape() -> None:
    print("\n[4] AA code scoped tasks (no over-breakdown)")
    session = scraper.build_session()
    code = "AA28570000000"
    api_total = int(
        scraper.hktv_fetch_page(session, categories=[code], page_number=0, page_size=1, timeout=30).get("total") or 0
    )
    cfg = {
        "method": "category",
        "category_urls": [],
        "category_codes": [code],
        "auto_breakdown_categories": True,
        "timeout": 30,
    }
    tasks = scraper.build_tasks(cfg, session=session)
    check("single task for AA code", len(tasks) == 1, f"tasks={len(tasks)}")
    check("task category matches", tasks[0].get("categories") == [code])
    check("API total ~15k", 14000 <= api_total <= 16000, f"total={api_total}")
    check("no auto-breakdown for AA", not scraper.should_auto_breakdown_root(code))


def scenario_url_breakdown_vs_aa_direct() -> None:
    print("\n[5] URL breakdown vs AA direct")
    session = scraper.build_session()
    url_cfg = {
        "method": "category",
        "category_urls": ["https://www.hktvmall.com/hktv/zh/mothernbaby"],
        "category_codes": [],
        "auto_breakdown_categories": True,
        "timeout": 30,
    }
    url_tasks = scraper.build_tasks(url_cfg, session=session)
    check("mothernbaby URL splits into many tasks", len(url_tasks) > 50, f"tasks={len(url_tasks)}")
    check("mothernbaby slug breakdown enabled", scraper.should_auto_breakdown_root("mothernbaby"))

    mixed_cfg = {
        "method": "category",
        "category_urls": ["https://www.hktvmall.com/hktv/zh/mothernbaby"],
        "category_codes": ["AA28570000000"],
        "auto_breakdown_categories": True,
        "timeout": 30,
    }
    roots = scraper.resolve_category_roots(mixed_cfg)
    check("mixed roots include both", "mothernbaby" in roots and "AA28570000000" in roots, str(roots))


def scenario_locale_from_url() -> None:
    print("\n[6] Product language from URL")
    zh_cfg = {"category_urls": ["https://www.hktvmall.com/hktv/zh/mothernbaby"], "category_codes": []}
    en_cfg = {"category_urls": ["https://www.hktvmall.com/hktv/en/mothernbaby"], "category_codes": []}
    check("zh URL → hktv_zh", scraper.resolve_website_key(zh_cfg) == "hktv_zh")
    check("en URL → hktv_en", scraper.resolve_website_key(en_cfg) == "hktv_en")

    session = scraper.build_session()
    result = scraper.hktv_fetch_page(session, categories=["mothernbaby"], page_number=0, page_size=1, timeout=30)
    hits = result.get("hits") or []
    if hits:
        row_zh = scraper.parse_hktv_hit(hits[0], task=scraper.new_task(method="category", categories=["mothernbaby"]), site_key="hktv_zh")
        row_en = scraper.parse_hktv_hit(hits[0], task=scraper.new_task(method="category", categories=["mothernbaby"]), site_key="hktv_en")
        check("zh name differs from en", row_zh.get("name") != row_en.get("name"), f"zh={row_zh.get('name', '')[:30]}")


def scenario_keyword_search() -> None:
    print("\n[7] Keyword search")
    session = scraper.build_session()
    cfg = {
        "method": "keyword",
        "keywords": ["牛奶"],
        "timeout": 30,
        "website_key": "hktv_zh",
    }
    tasks = scraper.build_tasks(cfg, session=session)
    check("keyword tasks created", len(tasks) == 1)
    scraper.init_state()
    scraper.reset_run_state(keep_settings=True)
    st_mock.session_state.settings = cfg
    st_mock.session_state.tasks = tasks
    step = scraper.run_one_step(cfg)
    added = len(st_mock.session_state.rows)
    check("keyword scrape adds rows", added > 0, f"rows={added}")
    check("step status ok", step.get("status") in {"progress", "task_done", "split"})


def scenario_dedup_and_parsing() -> None:
    print("\n[8] Product parsing quality")
    session = scraper.build_session()
    result = scraper.hktv_fetch_page(session, categories=["AA28570000000"], page_number=0, page_size=5, timeout=30)
    hits = result.get("hits") or []
    check("hits returned", len(hits) > 0)
    if hits:
        row = scraper.parse_hktv_hit(
            hits[0],
            task=scraper.new_task(method="category", categories=["AA28570000000"]),
            site_key="hktv_zh",
        )
        check("has product_code", bool(row.get("product_code")))
        check("has chinese name", bool(row.get("name")))
        check("no json blob in name", not str(row.get("name", "")).startswith("[{"))
        df = scraper.rows_to_dataframe([row])
        check("export has no Pack Size", "Pack Size" not in df.columns)


def scenario_exports() -> None:
    print("\n[9] CSV and Excel export")
    session = scraper.build_session()
    result = scraper.hktv_fetch_page(session, categories=["AA28570000000"], page_number=0, page_size=5, timeout=30)
    rows = [
        scraper.parse_hktv_hit(h, task=scraper.new_task(method="category", categories=["AA28570000000"]), site_key="hktv_zh")
        for h in result.get("hits") or []
    ]
    df = scraper.rows_to_dataframe(rows)
    csv_bytes = df[scraper.EXCEL_COLUMNS].to_csv(index=False).encode("utf-8-sig")
    excel_bytes = scraper.build_excel_with_images(df, session, embed_images=False, source_rows=rows)
    check("csv bytes", len(csv_bytes) > 100)
    check("excel bytes", len(excel_bytes) > 1000)
    for col in scraper.EXCEL_COLUMNS:
        for val in df[col].astype(str):
            if str(val).startswith("[{") or str(val).startswith('{"'):
                check(f"clean col {col}", False, str(val)[:60])
                break
        else:
            continue
        break
    else:
        check("all export cells clean", True)


def scenario_short_aa_scrape_simulation() -> None:
    print("\n[10] Short AA scrape simulation (dedup + scope)")
    session = scraper.build_session()
    code = "AA28570000000"
    api_total = int(
        scraper.hktv_fetch_page(session, categories=[code], page_number=0, page_size=1, timeout=30).get("total") or 0
    )
    cfg = {
        "method": "category",
        "category_urls": [],
        "category_codes": [code],
        "page_size": 60,
        "max_pages": 2,
        "limit": 0,
        "timeout": 30,
        "website_key": "hktv_zh",
        "sort_by": "salesVolume:desc",
        "use_pdp_fallback": False,
    }
    scraper.init_state()
    scraper.reset_run_state(keep_settings=True)
    st_mock.session_state.settings = cfg
    st_mock.session_state.tasks = scraper.build_tasks(cfg, session=session)
    check("simulation uses 1 task", len(st_mock.session_state.tasks) == 1)

    for _ in range(3):
        result = scraper.run_one_step(cfg)
        if result.get("status") in {"complete", "task_done", "limit", "idle"}:
            break

    collected = len(st_mock.session_state.rows)
    check("simulation collects rows", collected > 0, f"rows={collected}")
    check("simulation well below api total", collected < api_total, f"collected={collected}, api={api_total}")
    check("still single root task scope", len(scraper.build_tasks(cfg, session=session)) == 1)


def scenario_pagination_split() -> None:
    print("\n[11] Pagination split still works")
    session = scraper.build_session()
    task = scraper.new_task(method="category", categories=["AA32080500000"], label="test")
    subs = scraper.split_oversized_task(session, task, timeout=30)
    check("split creates subtasks", len(subs) > 0, f"subs={len(subs)}")


def main() -> int:
    print("=" * 60)
    print("HKTVmall Scraper — ALL SCENARIO CHECKS")
    print(f"Version: {scraper.APP_VERSION}")
    print("=" * 60)

    scenario_version_and_columns()
    scenario_category_urls()
    scenario_category_codes()
    scenario_aa_scope_no_overscrape()
    scenario_url_breakdown_vs_aa_direct()
    scenario_locale_from_url()
    scenario_keyword_search()
    scenario_dedup_and_parsing()
    scenario_exports()
    scenario_short_aa_scrape_simulation()
    scenario_pagination_split()

    print("\n" + "=" * 60)
    print(f"RESULTS: {PASS} passed, {FAIL} failed")
    print("=" * 60)
    scraper.close_session()
    return 0 if FAIL == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
