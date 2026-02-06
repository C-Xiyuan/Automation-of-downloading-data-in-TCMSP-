#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import os
import re
import sys
import time
import traceback
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import parse_qsl, urlencode, urljoin, urlsplit, urlunsplit

import pandas as pd
try:
    from bs4 import BeautifulSoup
except Exception:
    BeautifulSoup = None
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError, sync_playwright

BASE_URL = "https://tcmsp-e.com"
DEFAULT_ENTRY_URLS = [
    f"{BASE_URL}/tcmsp.php",
    f"{BASE_URL}/load_intro.php?id=43",
]

DEFAULT_TIMEOUT_MS = 30_000
RETRY_ATTEMPTS = 3
RETRY_SLEEP_SEC = 2

HEADLESS = os.environ.get("HEADLESS", "1").lower() in {"1", "true", "yes"}
SLOW_MO = int(os.environ.get("SLOW_MO", "0") or 0)
TRACE = os.environ.get("TRACE", "0").lower() in {"1", "true", "yes"}

UA = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/120.0.0.0 Safari/537.36"
)

TARGET_KEYWORDS = ["target", "protein", "uniprot", "gene", "symbol"]
RELATED_COLUMNS = [
    "molecule_ID",
    "MOL_ID",
    "molecule_name",
    "target_name",
    "target_ID",
    "drugbank_ID",
    "validated",
    "SVM_score",
    "RF_score",
]
NO_RESULTS_TEXT = "No items to display"


def sanitize_filename(name: str) -> str:
    name = re.sub(r"\s+", "_", name.strip())
    name = re.sub(r"[\\/:*?\"<>|]", "_", name)
    name = re.sub(r"[^0-9A-Za-z_\u4e00-\u9fff-]+", "_", name)
    return name.strip("_") or "herb"


def make_debug_dir(herb: str) -> Path:
    ts = time.strftime("%Y%m%d_%H%M%S")
    safe = sanitize_filename(herb)
    path = Path.cwd() / "debug" / f"{ts}_{safe}"
    path.mkdir(parents=True, exist_ok=True)
    return path


def normalize_url(url: str) -> str:
    try:
        parts = urlsplit(url)
        netloc = parts.netloc.replace("www.tcmsp-e.com", "tcmsp-e.com")
        query = urlencode(parse_qsl(parts.query, keep_blank_values=True))
        return urlunsplit((parts.scheme, netloc, parts.path, query, parts.fragment))
    except Exception:
        return url


def dump_debug(page, debug_dir: Path, label: str) -> None:
    try:
        html_path = debug_dir / f"{label}.html"
        html_path.write_text(page.content(), encoding="utf-8")
    except Exception:
        pass
    try:
        png_path = debug_dir / f"{label}.png"
        page.screenshot(path=str(png_path), full_page=True)
    except Exception:
        pass


def append_step(debug_dir: Path, text: str) -> None:
    try:
        with (debug_dir / "step_urls.txt").open("a", encoding="utf-8", errors="ignore") as f:
            f.write(text.rstrip() + "\n")
    except Exception:
        pass


def retry(action, attempts: int = RETRY_ATTEMPTS, sleep_sec: int = RETRY_SLEEP_SEC):
    last_exc = None
    for i in range(attempts):
        try:
            return action()
        except Exception as exc:
            last_exc = exc
            if i < attempts - 1:
                time.sleep(sleep_sec)
    raise last_exc


def safe_goto(page, url: str, debug_dir: Path, label: str) -> None:
    target = normalize_url(url)
    append_step(debug_dir, f"{label}: {target}")
    print(f"[GOTO] {target}")
    page.goto(target, wait_until="domcontentloaded")
    page.wait_for_load_state("networkidle")
    print(f"[URL] {page.url}")
    dump_debug(page, debug_dir, label)


def close_daily_popup(page) -> None:
    selectors = ["#dc", "button#dc", "text=我知道了", "text=关闭"]
    for sel in selectors:
        locator = page.locator(sel)
        if locator.count() > 0:
            try:
                locator.first.click(force=True)
                page.wait_for_timeout(300)
                return
            except Exception:
                continue
    try:
        page.evaluate(
            "() => {"
            " const el = document.getElementById('dp'); if (el) el.remove();"
            " const btn = document.getElementById('dc'); if (btn) btn.remove();"
            " const overlays = document.querySelectorAll('.k-overlay, .modal-backdrop');"
            " overlays.forEach(o => o.remove());"
            "}"
        )
    except Exception:
        pass


def load_entry_urls() -> List[str]:
    path = Path.cwd() / "step_urls.txt"
    if not path.exists():
        return DEFAULT_ENTRY_URLS
    urls = []
    for line in path.read_text(encoding="utf-8", errors="ignore").splitlines():
        line = line.strip()
        if line.startswith("http"):
            urls.append(line)
    return urls or DEFAULT_ENTRY_URLS


def log_state(page, label: str) -> None:
    try:
        grid = page.locator("#grid").count() > 0
        related = page.get_by_text(re.compile(r"Related\\s*Targets|相关靶点", re.I)).count() > 0
        grid2 = page.locator("#grid2").count() > 0
    except Exception:
        grid = False
        related = False
        grid2 = False
    print(f"[STATE] {label} url={page.url} grid={grid} related={related} grid2={grid2}")


def open_entry(page, debug_dir: Path) -> None:
    entry_urls = load_entry_urls()
    last_error = None
    for idx, url in enumerate(entry_urls, 1):
        try:
            safe_goto(page, url, debug_dir, f"entry_{idx}")
            return
        except Exception as exc:
            last_error = exc
    raise RuntimeError(f"All entry URLs failed: {last_error}")


def find_search_input(page):
    candidates = [
        "#inputVarTcm",
        "input[name='q']",
        "input[type='text']",
    ]
    for sel in candidates:
        locator = page.locator(sel)
        if locator.count() > 0 and locator.first.is_visible():
            return locator.first
    return None


def click_search(page):
    candidates = [
        "#searchBtTcm",
        "input[type='submit']",
        "button:has-text('Search')",
        "button",
    ]
    for sel in candidates:
        locator = page.locator(sel)
        if locator.count() == 0:
            continue
        try:
            locator.first.click()
            return True
        except Exception:
            continue
    return False


def wait_for_grid_ready(page) -> None:
    page.wait_for_selector("#grid .k-grid-content tbody tr", timeout=DEFAULT_TIMEOUT_MS)
    try:
        page.wait_for_function(
            "() => window.$ && $('#grid').data('kendoGrid') && $('#grid').data('kendoGrid').dataSource && $('#grid').data('kendoGrid').dataSource.total() > 0",
            timeout=DEFAULT_TIMEOUT_MS,
        )
    except Exception:
        pass


def search_has_no_results(page) -> bool:
    try:
        if page.locator(".k-grid-norecords").count() > 0:
            return True
    except Exception:
        pass
    try:
        text = page.inner_text("body")
        if NO_RESULTS_TEXT.lower() in text.lower():
            return True
    except Exception:
        pass
    return False


def is_search_list_page(page) -> bool:
    try:
        headers = []
        header_cells = page.locator("#grid .k-grid-header th")
        for i in range(header_cells.count()):
            try:
                headers.append(header_cells.nth(i).inner_text().strip().lower())
            except Exception:
                text = header_cells.nth(i).text_content() or ""
                headers.append(text.strip().lower())
        if any("latin name" in h for h in headers) or any("chinese name" in h for h in headers):
            return True
    except Exception:
        pass
    try:
        html = page.content()
        if "#grid" in html and "k-grid-content" in html and "Search by" in html:
            return True
    except Exception:
        pass
    return False


def get_latin_link(page) -> Optional[Tuple[str, Any]]:
    try:
        header_cells = page.locator("#grid .k-grid-header th")
        headers = []
        for i in range(header_cells.count()):
            try:
                headers.append(header_cells.nth(i).inner_text().strip().lower())
            except Exception:
                text = header_cells.nth(i).text_content() or ""
                headers.append(text.strip().lower())
        if not any("latin name" in h for h in headers):
            return None
    except Exception:
        pass
    # Required selector (per spec): first row, third column, <a>
    row = page.locator("#grid .k-grid-content tbody tr").first
    link = row.locator("td").nth(2).locator("a").first
    if link.count() == 0:
        return None
    href = link.get_attribute("href")
    return href, link


def extract_latin_href_from_html(html: str) -> Optional[str]:
    if BeautifulSoup is None:
        return None
    soup = BeautifulSoup(html, "lxml")
    grid = soup.find(id="grid")
    if not grid:
        return None
    headers = [th.get_text(" ", strip=True).lower() for th in grid.select(".k-grid-header th")]
    latin_idx = None
    for idx, h in enumerate(headers):
        if "latin name" in h:
            latin_idx = idx
            break
    if latin_idx is None:
        return None
    first_row = grid.select_one(".k-grid-content tbody tr")
    if not first_row:
        return None
    cells = first_row.find_all("td")
    if latin_idx >= len(cells):
        return None
    a = cells[latin_idx].find("a", href=True)
    if not a:
        return None
    return a.get("href")


def click_latin_link(page, debug_dir: Path, depth: int) -> Optional[object]:
    link_info = get_latin_link(page)
    if not link_info:
        return None
    href, locator = link_info
    if href:
        append_step(debug_dir, f"latin_href: {href}")
    try:
        locator.scroll_into_view_if_needed()
    except Exception:
        pass
    try:
        with page.expect_navigation(wait_until="domcontentloaded", timeout=10_000):
            locator.click(force=True)
        page.wait_for_load_state("networkidle")
        dump_debug(page, debug_dir, f"drill_click_{depth}")
        return page
    except Exception:
        pass
    return None


def goto_latin_link(page, debug_dir: Path, depth: int) -> Optional[object]:
    link_info = get_latin_link(page)
    if not link_info:
        return None
    href, _locator = link_info
    if not href:
        return None
    target = normalize_url(urljoin(BASE_URL + "/", href))
    safe_goto(page, target, debug_dir, f"drill_link_{depth}")
    return page


def drill_down_to_detail(page, herb: str, debug_dir: Path) -> object:
    last_url = None
    stagnant_rounds = 0
    for depth in range(4):
        close_daily_popup(page)
        log_state(page, f"drill_{depth}_before")
        dump_debug(page, debug_dir, f"drill_{depth}")
        append_step(debug_dir, f"drill_{depth}_url: {page.url}")

        if not is_search_list_page(page):
            return page

        wait_for_grid_ready(page)

        # 1) Try direct click
        clicked = click_latin_link(page, debug_dir, depth)
        if clicked is not None and not is_search_list_page(page):
            return clicked

        # 2) Fallback: direct navigation to href
        navigated = goto_latin_link(page, debug_dir, depth)
        if navigated is not None and not is_search_list_page(page):
            return navigated

        # 3) HTML-based fallback (even if locator fails)
        try:
            href = extract_latin_href_from_html(page.content())
            if href:
                append_step(debug_dir, f"latin_href_html: {href}")
                target = normalize_url(urljoin(BASE_URL + "/", href))
                safe_goto(page, target, debug_dir, f"drill_html_{depth}")
                if not is_search_list_page(page):
                    return page
        except Exception:
            pass

        # 3) If still on list, loop (multi drill-down)
        if last_url == page.url:
            stagnant_rounds += 1
        else:
            stagnant_rounds = 0
        last_url = page.url
        if stagnant_rounds >= 2:
            break

    return page


def normalize_headers(headers: List[str]) -> List[str]:
    cleaned = []
    seen = {}
    for h in headers:
        h = re.sub(r"\s+", " ", h).strip() or "col"
        key = h.lower()
        count = seen.get(key, 0) + 1
        seen[key] = count
        if count > 1:
            h = f"{h}_{count}"
        cleaned.append(h)
    return cleaned


def score_table(table: BeautifulSoup) -> int:
    headers = [th.get_text(" ", strip=True).lower() for th in table.find_all("th")]
    score = 0
    for h in headers:
        for k in TARGET_KEYWORDS:
            if k in h:
                score += 3
    for h in headers:
        if "target" in h:
            score += 2
    return score


def table_to_df(table: BeautifulSoup) -> pd.DataFrame:
    try:
        dfs = pd.read_html(str(table))
        if dfs:
            return dfs[0]
    except Exception:
        pass

    rows = []
    for tr in table.find_all("tr"):
        cells = [c.get_text(" ", strip=True) for c in tr.find_all(["th", "td"])]
        if cells:
            rows.append(cells)
    if not rows:
        return pd.DataFrame()
    headers = rows[0]
    data = rows[1:] if len(rows) > 1 else []
    return pd.DataFrame(data, columns=normalize_headers(headers))


def find_related_table_in_html(html: str) -> Optional[BeautifulSoup]:
    if BeautifulSoup is None:
        return None
    soup = BeautifulSoup(html, "lxml")
    related_text = soup.find_all(string=re.compile(r"(Related\\s*Targets|相关靶点)", re.I))
    for t in related_text:
        parent = t.parent
        table = parent.find_next("table") if parent else None
        if table:
            return table
    tables = soup.find_all("table")
    if not tables:
        return None
    scored = sorted(((score_table(t), t) for t in tables), key=lambda x: x[0], reverse=True)
    if scored and scored[0][0] >= 3:
        return scored[0][1]
    return None


def extract_from_frames(page) -> Optional[pd.DataFrame]:
    frames = list({page.main_frame} | set(page.frames))
    for frame in frames:
        try:
            html = frame.content()
        except Exception:
            continue
        table = find_related_table_in_html(html)
        if table:
            df = table_to_df(table)
            if not df.empty:
                return df
    return None


def wait_for_visible(page, selector: str) -> None:
    page.wait_for_function(
        "(sel) => { const el = document.querySelector(sel); return !!(el && el.offsetParent); }",
        selector,
        timeout=DEFAULT_TIMEOUT_MS,
    )


def click_tab_by_text(page, container_selector: str, text_re: re.Pattern) -> bool:
    tabs = page.locator(f"{container_selector} [role='tab']").filter(has_text=text_re)
    if tabs.count() == 0:
        return False
    try:
        tabs.first.scroll_into_view_if_needed()
    except Exception:
        pass
    tabs.first.click()
    return True


def select_kendo_tabstrip_index(page, selector: str, index: int) -> bool:
    return bool(
        page.evaluate(
            "(args) => {"
            " const sel = args.sel; const idx = args.idx;"
            " const el = document.querySelector(sel);"
            " if (!el) return false;"
            " const $ = window.$ || window.jQuery;"
            " if (!$) return false;"
            " const widget = $(el).data('kendoTabStrip');"
            " if (!widget) return false;"
            " widget.select(idx);"
            " return true;"
            "}",
            {"sel": selector, "idx": index},
        )
    )


def ensure_tab_visible(page, selector: str) -> bool:
    try:
        return bool(
            page.evaluate(
                "(sel) => { const el = document.querySelector(sel); return !!(el && el.offsetParent); }",
                selector,
            )
        )
    except Exception:
        return False


def force_show_tabstrip_content(page, container_selector: str, content_selector: str) -> None:
    page.evaluate(
        "(containerSel, contentSel) => {"
        " const container = document.querySelector(containerSel);"
        " if (!container) return;"
        " const contents = container.querySelectorAll('.k-content');"
        " contents.forEach(c => {"
        "   if (c.matches(contentSel)) {"
        "     c.style.display = 'block';"
        "     c.classList.add('k-state-active');"
        "     c.setAttribute('aria-hidden','false');"
        "     c.setAttribute('aria-expanded','true');"
        "   } else {"
        "     c.style.display = 'none';"
        "     c.classList.remove('k-state-active');"
        "     c.setAttribute('aria-hidden','true');"
        "     c.setAttribute('aria-expanded','false');"
        "   }"
        " });"
        " const tabs = container.querySelectorAll('.k-tabstrip-items li');"
        " tabs.forEach((li, idx) => {"
        "   const target = li.getAttribute('aria-controls');"
        "   if (target && ('#'+target) === contentSel) {"
        "     li.classList.add('k-state-active','k-tab-on-top');"
        "     li.setAttribute('aria-selected','true');"
        "   } else {"
        "     li.classList.remove('k-state-active','k-tab-on-top');"
        "     li.setAttribute('aria-selected','false');"
        "   }"
        " });"
        "}",
        container_selector,
        content_selector,
    )


def ensure_related_targets_section(page, debug_dir: Path) -> None:
    # Make several attempts in case popup or JS timing blocks clicks.
    for attempt in range(3):
        close_daily_popup(page)

        # Prefer Kendo API select to avoid overlay click interception.
        selected = select_kendo_tabstrip_index(page, "#tabstrip", 1)
        if not selected:
            click_tab_by_text(page, "#tabstrip", re.compile(r"Related\\s*Targets|相关靶点", re.I))
        page.wait_for_timeout(300)

        if ensure_tab_visible(page, "#tabstrip-2"):
            dump_debug(page, debug_dir, f"after_related_tab_{attempt+1}")
            break
        page.wait_for_timeout(500)

    if not ensure_tab_visible(page, "#tabstrip-2"):
        # Force show tabstrip-2 if clicks fail (e.g., blocked by modal)
        force_show_tabstrip_content(page, "#tabstrip", "#tabstrip-2")
        page.wait_for_timeout(300)
        dump_debug(page, debug_dir, "after_related_tab_forced")

    # Ensure nested Targets Information tab
    info_regex = re.compile(r"Targets\\s*Info", re.I)
    for attempt in range(3):
        close_daily_popup(page)
        selected = select_kendo_tabstrip_index(page, "#tabstrip-2", 1)
        if not selected:
            click_tab_by_text(page, "#tabstrip-2", info_regex)
        page.wait_for_timeout(300)
        if ensure_tab_visible(page, "#tabstrip2-2"):
            dump_debug(page, debug_dir, f"after_targets_info_tab_{attempt+1}")
            break
        page.wait_for_timeout(500)

    if not ensure_tab_visible(page, "#tabstrip2-2"):
        force_show_tabstrip_content(page, "#tabstrip-2", "#tabstrip2-2")
        page.wait_for_timeout(300)
        dump_debug(page, debug_dir, "after_targets_info_tab_forced")

    # Ensure grid2 is present (attached is enough; visibility handled above)
    page.wait_for_selector("#grid2 .k-grid-content tbody tr", state="attached", timeout=DEFAULT_TIMEOUT_MS)

    if not ensure_tab_visible(page, "#tabstrip-2"):
        raise RuntimeError("Related Targets tab not visible after selection.")


def get_kendo_state(page, grid_id: str) -> Optional[Dict[str, Any]]:
    return page.evaluate(
        "(gridSel) => {"
        " const el = document.querySelector(gridSel);"
        " if (!el || !window.$) return null;"
        " const grid = window.$(el).data('kendoGrid');"
        " if (!grid || !grid.dataSource) return null;"
        " const total = grid.dataSource.total();"
        " const pageSize = grid.dataSource.pageSize();"
        " const page = grid.dataSource.page();"
        " const pageCount = Math.ceil(total / pageSize) || 1;"
        " return { total, pageSize, page, pageCount };"
        "}",
        f"#{grid_id}",
    )


def set_kendo_page(page, grid_id: str, page_num: int) -> bool:
    ok = page.evaluate(
        "(gridSel, num) => {"
        " const el = document.querySelector(gridSel);"
        " if (!el || !window.$) return false;"
        " const grid = window.$(el).data('kendoGrid');"
        " if (!grid || !grid.dataSource) return false;"
        " grid.dataSource.page(num);"
        " return true;"
        "}",
        f"#{grid_id}",
        page_num,
    )
    if not ok:
        return False
    page.wait_for_function(
        "(gridSel, num) => {"
        " const el = document.querySelector(gridSel);"
        " if (!el || !window.$) return false;"
        " const grid = window.$(el).data('kendoGrid');"
        " return grid && grid.dataSource.page() === num;"
        "}",
        f"#{grid_id}",
        page_num,
        timeout=DEFAULT_TIMEOUT_MS,
    )
    return True


def wait_for_grid_idle(page, grid_id: str) -> None:
    page.wait_for_function(
        "(gridSel) => {"
        " const el = document.querySelector(gridSel);"
        " if (!el) return false;"
        " const mask = el.querySelector('.k-loading-mask');"
        " return !mask || mask.offsetParent === null;"
        "}",
        f"#{grid_id}",
        timeout=DEFAULT_TIMEOUT_MS,
    )


def read_grid_table(page, grid_id: str) -> Tuple[List[str], List[List[str]]]:
    headers = []
    header_cells = page.locator(f"#{grid_id} .k-grid-header th")
    for i in range(header_cells.count()):
        try:
            text = header_cells.nth(i).inner_text().strip()
        except Exception:
            text = header_cells.nth(i).text_content() or ""
            text = text.strip()
        headers.append(text or f"col_{i+1}")

    rows = []
    row_locator = page.locator(f"#{grid_id} .k-grid-content tbody tr")
    for r in range(row_locator.count()):
        cells = row_locator.nth(r).locator("td")
        row_vals = []
        for c in range(cells.count()):
            try:
                cell_text = cells.nth(c).inner_text().strip()
            except Exception:
                cell_text = cells.nth(c).text_content() or ""
                cell_text = cell_text.strip()
            row_vals.append(cell_text)
        rows.append(row_vals)
    return headers, rows


def extract_kendo_grid_all_pages(page, grid_id: str, debug_dir: Path) -> pd.DataFrame:
    state = get_kendo_state(page, grid_id)
    if not state:
        raise RuntimeError(f"Kendo grid state unavailable for #{grid_id}.")
    total_pages = int(state.get("pageCount") or 1)
    current_page = int(state.get("page") or 1)
    append_step(debug_dir, f"{grid_id}_pages: {total_pages}")
    print(f"[GRID] #{grid_id} pages={total_pages} page_size={state.get('pageSize')} total={state.get('total')}")

    headers: List[str] = []
    rows: List[List[str]] = []
    seen = set()

    for page_num in range(1, total_pages + 1):
        if page_num != current_page:
            set_kendo_page(page, grid_id, page_num)
        wait_for_grid_idle(page, grid_id)
        page.wait_for_selector(f"#{grid_id} .k-grid-content tbody tr", timeout=DEFAULT_TIMEOUT_MS)
        if page_num == 1 or page_num == total_pages:
            dump_debug(page, debug_dir, f"{grid_id}_page_{page_num}")

        header_row, page_rows = read_grid_table(page, grid_id)
        if not headers:
            headers = header_row
        for row in page_rows:
            key = tuple(row)
            if key in seen:
                continue
            seen.add(key)
            rows.append(row)
        append_step(debug_dir, f"{grid_id}_page_done: {page_num}/{total_pages} rows={len(rows)}")
        current_page = page_num

    if not headers:
        return pd.DataFrame()

    normalized = normalize_headers(headers)
    normalized_rows = []
    for row in rows:
        if len(row) < len(normalized):
            row = row + [""] * (len(normalized) - len(row))
        elif len(row) > len(normalized):
            row = row[: len(normalized)]
        normalized_rows.append(row)
    return pd.DataFrame(normalized_rows, columns=normalized)


def extract_from_xhr(response_log: List[Dict[str, Any]]) -> Optional[pd.DataFrame]:
    for item in response_log:
        body = item.get("body_text")
        if not body:
            continue
        content_type = (item.get("content_type") or "").lower()
        if "application/json" in content_type or body.strip().startswith(("{", "[")):
            try:
                data = json.loads(body)
            except Exception:
                continue
            df = json_to_df(data)
            if df is not None and not df.empty:
                return df
        if "text/html" in content_type or "<table" in body.lower():
            table = find_related_table_in_html(body)
            if table:
                df = table_to_df(table)
                if not df.empty:
                    return df
    return None


def json_to_df(data: Any) -> Optional[pd.DataFrame]:
    candidates = []

    def walk(obj):
        if isinstance(obj, list):
            if obj and all(isinstance(x, dict) for x in obj):
                candidates.append(obj)
            for x in obj:
                walk(x)
        elif isinstance(obj, dict):
            for v in obj.values():
                walk(v)

    walk(data)
    if not candidates:
        return None

    def score_rows(rows: List[Dict[str, Any]]) -> int:
        keys = set()
        for r in rows:
            keys.update(str(k).lower() for k in r.keys())
        return sum(2 for k in keys if any(kw in k for kw in TARGET_KEYWORDS)) + len(keys)

    best = max(candidates, key=score_rows)
    return pd.DataFrame(best)


def extract_grid2_data_from_html(html: str) -> Optional[pd.DataFrame]:
    marker_patterns = [
        r'\$\(["\']#grid2["\']\)\.kendoGrid\(',
        r'$("#grid2")\.kendoGrid\(',
        r'grid2"\)\.kendoGrid\(',
    ]
    start_idx = -1
    for pat in marker_patterns:
        m = re.search(pat, html)
        if m:
            start_idx = m.start()
            break
    if start_idx == -1:
        return None

    segment = html[start_idx:]
    ds_idx = segment.find("dataSource")
    if ds_idx == -1:
        return None
    data_idx = segment.find("data:", ds_idx)
    if data_idx == -1:
        return None
    array_start = segment.find("[", data_idx)
    if array_start == -1:
        return None

    depth = 0
    end_idx = -1
    for i in range(array_start, len(segment)):
        ch = segment[i]
        if ch == "[":
            depth += 1
        elif ch == "]":
            depth -= 1
            if depth == 0:
                end_idx = i
                break
    if end_idx == -1:
        return None

    json_text = segment[array_start : end_idx + 1]
    try:
        data = json.loads(json_text)
    except Exception:
        return None
    if not isinstance(data, list) or not data:
        return None
    if not all(isinstance(x, dict) for x in data):
        return None
    return pd.DataFrame(data)


def extract_grid2_data_from_kendo(page) -> Optional[pd.DataFrame]:
    try:
        data = page.evaluate(
            "() => {"
            " const $ = window.$ || window.jQuery;"
            " if (!$) return null;"
            " const grid = $('#grid2').data('kendoGrid');"
            " if (!grid || !grid.dataSource) return null;"
            " const data = grid.dataSource.data();"
            " if (!data) return null;"
            " return data.toJSON ? data.toJSON() : data;"
            "}"
        )
    except Exception:
        return None
    if not data or not isinstance(data, list):
        return None
    if not all(isinstance(x, dict) for x in data):
        return None
    return pd.DataFrame(data)


def normalize_related_targets_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df = df.replace({"": pd.NA})
    for col in ["molecule_ID", "target_ID"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
            if df[col].notna().all():
                df[col] = df[col].astype(int)
    for col in ["SVM_score", "RF_score"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    if "validated" in df.columns:
        df["validated"] = df["validated"].where(df["validated"].notna(), pd.NA)
    cols = [c for c in RELATED_COLUMNS if c in df.columns]
    extras = [c for c in df.columns if c not in cols]
    df = df[cols + extras]
    return df


def get_success_index_from_args(argv: List[str]) -> Optional[str]:
    if "--success-index" in argv:
        idx = argv.index("--success-index")
        if idx + 1 < len(argv):
            return argv[idx + 1].strip()
    return None


def save_success_xlsx(df: pd.DataFrame, herb: str, argv: List[str]) -> Path:
    desktop = Path.home() / "Desktop"
    idx = os.environ.get("SUCCESS_INDEX") or get_success_index_from_args(argv)
    suffix = f"{idx}" if idx else ""
    out_path = desktop / f"Success{suffix}.xlsx"
    df.to_excel(out_path, index=False)
    if not out_path.exists() or out_path.stat().st_size == 0:
        raise RuntimeError("Success.xlsx was not created or is empty.")
    return out_path


def compare_with_reference(df: pd.DataFrame, ref_path: Path) -> bool:
    if not ref_path.exists():
        return False
    ref = pd.read_excel(ref_path)
    df_norm = normalize_related_targets_df(df)
    ref_norm = normalize_related_targets_df(ref)
    # Align columns to reference for strict comparison
    if list(ref_norm.columns) != list(df_norm.columns):
        df_norm = df_norm.reindex(columns=ref_norm.columns)
    # Cast to reference dtypes to avoid False due to dtype differences
    for col, dtype in ref_norm.dtypes.items():
        if col not in df_norm.columns:
            continue
        kind = getattr(dtype, "kind", None)
        if kind in {"i", "u"}:
            s = pd.to_numeric(df_norm[col], errors="coerce")
            if s.notna().all():
                s = s.astype(int)
            df_norm[col] = s
        elif kind == "f":
            df_norm[col] = pd.to_numeric(df_norm[col], errors="coerce")
        else:
            # Match pandas StringDtype if reference uses it
            try:
                is_string = isinstance(dtype, pd.StringDtype) or str(dtype) in {"string", "str"}
            except Exception:
                is_string = str(dtype) in {"string", "str"}
            if is_string:
                df_norm[col] = df_norm[col].astype("string")
            else:
                df_norm[col] = df_norm[col].astype(object)
    try:
        pd.testing.assert_frame_equal(df_norm, ref_norm, check_dtype=False)
        return True
    except Exception:
        return False


def extract_related_targets(page, response_log: List[Dict[str, Any]], debug_dir: Path) -> Tuple[pd.DataFrame, str]:
    # Try to ensure tab selection, but do not fail hard if UI is blocked.
    try:
        ensure_related_targets_section(page, debug_dir)
    except Exception as exc:
        print(f"[WARN] ensure_related_targets_section failed: {exc}")

    # First try Kendo dataSource (authoritative values)
    df = extract_grid2_data_from_kendo(page)
    if df is not None and not df.empty:
        return normalize_related_targets_df(df), "Kendo-DataSource"

    # Fallback: embedded HTML dataSource
    html = page.content()
    df = extract_grid2_data_from_html(html)
    if df is not None and not df.empty:
        return normalize_related_targets_df(df), "HTML-embedded"

    # Last fallback: paginate table view (may have display labels instead of IDs)
    try:
        df = extract_kendo_grid_all_pages(page, "grid2", debug_dir)
        if not df.empty:
            return normalize_related_targets_df(df), "DOM-Kendo"
    except Exception:
        pass

    df = extract_from_frames(page)
    if df is not None and not df.empty:
        return df, "DOM"

    df = extract_from_xhr(response_log)
    if df is not None and not df.empty:
        return normalize_related_targets_df(df), "XHR"

    raise RuntimeError("Related Targets not found in DOM, Kendo grid, or XHR.")


def save_xlsx(df: pd.DataFrame, herb: str) -> Path:
    desktop = Path.home() / "Desktop"
    filename = f"{sanitize_filename(herb)}_RelatedTargets.xlsx"
    out_path = desktop / filename
    df.to_excel(out_path, index=False)
    return out_path


def chunk_string(text: str, max_len: int = 30000) -> List[str]:
    if not text:
        return [""]
    chunks = []
    for i in range(0, len(text), max_len):
        chunks.append(text[i : i + max_len])
    return chunks


def text_to_df(text: str, label: str) -> pd.DataFrame:
    rows = []
    for idx, line in enumerate(text.splitlines(), 1):
        if len(line) > 30000:
            for sub_idx, chunk in enumerate(chunk_string(line), 1):
                rows.append((f"{idx}.{sub_idx}", chunk))
        else:
            rows.append((str(idx), line))
    if not rows:
        rows = [("1", "")]
    return pd.DataFrame(rows, columns=["line", label])


def html_to_df(html: str) -> pd.DataFrame:
    chunks = chunk_string(html or "")
    rows = [(i + 1, chunk) for i, chunk in enumerate(chunks)]
    return pd.DataFrame(rows, columns=["chunk", "html"])


def extract_tables_from_html(html: str) -> List[Tuple[str, pd.DataFrame]]:
    tables: List[Tuple[str, pd.DataFrame]] = []
    try:
        dfs = pd.read_html(html)
    except Exception:
        dfs = []
    for idx, df in enumerate(dfs, 1):
        if df is not None and not df.empty:
            tables.append((f"table_{idx}", df))
    return tables


def extract_detail_page_content(page, debug_dir: Path) -> Dict[str, pd.DataFrame]:
    dump_debug(page, debug_dir, "detail_page")
    html = page.content()
    try:
        text = page.inner_text("body")
    except Exception:
        try:
            text = page.evaluate("() => document.body && document.body.innerText ? document.body.innerText : ''")
        except Exception:
            text = ""

    sheets: Dict[str, pd.DataFrame] = {}
    sheets["page_text"] = text_to_df(text, "text")
    sheets["page_html"] = html_to_df(html)

    for name, df in extract_tables_from_html(html):
        sheets[name] = df

    return sheets


def save_xlsx_multi(sheets: Dict[str, pd.DataFrame], herb: str) -> Path:
    desktop = Path.home() / "Desktop"
    filename = f"{sanitize_filename(herb)}_DetailPage.xlsx"
    out_path = desktop / filename
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            safe_name = re.sub(r"[^0-9A-Za-z_]+", "_", sheet_name)[:31] or "sheet"
            df.to_excel(writer, index=False, sheet_name=safe_name)
    return out_path


def run_once(herb: str, headless: bool) -> None:
    debug_dir = make_debug_dir(herb)
    response_log: List[Dict[str, Any]] = []

    def on_response(resp):
        try:
            rtype = resp.request.resource_type
        except Exception:
            rtype = ""
        if rtype not in {"xhr", "fetch"}:
            return
        try:
            headers = resp.headers
        except Exception:
            headers = {}
        content_type = headers.get("content-type", "")
        entry: Dict[str, Any] = {
            "url": resp.url,
            "status": resp.status,
            "resource_type": rtype,
            "content_type": content_type,
        }
        try:
            length = int(headers.get("content-length", "0") or "0")
        except Exception:
            length = 0
        if length and length > 2_000_000:
            response_log.append(entry)
            return
        try:
            entry["body_text"] = resp.text()
        except Exception:
            entry["body_text"] = None
        response_log.append(entry)

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=headless,
            slow_mo=SLOW_MO if not headless else 0,
            args=["--disable-blink-features=AutomationControlled", "--no-sandbox"],
        )
        context = browser.new_context(
            user_agent=UA,
            locale="zh-CN",
            viewport={"width": 1280, "height": 800},
        )
        today = time.strftime("%Y-%m-%d")
        context.add_init_script(
            f"try{{localStorage.setItem('tcmsp_daily_popup_date','{today}')}}catch(e){{}}"
        )
        if TRACE:
            context.tracing.start(screenshots=True, snapshots=True, sources=False)
        context.on("response", on_response)
        page = context.new_page()
        page.set_default_timeout(DEFAULT_TIMEOUT_MS)

        try:
            retry(lambda: open_entry(page, debug_dir))
            close_daily_popup(page)
            log_state(page, "after_entry")
            search_input = retry(lambda: find_search_input(page))
            if not search_input:
                raise RuntimeError("Search input not found.")
            search_input.fill(herb)
            if not click_search(page):
                raise RuntimeError("Search button not found or click failed.")

            page.wait_for_load_state("networkidle")
            dump_debug(page, debug_dir, "after_search")
            log_state(page, "after_search")

            if search_has_no_results(page):
                empty_df = pd.DataFrame(columns=RELATED_COLUMNS)
                out_path = save_success_xlsx(empty_df, herb, sys.argv)
                print("[WARN] No search results for herb. Exported empty file.")
                print(f"[SAVED] {out_path}")
                print(f"[DEBUG] {debug_dir}")
                return

            page = drill_down_to_detail(page, herb, debug_dir)

            df, source = extract_related_targets(page, response_log, debug_dir)
            df = normalize_related_targets_df(df)
            out_path = save_success_xlsx(df, herb, sys.argv)
            print(f"[OK] Related Targets extracted via {source}")
            print(f"[SAVED] {out_path}")

            ref_path = Path.home() / "Desktop" / f"{herb}_RelatedTargets.xlsx"
            if ref_path.exists():
                matched = compare_with_reference(df, ref_path)
                print(f"[COMPARE] {ref_path} matched={matched}")
                if not matched:
                    raise RuntimeError("Output does not match reference file.")
            print(f"[DEBUG] {debug_dir}")

        except Exception as exc:
            dump_debug(page, debug_dir, "failure")
            debug_path = debug_dir / "error.txt"
            debug_path.write_text(
                f"{exc}\n\n{traceback.format_exc()}",
                encoding="utf-8",
            )
            log_path = debug_dir / "xhr_log.json"
            try:
                log_path.write_text(json.dumps(response_log, ensure_ascii=False, indent=2), encoding="utf-8")
            except Exception:
                pass
            print(f"[FAIL] {exc}")
            print(f"[DEBUG] {debug_dir}")
            raise
        finally:
            try:
                log_path = debug_dir / "xhr_log.json"
                if not log_path.exists():
                    log_path.write_text(
                        json.dumps(response_log, ensure_ascii=False, indent=2),
                        encoding="utf-8",
                    )
            except Exception:
                pass
            if TRACE:
                try:
                    context.tracing.stop(path=str(debug_dir / "trace.zip"))
                except Exception:
                    pass
            context.close()
            browser.close()


def self_check(herb: str) -> None:
    print("[SELF-CHECK] Run 1 (headless)")
    run_once(herb, True)
    print("[SELF-CHECK] Run 2 (headless)")
    run_once(herb, True)
    print("[SELF-CHECK] Run 3 (headed)")
    run_once(herb, False)


def main() -> None:
    herb = sys.argv[1].strip() if len(sys.argv) > 1 else input("请输入药名：").strip()
    if not herb:
        print("药名不能为空。")
        return
    if os.environ.get("SELF_CHECK", "0") in {"1", "true", "yes"}:
        self_check(herb)
    else:
        run_once(herb, HEADLESS)


if __name__ == "__main__":
    main()
