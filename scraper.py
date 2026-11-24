"""Playwright-based scraper for Rajasthan SEC Contesting Sarpanch data."""

from __future__ import annotations

import csv
import html.parser
import logging
import os
import random
import re
import time
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from playwright.sync_api import Browser, BrowserContext, Page, TimeoutError as PlaywrightTimeoutError, sync_playwright

from progress_tracker import ProgressTracker, atomic_write_csv, make_ps_key
from schema import PORTAL_FIELD_KEYS, clean_text, sanitize_filename

# Verified exact selectors on https://sec.rajasthan.gov.in/GramPanchayatDetails.aspx
SEL_ELECTION_TYPE = "#ContentPlaceHolder1_DropDownListElection"
SEL_ELECTION_PERIOD = "#ContentPlaceHolder1_DropDownListPeriod"
SEL_DISTRICT = "#ContentPlaceHolder1_DistrictDropDownList"
SEL_PANCHAYAT_SAMITI = "#PanchayatSamitiDropDownList"  # Note: no ContentPlaceHolder1_ prefix on id
SEL_GRAM_PANCHAYAT = "#ContentPlaceHolder1_GramPanchayatdrpdwn"
SEL_SUBMIT = "#ContentPlaceHolder1_btnSubmit"
SEL_UPDATE_PANEL = "#ContentPlaceHolder1_updatepanel"
SEL_TAB_CONTESTING_SARPANCH = "#tabs-1"
SEL_EXPORT_LINK = "a[href*='btnExport_repeatorSarpanch']"


class ParseStructureError(RuntimeError):
    """Raised when the returned portal table structure deviates from the expected 17 columns."""


class _HtmlTableParser(html.parser.HTMLParser):
    """Minimal dependency-free parser for verifying exported .xls (HTML table) parity."""

    def __init__(self) -> None:
        super().__init__()
        self.in_td = False
        self.in_th = False
        self.in_tr = False
        self.current_row: List[str] = []
        self.current_cell: List[str] = []
        self.is_header_row = False
        self.data_rows: List[List[str]] = []

    def handle_starttag(self, tag: str, attrs: List[Tuple[str, Optional[str]]]) -> None:
        tag = tag.lower()
        if tag == "tr":
            self.in_tr = True
            self.current_row = []
            self.is_header_row = False
        elif tag == "td" and self.in_tr:
            self.in_td = True
            self.current_cell = []
        elif tag == "th" and self.in_tr:
            self.in_th = True
            self.is_header_row = True

    def handle_data(self, data: str) -> None:
        if self.in_td:
            self.current_cell.append(data)

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if tag == "td" and self.in_td:
            self.in_td = False
            self.current_row.append(clean_text("".join(self.current_cell)))
        elif tag == "th" and self.in_th:
            self.in_th = False
        elif tag == "tr" and self.in_tr:
            self.in_tr = False
            if not self.is_header_row and len(self.current_row) == 17:
                self.data_rows.append(self.current_row)


def _is_placeholder_option(value: str, text: str) -> bool:
    v = value.strip()
    t = text.strip().lower()
    if v in ("", "0", "-1"):
        return True
    if t.startswith("--") or "please select" in t or t in ("select", "all", "--all--"):
        return True
    return False


class SecRajasthanScraper:
    """Automated Playwright scraper for Rajasthan SEC Panchayat Samiti-wise Sarpanch data."""

    def __init__(self, config: Dict[str, Any], base_dir: Path, logger: logging.Logger) -> None:
        self.config = config
        self.base_dir = base_dir
        self.logger = logger

        paths = config["paths"]
        self.output_dir = base_dir / paths["output_dir"]
        self.ps_data_dir = base_dir / paths["ps_data_dir"]
        self.logs_dir = base_dir / paths["logs_dir"]
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.ps_data_dir.mkdir(parents=True, exist_ok=True)
        self.logs_dir.mkdir(parents=True, exist_ok=True)

        self.tracker = ProgressTracker(
            catalog_path=base_dir / paths["catalog_file"],
            progress_json_path=base_dir / paths["progress_json"],
            progress_csv_path=base_dir / paths["progress_csv"],
            skip_list=config.get("skip_panchayat_samitis", []),
        )

        self.target_url: str = config["target_url"]
        self.election_type_val: str = str(config["election_type_value"])
        self.election_period_val: str = str(config["election_period_value"])

        bcfg = config["browser"]
        self.headless: bool = bool(bcfg.get("headless", True))
        self.browser_channel: Optional[str] = bcfg.get("channel")
        self.nav_timeout_ms: int = int(bcfg.get("navigation_timeout_ms", 45000))
        self.action_timeout_ms: int = int(bcfg.get("action_timeout_ms", 30000))
        self.postback_timeout_ms: int = int(bcfg.get("postback_timeout_ms", 45000))
        self.block_media: bool = bool(bcfg.get("block_media_resources", True))

        scfg = config["scraping"]
        self.min_delay: float = float(scfg.get("min_delay_seconds", 1.5))
        self.max_delay: float = float(scfg.get("max_delay_seconds", 2.5))
        self.district_delay: float = float(scfg.get("district_delay_seconds", 1.0))
        self.max_retries: int = int(scfg.get("max_retries", 3))
        self.retry_backoff: List[float] = [float(x) for x in scfg.get("retry_backoff_seconds", [3, 8, 18])]
        self.reset_every_n_ps: int = int(scfg.get("reset_page_every_n_ps", 5))
        self.verify_export_parity_sample: bool = bool(scfg.get("verify_export_parity_sample", True))

        self._parity_verified = False

    def _polite_sleep(self, fixed_seconds: Optional[float] = None) -> None:
        delay = fixed_seconds if fixed_seconds is not None else random.uniform(self.min_delay, self.max_delay)
        if delay > 0:
            time.sleep(delay)

    def _launch_browser(self, playwright: Any) -> Browser:
        """Launch Playwright Chromium, falling back cleanly to system Chrome if bundled revision is missing."""
        if self.browser_channel:
            return playwright.chromium.launch(headless=self.headless, channel=self.browser_channel)
        try:
            return playwright.chromium.launch(headless=self.headless)
        except Exception as exc:
            if "Executable doesn't exist" in str(exc):
                self.logger.info(
                    "Bundled Playwright Chromium binary not found; falling back to system Chrome (channel='chrome')."
                )
                return playwright.chromium.launch(headless=self.headless, channel="chrome")
            raise

    def _new_context_and_page(self, browser: Browser) -> Tuple[BrowserContext, Page]:
        context = browser.new_context(
            accept_downloads=True,
            viewport={"width": 1366, "height": 900},
            user_agent=(
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/128.0.0.0 Safari/537.36"
            ),
        )
        page = context.new_page()
        page.set_default_navigation_timeout(self.nav_timeout_ms)
        page.set_default_timeout(self.action_timeout_ms)
        if self.block_media:
            page.route(
                "**/*",
                lambda route: route.abort()
                if route.request.resource_type in ("image", "media", "font")
                else route.continue_(),
            )
        return context, page

    def _is_sec_postback(self, response: Any) -> bool:
        return (
            "GramPanchayatDetails.aspx" in response.url
            and response.request.method == "POST"
        )

    def _load_clean_base_page(self, page: Page) -> None:
        """Load GramPanchayatDetails.aspx freshly and set Election Type + Election Period."""
        resp = page.goto(self.target_url, wait_until="domcontentloaded", timeout=self.nav_timeout_ms)
        if resp is None or resp.status >= 400:
            raise RuntimeError(f"Failed to load base page: HTTP {resp.status if resp else 'None'}")

        page.wait_for_selector(SEL_ELECTION_TYPE, state="visible", timeout=self.action_timeout_ms)

        # Ensure Election Type = General Election (1)
        curr_type = page.locator(SEL_ELECTION_TYPE).input_value()
        if curr_type != self.election_type_val:
            with page.expect_response(self._is_sec_postback, timeout=self.postback_timeout_ms) as rinfo:
                page.select_option(SEL_ELECTION_TYPE, value=self.election_type_val)
            if rinfo.value.status >= 400:
                raise RuntimeError(f"Election Type postback failed: HTTP {rinfo.value.status}")
            page.wait_for_load_state("domcontentloaded")
            page.wait_for_selector(SEL_ELECTION_PERIOD, state="visible")

        # Select Election Duration = OCT-NOV 2026 (64)
        curr_period = page.locator(SEL_ELECTION_PERIOD).input_value()
        if curr_period != self.election_period_val:
            with page.expect_response(self._is_sec_postback, timeout=self.postback_timeout_ms) as rinfo:
                page.select_option(SEL_ELECTION_PERIOD, value=self.election_period_val)
            if rinfo.value.status >= 400:
                raise RuntimeError(f"Election Period postback failed: HTTP {rinfo.value.status}")
            page.wait_for_load_state("domcontentloaded")

        # Wait until District dropdown has options populated
        page.wait_for_function(
            f"() => {{ const s = document.querySelector('{SEL_DISTRICT}'); return s && s.options.length > 5; }}",
            timeout=self.postback_timeout_ms,
        )

    def _read_select_options(self, page: Page, selector: str) -> List[Dict[str, str]]:
        raw = page.eval_on_selector(
            selector,
            """(el) => Array.from(el.options).map(o => ({
                value: (o.value || '').trim(),
                text: (o.textContent || '').replace(/\\u00a0/g, ' ').trim()
            }))""",
        )
        return [opt for opt in raw if not _is_placeholder_option(opt["value"], opt["text"])]

    def _select_district(self, page: Page, district_id: str) -> List[Dict[str, str]]:
        """Select District and wait for UpdatePanel delta POST to refresh PS dropdown."""
        if not page.locator(SEL_DISTRICT).is_visible():
            self.logger.debug("District dropdown not visible (accordion collapsed); reloading clean base page.")
            self._load_clean_base_page(page)

        # Snapshot current PS dropdown HTML to ensure we detect an actual DOM refresh
        prev_marker = page.evaluate(
            f"""() => {{
                const ps = document.querySelector('{SEL_PANCHAYAT_SAMITI}');
                if (ps) ps.setAttribute('data-stale-check', '1');
                return ps ? ps.options.length : 0;
            }}"""
        )
        _ = prev_marker

        with page.expect_response(self._is_sec_postback, timeout=self.postback_timeout_ms) as rinfo:
            page.select_option(SEL_DISTRICT, value=str(district_id))
        resp = rinfo.value
        if resp.status >= 400:
            raise RuntimeError(f"District {district_id} postback failed: HTTP {resp.status}")

        # Wait until PS select has been replaced (marker gone) and options loaded
        page.wait_for_function(
            f"""() => {{
                const d = document.querySelector('{SEL_DISTRICT}');
                const ps = document.querySelector('{SEL_PANCHAYAT_SAMITI}');
                return d && d.value === '{district_id}' && ps && !ps.hasAttribute('data-stale-check') && ps.options.length >= 1;
            }}""",
            timeout=self.postback_timeout_ms,
        )
        # Small settle delay for MicrosoftAjax PageRequestManager endRequest
        page.wait_for_timeout(150)
        return self._read_select_options(page, SEL_PANCHAYAT_SAMITI)

    def _select_ps_and_submit(
        self,
        page: Page,
        district_id: str,
        ps_id: str,
    ) -> Tuple[List[str], Dict[str, Any]]:
        """Select Panchayat Samiti (full POST), keep GP=--All--, click Submit, and parse table."""
        if not page.locator(SEL_PANCHAYAT_SAMITI).is_visible():
            self.logger.debug("PS dropdown hidden; reloading clean base page and re-selecting district %s.", district_id)
            self._load_clean_base_page(page)
            self._select_district(page, district_id)

        # Ensure expected PS option exists in dropdown
        has_ps = page.evaluate(
            f"""() => {{
                const ps = document.querySelector('{SEL_PANCHAYAT_SAMITI}');
                return !!ps && Array.from(ps.options).some(o => o.value.trim() === '{ps_id}');
            }}"""
        )
        if not has_ps:
            self._load_clean_base_page(page)
            self._select_district(page, district_id)

        # Step 1: Select Panchayat Samiti (triggers full synchronous POST that populates GramPanchayatdrpdwn)
        with page.expect_response(self._is_sec_postback, timeout=self.postback_timeout_ms) as rinfo:
            page.select_option(SEL_PANCHAYAT_SAMITI, value=str(ps_id))
        if rinfo.value.status >= 400:
            raise RuntimeError(f"PS {ps_id} selection postback failed: HTTP {rinfo.value.status}")
        page.wait_for_load_state("domcontentloaded")
        page.wait_for_selector(SEL_GRAM_PANCHAYAT, state="visible", timeout=self.postback_timeout_ms)

        # Verify PS is actually selected
        selected_ps = page.locator(SEL_PANCHAYAT_SAMITI).input_value()
        if str(selected_ps).strip() != str(ps_id).strip():
            raise RuntimeError(f"PS selection mismatch: expected {ps_id}, got {selected_ps}")

        # Capture all Gram Panchayat names listed in the dropdown for completeness verification
        gp_options = self._read_select_options(page, SEL_GRAM_PANCHAYAT)
        gp_dropdown_names = [opt["text"] for opt in gp_options]

        # Ensure Gram Panchayat is set to '0' (--All--)
        curr_gp = page.locator(SEL_GRAM_PANCHAYAT).input_value()
        if curr_gp != "0":
            has_all = page.evaluate(
                f"() => Array.from(document.querySelector('{SEL_GRAM_PANCHAYAT}').options).some(o => o.value === '0')"
            )
            if has_all:
                page.select_option(SEL_GRAM_PANCHAYAT, value="0")

        # Step 2: Click Submit (full synchronous POST returning complete HTML table without pagination)
        with page.expect_response(self._is_sec_postback, timeout=self.postback_timeout_ms) as rinfo:
            page.click(SEL_SUBMIT)
        if rinfo.value.status >= 400:
            raise RuntimeError(f"Submit postback for PS {ps_id} failed: HTTP {rinfo.value.status}")
        page.wait_for_load_state("domcontentloaded")
        page.wait_for_selector(SEL_UPDATE_PANEL, state="attached", timeout=self.postback_timeout_ms)

        extracted = self._extract_contesting_sarpanch_table(page)
        return gp_dropdown_names, extracted

    def _extract_contesting_sarpanch_table(self, page: Page) -> Dict[str, Any]:
        """Extract all 17 columns from #tabs-1 table.Nomination strictly by <td> index."""
        payload = page.evaluate(
            """() => {
                const panel = document.querySelector('#ContentPlaceHolder1_updatepanel');
                if (!panel) {
                    return { panel_found: false, table_found: false, headers: [], rows: [], bad_rows: [], message: 'updatepanel not found' };
                }
                const tab1 = panel.querySelector('#tabs-1') || document.querySelector('#tabs-1') || panel;
                const table = tab1.querySelector('table.Nomination');
                if (!table) {
                    const txt = (tab1.textContent || '').replace(/\\s+/g, ' ').trim().slice(0, 300);
                    return { panel_found: true, table_found: false, headers: [], rows: [], bad_rows: [], message: txt };
                }

                const headerCells = Array.from(table.querySelectorAll('thead th')).map(
                    th => (th.textContent || '').replace(/\\u00a0/g, ' ').replace(/\\s+/g, ' ').trim()
                );

                // Parse every tr that has td cells (skip thead header rows)
                const trs = Array.from(table.querySelectorAll('tbody tr'));
                const allTrs = trs.length > 0 ? trs : Array.from(table.querySelectorAll('tr'));
                const rows = [];
                const bad_rows = [];
                let message = '';

                for (const tr of allTrs) {
                    const tds = Array.from(tr.querySelectorAll('td'));
                    if (tds.length === 0) continue;
                    const vals = tds.map(td => (td.textContent || '').replace(/\\u00a0/g, ' ').replace(/\\s+/g, ' ').trim());
                    if (vals.length === 17) {
                        rows.push(vals);
                    } else if (vals.length === 1 && /no\\s+record|no\\s+data|not\\s+found/i.test(vals[0])) {
                        message = vals[0];
                    } else {
                        bad_rows.push({ col_count: vals.length, sample: vals.slice(0, 5) });
                    }
                }

                return {
                    panel_found: true,
                    table_found: true,
                    headers: headerCells,
                    rows: rows,
                    bad_rows: bad_rows,
                    message: message,
                };
            }"""
        )

        if not payload["panel_found"]:
            raise RuntimeError("Result container #ContentPlaceHolder1_updatepanel missing after Submit.")

        if payload["bad_rows"]:
            raise ParseStructureError(
                f"Unexpected column counts in table rows: {payload['bad_rows'][:3]}"
            )

        if payload["table_found"] and payload["rows"]:
            header_joined = " ".join(payload["headers"]).lower()
            expected_tokens = ["grampanchayat", "contesting candidate", "capital assets", "mobile"]
            missing_tokens = [tok for tok in expected_tokens if tok not in header_joined]
            if missing_tokens:
                raise ParseStructureError(
                    f"Table headers do not match Contesting Sarpanch schema; missing tokens {missing_tokens}. "
                    f"Headers seen: {payload['headers']}"
                )

        return payload

    def _verify_export_parity(self, page: Page, dom_rows: List[List[str]], ps_name: str) -> bool:
        """Click Export to Excel once during smoke test and verify 100% parity with DOM rows."""
        export_btn = page.locator(SEL_EXPORT_LINK)
        if export_btn.count() == 0:
            self.logger.warning("Parity check skipped for %s: export link not found on page.", ps_name)
            return False

        parity_file = self.logs_dir / f"parity_sample_{sanitize_filename(ps_name)}.xls"
        with page.expect_download(timeout=self.postback_timeout_ms) as dl_info:
            export_btn.first.click()
        dl_info.value.save_as(str(parity_file))

        raw_html = parity_file.read_text(encoding="utf-8", errors="replace")
        parser = _HtmlTableParser()
        parser.feed(raw_html)
        xls_rows = parser.data_rows

        if len(xls_rows) != len(dom_rows):
            self.logger.error(
                "EXPORT PARITY MISMATCH for %s: DOM=%d rows vs Export .xls=%d rows!",
                ps_name,
                len(dom_rows),
                len(xls_rows),
            )
            return False

        mismatches = 0
        for idx, (d_row, x_row) in enumerate(zip(dom_rows, xls_rows), start=1):
            if d_row != x_row:
                mismatches += 1
                if mismatches <= 3:
                    self.logger.warning(
                        "Parity cell diff at row %d for %s:\n  DOM: %s\n  XLS: %s",
                        idx,
                        ps_name,
                        d_row,
                        x_row,
                    )

        if mismatches == 0:
            self.logger.info(
                "EXPORT PARITY VERIFIED (100%% exact cell-for-cell match on %d rows) for %s (%s).",
                len(dom_rows),
                ps_name,
                parity_file.name,
            )
            return True

        self.logger.warning("Export parity had %d differing rows for %s.", mismatches, ps_name)
        return False

    def _audit_ps_rows(
        self,
        raw_rows: List[List[str]],
        gp_dropdown_names: List[str],
    ) -> Dict[str, Any]:
        """Compute non-destructive quality & completeness metrics for one Panchayat Samiti."""
        gps_in_table = {clean_text(row[1]) for row in raw_rows if clean_text(row[1])}
        gps_in_table_lower = {g.lower() for g in gps_in_table}
        missing_gps = [clean_text(g) for g in gp_dropdown_names if clean_text(g).lower() not in gps_in_table_lower]

        blank_field_rows = sum(1 for row in raw_rows if any(cell == "" for cell in row))

        # Check duplicate Candidate Serial No. within same GP (reported only; never dropped)
        gp_serial_counts: Counter[Tuple[str, str]] = Counter()
        for row in raw_rows:
            gp_serial_counts[(row[1], row[3])] += 1
        dup_serial_in_gp = sum(1 for cnt in gp_serial_counts.values() if cnt > 1)

        # Check exact full-row duplicates (excluding Sr. No. at col 0)
        row_sig_counts: Counter[Tuple[str, ...]] = Counter(tuple(row[1:]) for row in raw_rows)
        exact_dup_rows = sum(cnt - 1 for cnt in row_sig_counts.values() if cnt > 1)

        return {
            "gp_dropdown_count": len(gp_dropdown_names),
            "gp_in_table_count": len(gps_in_table),
            "missing_gps_count": len(missing_gps),
            "missing_gps_in_table": missing_gps,
            "blank_field_rows_count": blank_field_rows,
            "duplicate_serial_in_gp_count": dup_serial_in_gp,
            "exact_duplicate_rows_count": exact_dup_rows,
        }

    def _save_ps_csv(
        self,
        district_id: str,
        district_name: str,
        ps_id: str,
        ps_name: str,
        raw_rows: List[List[str]],
    ) -> Path:
        filename = f"{district_id.zfill(2)}_{sanitize_filename(district_name)}__{ps_id.zfill(3)}_{sanitize_filename(ps_name)}.csv"
        out_path = self.ps_data_dir / filename
        columns = PORTAL_FIELD_KEYS + ["district_id", "district", "ps_id", "panchayat_samiti"]
        dict_rows: List[Dict[str, Any]] = []
        for row in raw_rows:
            rec = {PORTAL_FIELD_KEYS[i]: row[i] for i in range(17)}
            rec["district_id"] = district_id
            rec["district"] = district_name
            rec["ps_id"] = ps_id
            rec["panchayat_samiti"] = ps_name
            dict_rows.append(rec)
        atomic_write_csv(out_path, columns, dict_rows)
        return out_path

    def discover_all(self, force_refresh: bool = False) -> Dict[str, Any]:
        """Enumerate all Districts and all Panchayat Samitis dynamically from the live SEC portal."""
        if self.tracker.has_catalog() and not force_refresh:
            counts = self.tracker.summary_counts()
            self.logger.info(
                "Using existing catalog (%d districts, %d Panchayat Samitis).",
                self.tracker.state["meta"].get("total_districts", 0),
                counts["total"],
            )
            return counts

        self.logger.info("Starting statewide District & Panchayat Samiti discovery on %s ...", self.target_url)
        districts: List[Dict[str, str]] = []
        panchayat_samitis: List[Dict[str, str]] = []

        with sync_playwright() as p:
            browser = self._launch_browser(p)
            context, page = self._new_context_and_page(browser)
            try:
                self._load_clean_base_page(page)
                districts = self._read_select_options(page, SEL_DISTRICT)
                self.logger.info("Discovered %d Districts on SEC portal.", len(districts))

                for idx, dist in enumerate(districts, start=1):
                    d_id = dist["value"]
                    d_name = dist["text"]
                    ps_list: List[Dict[str, str]] = []
                    for attempt in range(1, self.max_retries + 1):
                        try:
                            ps_list = self._select_district(page, d_id)
                            break
                        except Exception as exc:
                            self.logger.warning(
                                "Discovery retry %d/%d for district %s (%s): %s",
                                attempt,
                                self.max_retries,
                                d_name,
                                d_id,
                                exc,
                            )
                            if attempt == self.max_retries:
                                raise
                            time.sleep(self.retry_backoff[min(attempt - 1, len(self.retry_backoff) - 1)])
                            self._load_clean_base_page(page)

                    self.logger.info(
                        "  [%02d/%02d] %-20s (id=%2s): %2d Panchayat Samitis",
                        idx,
                        len(districts),
                        d_name,
                        d_id,
                        len(ps_list),
                    )
                    for ps in ps_list:
                        panchayat_samitis.append(
                            {
                                "district_id": d_id,
                                "district_name": d_name,
                                "ps_id": ps["value"],
                                "ps_name": ps["text"],
                            }
                        )
                    self._polite_sleep(self.district_delay)
            finally:
                context.close()
                browser.close()

        self.tracker.save_catalog_and_init_progress(districts, panchayat_samitis)
        counts = self.tracker.summary_counts()
        self.logger.info(
            "Discovery complete: %d Districts, %d total Panchayat Samitis (%d skipped AMBER, %d pending).",
            len(districts),
            counts["total"],
            counts["skipped_amber"],
            counts["pending"],
        )
        return counts

    def scrape_items(
        self,
        items: List[Dict[str, Any]],
        run_export_parity: bool = False,
    ) -> Dict[str, int]:
        """Scrape a list of Panchayat Samitis grouped by District, with retries and checkpointing."""
        if not items:
            self.logger.info("No actionable Panchayat Samitis to scrape.")
            return self.tracker.summary_counts()

        # Group items by District so we minimize District reloads while keeping ViewState clean
        by_district: Dict[Tuple[str, str], List[Dict[str, Any]]] = defaultdict(list)
        for item in items:
            by_district[(item["district_id"], item["district_name"])].append(item)

        total_to_run = len(items)
        processed = 0
        run_start = time.monotonic()

        with sync_playwright() as p:
            browser = self._launch_browser(p)
            context, page = self._new_context_and_page(browser)
            try:
                for (d_id, d_name), dist_items in sorted(by_district.items(), key=lambda kv: kv[0][1]):
                    self.logger.info(
                        "=== District: %s (id=%s) — %d Panchayat Samitis queued ===",
                        d_name,
                        d_id,
                        len(dist_items),
                    )
                    ps_since_reset = 0
                    district_ready = False

                    for item in dist_items:
                        processed += 1
                        ps_key = item["ps_key"]
                        ps_id = item["ps_id"]
                        ps_name = item["ps_name"]
                        ps_start = time.monotonic()

                        # Double-check AMBER skip safeguard
                        if self.tracker.is_skipped_amber(d_id, ps_id, ps_name):
                            self.logger.info("[%d/%d] Skipping AMBER: %s", processed, total_to_run, ps_name)
                            self.tracker.update_item(ps_key, status="skipped_amber")
                            continue

                        last_error = ""
                        succeeded = False
                        attempts_done = int(item.get("attempts") or 0)

                        for attempt in range(1, self.max_retries + 1):
                            attempts_done += 1
                            try:
                                # Keep ASP.NET ViewState lean and avoid accordion drift
                                if (not district_ready) or (ps_since_reset >= self.reset_every_n_ps) or (attempt > 1):
                                    self._load_clean_base_page(page)
                                    self._select_district(page, d_id)
                                    district_ready = True
                                    ps_since_reset = 0

                                gp_dropdown_names, extracted = self._select_ps_and_submit(page, d_id, ps_id)
                                raw_rows = extracted["rows"]
                                ps_since_reset += 1

                                # Optional one-time parity check against portal's Export to Excel (.xls)
                                if run_export_parity and not self._parity_verified and raw_rows:
                                    self._parity_verified = self._verify_export_parity(page, raw_rows, ps_name)
                                    # Reload cleanly after clicking Export LinkButton
                                    district_ready = False

                                audit = self._audit_ps_rows(raw_rows, gp_dropdown_names)
                                elapsed = round(time.monotonic() - ps_start, 2)

                                if len(raw_rows) == 0:
                                    msg = extracted.get("message") or "No contesting candidate table/rows returned by portal"
                                    self.tracker.update_item(
                                        ps_key,
                                        status="empty_on_portal",
                                        attempts=attempts_done,
                                        record_count=0,
                                        portal_message=msg,
                                        csv_path="",
                                        duration_sec=elapsed,
                                        error_message="",
                                        **audit,
                                    )
                                    self.logger.warning(
                                        "[%d/%d] EMPTY ON PORTAL: %s -> %s (id=%s) | GPs in dropdown=%d | msg=%s (%.1fs)",
                                        processed,
                                        total_to_run,
                                        d_name,
                                        ps_name,
                                        ps_id,
                                        audit["gp_dropdown_count"],
                                        msg[:120],
                                        elapsed,
                                    )
                                else:
                                    csv_path = self._save_ps_csv(d_id, d_name, ps_id, ps_name, raw_rows)
                                    rel_csv = str(csv_path.relative_to(self.base_dir))
                                    self.tracker.update_item(
                                        ps_key,
                                        status="completed",
                                        attempts=attempts_done,
                                        record_count=len(raw_rows),
                                        portal_message="",
                                        csv_path=rel_csv,
                                        duration_sec=elapsed,
                                        error_message="",
                                        **audit,
                                    )
                                    self.logger.info(
                                        "[%d/%d] OK: %s -> %s (id=%s) | rows=%d | GPs=%d/%d | %.1fs",
                                        processed,
                                        total_to_run,
                                        d_name,
                                        ps_name,
                                        ps_id,
                                        len(raw_rows),
                                        audit["gp_in_table_count"],
                                        audit["gp_dropdown_count"],
                                        elapsed,
                                    )
                                succeeded = True
                                break

                            except (PlaywrightTimeoutError, ParseStructureError, Exception) as exc:
                                last_error = f"{type(exc).__name__}: {exc}"
                                district_ready = False
                                self.logger.warning(
                                    "[%d/%d] Attempt %d/%d FAILED for %s -> %s (id=%s): %s",
                                    processed,
                                    total_to_run,
                                    attempt,
                                    self.max_retries,
                                    d_name,
                                    ps_name,
                                    ps_id,
                                    last_error,
                                )
                                if attempt < self.max_retries:
                                    backoff = self.retry_backoff[min(attempt - 1, len(self.retry_backoff) - 1)]
                                    self._polite_sleep(backoff)
                                    # Recreate page if browser page got into a bad state
                                    try:
                                        page.close()
                                    except Exception:
                                        pass
                                    page = context.new_page()
                                    page.set_default_navigation_timeout(self.nav_timeout_ms)
                                    page.set_default_timeout(self.action_timeout_ms)
                                    if self.block_media:
                                        page.route(
                                            "**/*",
                                            lambda route: route.abort()
                                            if route.request.resource_type in ("image", "media", "font")
                                            else route.continue_(),
                                        )

                        if not succeeded:
                            elapsed = round(time.monotonic() - ps_start, 2)
                            self.tracker.update_item(
                                ps_key,
                                status="failed",
                                attempts=attempts_done,
                                duration_sec=elapsed,
                                error_message=last_error,
                            )
                            self.logger.error(
                                "[%d/%d] FAILED PERMANENTLY (for this run): %s -> %s (id=%s): %s",
                                processed,
                                total_to_run,
                                d_name,
                                ps_name,
                                ps_id,
                                last_error,
                            )

                        # Periodic progress heartbeat every 20 PSs
                        if processed % 20 == 0 or processed == total_to_run:
                            c = self.tracker.summary_counts()
                            mins = (time.monotonic() - run_start) / 60.0
                            self.logger.info(
                                "--- Progress Heartbeat [%.1f min]: completed=%d, skipped_amber=%d, empty=%d, failed=%d, pending=%d, total_rows=%d ---",
                                mins,
                                c["completed"],
                                c["skipped_amber"],
                                c["empty_on_portal"],
                                c["failed"],
                                c["pending"],
                                c["total_records_including_amber"],
                            )

                        self._polite_sleep()
            finally:
                context.close()
                browser.close()

        return self.tracker.summary_counts()

    def run_smoke_test(self, count: int = 4) -> Dict[str, int]:
        """Run discovery (if needed) and test extraction on `count` PSs across at least 2 districts."""
        self.discover_all(force_refresh=False)
        pending = self.tracker.get_actionable_items(include_failed=True, include_empty=False)
        if not pending:
            self.logger.info("No pending items left for smoke test.")
            return self.tracker.summary_counts()

        # Pick 2 PSs from the first district and 2 PSs from the second district (to test both intra-district & cross-district flows)
        selected: List[Dict[str, Any]] = []
        by_dist: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
        for item in pending:
            by_dist[item["district_name"]].append(item)

        dist_names = sorted(by_dist.keys())
        # Take up to 2 from first district, then up to 2 from second district, etc.
        for d in dist_names:
            for item in by_dist[d][:2]:
                selected.append(item)
                if len(selected) >= count:
                    break
            if len(selected) >= count:
                break

        self.logger.info(
            "Running smoke test on %d Panchayat Samitis: %s",
            len(selected),
            [(x["district_name"], x["ps_name"], x["ps_id"]) for x in selected],
        )
        return self.scrape_items(selected, run_export_parity=self.verify_export_parity_sample)

    def run_full(
        self,
        max_items: Optional[int] = None,
        retry_failed: bool = True,
        retry_empty: bool = False,
        force_rediscover: bool = False,
    ) -> Dict[str, int]:
        """Execute full resumable statewide scraping across all remaining Panchayat Samitis."""
        self.discover_all(force_refresh=force_rediscover)
        items = self.tracker.get_actionable_items(
            include_failed=retry_failed,
            include_empty=retry_empty,
        )
        if max_items is not None:
            items = items[:max_items]
        self.logger.info("Queued %d Panchayat Samitis for extraction.", len(items))
        return self.scrape_items(items, run_export_parity=self.verify_export_parity_sample)
