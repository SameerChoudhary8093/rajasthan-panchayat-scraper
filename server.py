"""One-page local web UI for exploring and downloading the scraped Rajasthan Sarpanch data.

Run:
    python3 server.py            # then open http://127.0.0.1:8000

Uses only the Python standard library for the web server; data loading and Excel export reuse
excel_builder.py, and the optional "Refresh from SEC portal" button reuses scraper.py (Playwright).
"""

from __future__ import annotations

import argparse
import csv
import io
import json
import logging
import mimetypes
import re
import threading
import time
import urllib.parse
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl

from excel_builder import CANDIDATE_SHEET_LAYOUT, load_amber_excel_untouched, load_scraped_ps_csv, write_candidates_sheet
from progress_tracker import ProgressTracker, make_ps_key

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
MAX_PAGE_SIZE = 500

logger = logging.getLogger("raj_sec_ui")


def _slug(text: str) -> str:
    return re.sub(r"[^A-Za-z0-9]+", "_", text).strip("_") or "All"


class DataStore:
    """In-memory index of AMBER + all scraped Panchayat Samiti records."""

    def __init__(self, config: Dict[str, Any]) -> None:
        self.config = config
        self.lock = threading.RLock()
        self.items: List[Dict[str, Any]] = []
        self.records_by_ps: Dict[str, List[Dict[str, Any]]] = {}
        self.loaded_at = ""

    def _tracker(self) -> ProgressTracker:
        paths = self.config["paths"]
        return ProgressTracker(
            catalog_path=BASE_DIR / paths["catalog_file"],
            progress_json_path=BASE_DIR / paths["progress_json"],
            progress_csv_path=BASE_DIR / paths["progress_csv"],
            skip_list=self.config.get("skip_panchayat_samitis", []),
        )

    def _load_item_records(self, item: Dict[str, Any]) -> List[Dict[str, Any]]:
        status = item.get("status")
        if status == "skipped_amber":
            recs, _ = load_amber_excel_untouched(
                BASE_DIR / self.config["amber_excel_path"], item["district_name"], item["ps_name"]
            )
        elif item.get("csv_path") and (BASE_DIR / item["csv_path"]).exists():
            # Show the last successfully scraped data even if a later refresh attempt failed
            recs = load_scraped_ps_csv(BASE_DIR / item["csv_path"])
        else:
            recs = []
        for r in recs:
            r["district_id"] = str(item["district_id"])
            r["ps_id"] = str(item["ps_id"])
        return recs

    def load(self) -> None:
        start = time.monotonic()
        items = self._tracker().get_all_items()
        by_ps = {item["ps_key"]: self._load_item_records(item) for item in items}
        with self.lock:
            self.items = items
            self.records_by_ps = by_ps
            self.loaded_at = time.strftime("%Y-%m-%d %H:%M:%S")
        logger.info(
            "Loaded %d records across %d Panchayat Samitis in %.1fs",
            sum(len(v) for v in by_ps.values()),
            len(items),
            time.monotonic() - start,
        )

    def reload_ps(self, ps_key: str) -> None:
        items = self._tracker().get_all_items()
        item = next(i for i in items if i["ps_key"] == ps_key)
        recs = self._load_item_records(item)
        with self.lock:
            self.items = items
            self.records_by_ps[ps_key] = recs

    # ---- Queries -------------------------------------------------------
    def names(self, district_id: str, ps_id: str) -> Tuple[str, str]:
        with self.lock:
            d_name = next((it["district_name"] for it in self.items if district_id and it["district_id"] == district_id), "")
            ps_name = next(
                (it["ps_name"] for it in self.items
                 if ps_id and it["ps_id"] == ps_id and (not district_id or it["district_id"] == district_id)),
                "",
            )
            return d_name, ps_name

    def summary(self) -> Dict[str, Any]:
        with self.lock:
            all_recs = [r for recs in self.records_by_ps.values() for r in recs]
            statuses: Dict[str, int] = {}
            for it in self.items:
                statuses[it.get("status", "pending")] = statuses.get(it.get("status", "pending"), 0) + 1
            return {
                "districts": len({it["district_id"] for it in self.items}),
                "panchayat_samitis": len(self.items),
                "gram_panchayats": len({(r["ps_id"], r.get("gram_panchayat_name")) for r in all_recs}),
                "candidates": len(all_recs),
                "status_counts": statuses,
                "loaded_at": self.loaded_at,
            }

    def districts(self) -> List[Dict[str, Any]]:
        with self.lock:
            out: Dict[str, Dict[str, Any]] = {}
            for it in self.items:
                d = out.setdefault(
                    it["district_id"],
                    {"id": it["district_id"], "name": it["district_name"], "ps_count": 0, "candidates": 0},
                )
                d["ps_count"] += 1
                d["candidates"] += len(self.records_by_ps.get(it["ps_key"], []))
            return sorted(out.values(), key=lambda d: d["name"])

    def panchayat_samitis(self, district_id: str) -> List[Dict[str, Any]]:
        with self.lock:
            out = []
            for it in self.items:
                if it["district_id"] != district_id:
                    continue
                recs = self.records_by_ps.get(it["ps_key"], [])
                out.append(
                    {
                        "id": it["ps_id"],
                        "name": it["ps_name"],
                        "status": it.get("status"),
                        "candidates": len(recs),
                        "gram_panchayats": len({r.get("gram_panchayat_name") for r in recs}),
                        "last_updated": it.get("last_updated", ""),
                    }
                )
            return sorted(out, key=lambda p: p["name"])

    def gram_panchayats(self, district_id: str, ps_id: str) -> List[Dict[str, Any]]:
        with self.lock:
            recs = self.records_by_ps.get(make_ps_key(district_id, ps_id), [])
            counts: Dict[str, int] = {}
            for r in recs:
                gp = r.get("gram_panchayat_name") or ""
                counts[gp] = counts.get(gp, 0) + 1
            return [{"name": k, "candidates": v} for k, v in sorted(counts.items(), key=lambda kv: kv[0].lower())]

    def filter(self, district_id: str = "", ps_id: str = "", gp: str = "", q: str = "") -> List[Dict[str, Any]]:
        q_low = q.strip().lower()
        with self.lock:
            out: List[Dict[str, Any]] = []
            for it in self.items:
                if district_id and it["district_id"] != district_id:
                    continue
                if ps_id and it["ps_id"] != ps_id:
                    continue
                for r in self.records_by_ps.get(it["ps_key"], []):
                    if gp and (r.get("gram_panchayat_name") or "") != gp:
                        continue
                    if q_low and q_low not in f"{r.get('candidate_name') or ''} {r.get('father_husband_name') or ''}".lower():
                        continue
                    out.append(r)
            return out


class RefreshManager:
    """Re-scrapes one Panchayat Samiti live from the SEC portal (one at a time, politely)."""

    def __init__(self, config: Dict[str, Any], store: DataStore) -> None:
        self.config = config
        self.store = store
        self.lock = threading.Lock()

    def refresh(self, district_id: str, ps_id: str) -> Dict[str, Any]:
        if not self.lock.acquire(blocking=False):
            return {"ok": False, "message": "Another refresh is already running. Please wait for it to finish."}
        try:
            from scraper import SecRajasthanScraper  # lazy import: Playwright only needed here

            scraper = SecRajasthanScraper(config=self.config, base_dir=BASE_DIR, logger=logger)
            ps_key = make_ps_key(district_id, ps_id)
            item = scraper.tracker.state["items"].get(ps_key)
            if not item:
                return {"ok": False, "message": "Unknown Panchayat Samiti."}
            if item.get("status") == "skipped_amber":
                return {"ok": False, "message": "AMBER comes from your original Excel file and is kept untouched."}
            previous = dict(item)
            scraper.scrape_items([dict(item)], run_export_parity=False)
            updated = scraper.tracker.state["items"][ps_key]
            ok = updated.get("status") == "completed"
            if not ok and previous.get("status") == "completed":
                # A failed refresh must never discard previously verified data: restore the old state
                failure = updated.get("error_message") or updated.get("portal_message") or updated.get("status")
                if "ERR_NAME_NOT_RESOLVED" in str(failure) or "ERR_INTERNET_DISCONNECTED" in str(failure):
                    failure = "Could not reach sec.rajasthan.gov.in (check your internet connection)."
                elif "Timeout" in str(failure):
                    failure = "The SEC website took too long to respond. Please try again later."
                else:
                    failure = str(failure).splitlines()[0][:200]
                restored = dict(previous)
                restored["error_message"] = f"Last live refresh failed (old data kept): {failure}"
                scraper.tracker.state["items"][ps_key] = restored
                scraper.tracker.save()
                updated = dict(updated, error_message=failure)
            self.store.reload_ps(ps_key)
            return {
                "ok": ok,
                "status": updated.get("status"),
                "candidates": updated.get("record_count", 0),
                "message": (
                    f"Refreshed {updated['ps_name']} from the SEC portal: {updated.get('record_count', 0)} candidates."
                    if ok
                    else f"Refresh did not complete ({updated.get('status')}); previously saved data is kept. "
                    f"Details: {updated.get('error_message') or updated.get('portal_message')}"
                ),
            }
        finally:
            self.lock.release()


def records_to_xlsx(records: List[Dict[str, Any]]) -> bytes:
    wb = openpyxl.Workbook(write_only=True)
    write_candidates_sheet(wb, "Contesting Candidates", records)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def records_to_csv(records: List[Dict[str, Any]]) -> bytes:
    headers = []
    for key, h1, h2, _, _ in CANDIDATE_SHEET_LAYOUT:
        label = " - ".join(x for x in ("Children" if key.startswith("children_") else h1, h2) if x)
        headers.append(label.replace("\n", " "))
    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(headers)
    for idx, r in enumerate(records, start=1):
        writer.writerow([idx if key == "overall_sno" else ("" if r.get(key) is None else r.get(key))
                         for key, *_ in CANDIDATE_SHEET_LAYOUT])
    return ("\ufeff" + buf.getvalue()).encode("utf-8")  # BOM so Excel opens Hindi/UTF-8 text correctly


def make_handler(config: Dict[str, Any], store: DataStore, refresher: RefreshManager):
    combined_excel = BASE_DIR / config["paths"]["combined_excel"]

    class Handler(BaseHTTPRequestHandler):
        server_version = "RajSarpanchUI/1.0"

        def log_message(self, fmt: str, *args: Any) -> None:
            logger.info("%s - %s", self.address_string(), fmt % args)

        # ---- helpers ----
        def _send(self, status: int, body: bytes, ctype: str, extra: Optional[Dict[str, str]] = None) -> None:
            self.send_response(status)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(body)))
            self.send_header("X-Content-Type-Options", "nosniff")
            for k, v in (extra or {}).items():
                self.send_header(k, v)
            self.end_headers()
            self.wfile.write(body)

        def _json(self, data: Any, status: int = 200) -> None:
            self._send(status, json.dumps(data, ensure_ascii=False).encode("utf-8"), "application/json; charset=utf-8")

        def _params(self) -> Tuple[str, Dict[str, str]]:
            parsed = urllib.parse.urlparse(self.path)
            qs = {k: v[0] for k, v in urllib.parse.parse_qs(parsed.query).items()}
            return parsed.path, qs

        # ---- routes ----
        def do_GET(self) -> None:  # noqa: N802
            path, qs = self._params()
            try:
                if path in ("/", "/index.html"):
                    return self._static("index.html")
                if path.startswith("/static/"):
                    return self._static(path[len("/static/"):])
                if path == "/api/summary":
                    return self._json(store.summary())
                if path == "/api/districts":
                    return self._json(store.districts())
                if path == "/api/panchayat-samitis":
                    return self._json(store.panchayat_samitis(qs.get("district_id", "")))
                if path == "/api/gram-panchayats":
                    return self._json(store.gram_panchayats(qs.get("district_id", ""), qs.get("ps_id", "")))
                if path == "/api/records":
                    return self._records(qs)
                if path == "/api/export":
                    return self._export(qs)
                self._json({"error": "Not found"}, 404)
            except Exception as exc:  # pragma: no cover - defensive
                logger.exception("GET %s failed", path)
                self._json({"error": str(exc)}, 500)

        def do_POST(self) -> None:  # noqa: N802
            path, _ = self._params()
            if path != "/api/refresh":
                return self._json({"error": "Not found"}, 404)
            length = int(self.headers.get("Content-Length") or 0)
            try:
                payload = json.loads(self.rfile.read(length) or b"{}")
                result = refresher.refresh(str(payload.get("district_id", "")), str(payload.get("ps_id", "")))
                self._json(result, 200)
            except Exception as exc:
                logger.exception("Refresh failed")
                self._json({"ok": False, "message": f"Refresh failed: {exc}"}, 500)

        def _static(self, rel: str) -> None:
            target = (STATIC_DIR / rel).resolve()
            if STATIC_DIR.resolve() not in target.parents or not target.is_file():
                return self._json({"error": "Not found"}, 404)
            ctype = mimetypes.guess_type(target.name)[0] or "application/octet-stream"
            if ctype.startswith("text/") or ctype in ("application/javascript",):
                ctype += "; charset=utf-8"
            self._send(200, target.read_bytes(), ctype, {"Cache-Control": "no-cache"})

        def _records(self, qs: Dict[str, str]) -> None:
            rows = store.filter(qs.get("district_id", ""), qs.get("ps_id", ""), qs.get("gp", ""), qs.get("q", ""))
            page_size = max(1, min(MAX_PAGE_SIZE, int(qs.get("page_size", 100) or 100)))
            pages = max(1, -(-len(rows) // page_size))
            page = max(1, min(pages, int(qs.get("page", 1) or 1)))
            start = (page - 1) * page_size
            page_rows = []
            for i, r in enumerate(rows[start:start + page_size], start=start + 1):
                rec = {k: r.get(k) for k, *_ in CANDIDATE_SHEET_LAYOUT if k != "overall_sno"}
                rec["overall_sno"] = i
                page_rows.append(rec)
            self._json(
                {
                    "total": len(rows),
                    "gram_panchayats": len({(r["ps_id"], r.get("gram_panchayat_name")) for r in rows}),
                    "page": page,
                    "pages": pages,
                    "page_size": page_size,
                    "rows": page_rows,
                }
            )

        def _export(self, qs: Dict[str, str]) -> None:
            fmt = qs.get("format", "xlsx")
            d_id, ps_id, gp, q = qs.get("district_id", ""), qs.get("ps_id", ""), qs.get("gp", ""), qs.get("q", "")
            no_filter = not any((d_id, ps_id, gp, q))
            xlsx_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

            if fmt == "xlsx" and no_filter and combined_excel.exists():
                body = combined_excel.read_bytes()
                name = combined_excel.name
                ctype = xlsx_type
            else:
                rows = store.filter(d_id, ps_id, gp, q)
                d_name, ps_name = store.names(d_id, ps_id)
                parts = ["Sarpanch"]
                if d_name:
                    parts.append(d_name)
                if ps_name:
                    parts.append(ps_name.replace(" PANCHAYAT SAMITI", "_PS"))
                if gp:
                    parts.append(gp)
                if q:
                    parts.append(f"search_{q}")
                if len(parts) == 1:
                    parts.append("Rajasthan_All")
                base = "_".join(_slug(p) for p in parts)
                if fmt == "csv":
                    body, name, ctype = records_to_csv(rows), f"{base}.csv", "text/csv; charset=utf-8"
                else:
                    body, name, ctype = records_to_xlsx(rows), f"{base}.xlsx", xlsx_type
            self._send(200, body, ctype, {"Content-Disposition": f'attachment; filename="{name}"'})

    return Handler


def main() -> int:
    parser = argparse.ArgumentParser(description="Local web UI for Rajasthan Sarpanch candidate data")
    parser.add_argument("--host", default="127.0.0.1", help="Bind address (default: 127.0.0.1, local only)")
    parser.add_argument("--port", type=int, default=8000)
    parser.add_argument("--config", default="config.json")
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S")
    config = json.loads((BASE_DIR / args.config).read_text(encoding="utf-8"))
    store = DataStore(config)
    store.load()
    refresher = RefreshManager(config, store)

    httpd = ThreadingHTTPServer((args.host, args.port), make_handler(config, store, refresher))
    logger.info("Open http://%s:%d in your browser (Ctrl+C to stop)", args.host, args.port)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        httpd.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
