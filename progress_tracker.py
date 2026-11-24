"""Atomic, resumable progress and catalog tracking for Panchayat Samiti scraping."""

from __future__ import annotations

import csv
import json
import os
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple


PROGRESS_CSV_COLUMNS = [
    "ps_key",
    "district_id",
    "district_name",
    "ps_id",
    "ps_name",
    "status",
    "attempts",
    "gp_dropdown_count",
    "gp_in_table_count",
    "missing_gps_count",
    "missing_gps_in_table",
    "record_count",
    "blank_field_rows_count",
    "duplicate_serial_in_gp_count",
    "exact_duplicate_rows_count",
    "portal_message",
    "csv_path",
    "duration_sec",
    "last_updated",
    "error_message",
]


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


def make_ps_key(district_id: str, ps_id: str) -> str:
    return f"{district_id}:{ps_id}"


def atomic_write_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = path.with_suffix(path.suffix + ".tmp")
    with tmp_path.open("w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    os.replace(tmp_path, path)


def atomic_write_csv(path: Path, columns: List[str], rows: List[Dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = path.with_suffix(path.suffix + ".tmp")
    with tmp_path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=columns, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow(row)
    os.replace(tmp_path, path)


class ProgressTracker:
    """Manages catalog discovery and resumable per-Panchayat-Samiti extraction state."""

    def __init__(
        self,
        catalog_path: Path,
        progress_json_path: Path,
        progress_csv_path: Path,
        skip_list: List[Dict[str, str]],
    ) -> None:
        self.catalog_path = catalog_path
        self.progress_json_path = progress_json_path
        self.progress_csv_path = progress_csv_path
        self.skip_ps_ids = {
            (str(item.get("district_id", "")).strip(), str(item.get("ps_id", "")).strip())
            for item in skip_list
        }
        self.skip_ps_names = {
            str(item.get("ps_name", "")).strip().upper()
            for item in skip_list
            if item.get("ps_name")
        }
        self.state: Dict[str, Any] = self._load_progress()

    def _load_progress(self) -> Dict[str, Any]:
        if self.progress_json_path.exists():
            with self.progress_json_path.open("r", encoding="utf-8") as f:
                return json.load(f)
        return {
            "meta": {
                "created_at": utc_now_iso(),
                "last_updated": utc_now_iso(),
                "total_districts": 0,
                "total_panchayat_samitis": 0,
            },
            "items": {},
        }

    def is_skipped_amber(self, district_id: str, ps_id: str, ps_name: str) -> bool:
        if (str(district_id).strip(), str(ps_id).strip()) in self.skip_ps_ids:
            return True
        if ps_name.strip().upper() in self.skip_ps_names:
            return True
        return False

    def has_catalog(self) -> bool:
        return self.catalog_path.exists() and len(self.state.get("items", {})) > 0

    def save_catalog_and_init_progress(
        self,
        districts: List[Dict[str, str]],
        panchayat_samitis: List[Dict[str, str]],
    ) -> None:
        catalog = {
            "discovered_at": utc_now_iso(),
            "district_count": len(districts),
            "panchayat_samiti_count": len(panchayat_samitis),
            "districts": districts,
            "panchayat_samitis": panchayat_samitis,
        }
        atomic_write_json(self.catalog_path, catalog)

        items: Dict[str, Dict[str, Any]] = self.state.setdefault("items", {})
        for ps in panchayat_samitis:
            key = make_ps_key(ps["district_id"], ps["ps_id"])
            if key in items:
                # Keep existing scraping progress, just refresh labels if needed
                items[key]["district_name"] = ps["district_name"]
                items[key]["ps_name"] = ps["ps_name"]
                if self.is_skipped_amber(ps["district_id"], ps["ps_id"], ps["ps_name"]):
                    items[key]["status"] = "skipped_amber"
                continue

            is_amber = self.is_skipped_amber(ps["district_id"], ps["ps_id"], ps["ps_name"])
            items[key] = {
                "ps_key": key,
                "district_id": str(ps["district_id"]),
                "district_name": ps["district_name"],
                "ps_id": str(ps["ps_id"]),
                "ps_name": ps["ps_name"],
                "status": "skipped_amber" if is_amber else "pending",
                "attempts": 0,
                "gp_dropdown_count": 25 if is_amber else None,
                "gp_in_table_count": 25 if is_amber else None,
                "missing_gps_count": 0 if is_amber else None,
                "missing_gps_in_table": [],
                "record_count": 209 if is_amber else 0,
                "blank_field_rows_count": 1 if is_amber else 0,
                "duplicate_serial_in_gp_count": 0,
                "exact_duplicate_rows_count": 0,
                "portal_message": "Skipped (provided in Sarpanch Result (AMBER) New.xlsx)" if is_amber else "",
                "csv_path": "",
                "duration_sec": 0.0,
                "last_updated": utc_now_iso(),
                "error_message": "",
            }

        self.state["meta"]["total_districts"] = len(districts)
        self.state["meta"]["total_panchayat_samitis"] = len(items)
        self.save()

    def update_item(self, ps_key: str, **updates: Any) -> None:
        item = self.state["items"][ps_key]
        item.update(updates)
        item["last_updated"] = utc_now_iso()
        self.save()

    def save(self) -> None:
        self.state["meta"]["last_updated"] = utc_now_iso()
        atomic_write_json(self.progress_json_path, self.state)

        sorted_items = self.get_all_items()
        csv_rows = []
        for item in sorted_items:
            row = dict(item)
            missing_list = row.get("missing_gps_in_table") or []
            if isinstance(missing_list, list):
                row["missing_gps_in_table"] = "; ".join(missing_list)
            csv_rows.append(row)
        atomic_write_csv(self.progress_csv_path, PROGRESS_CSV_COLUMNS, csv_rows)

    def get_all_items(self) -> List[Dict[str, Any]]:
        items = list(self.state.get("items", {}).values())
        return sorted(
            items,
            key=lambda x: (
                x.get("district_name", ""),
                x.get("ps_name", ""),
                int(x["ps_id"]) if str(x.get("ps_id", "")).isdigit() else 0,
            ),
        )

    def get_actionable_items(
        self,
        include_failed: bool = True,
        include_empty: bool = False,
    ) -> List[Dict[str, Any]]:
        allowed = {"pending"}
        if include_failed:
            allowed.add("failed")
        if include_empty:
            allowed.add("empty_on_portal")
        return [item for item in self.get_all_items() if item.get("status") in allowed]

    def summary_counts(self) -> Dict[str, int]:
        counts = {
            "total": 0,
            "completed": 0,
            "skipped_amber": 0,
            "empty_on_portal": 0,
            "failed": 0,
            "pending": 0,
            "total_records_scraped": 0,
            "total_records_including_amber": 0,
        }
        for item in self.state.get("items", {}).values():
            counts["total"] += 1
            st = item.get("status", "pending")
            if st in counts:
                counts[st] += 1
            rc = int(item.get("record_count") or 0)
            if st == "completed":
                counts["total_records_scraped"] += rc
                counts["total_records_including_amber"] += rc
            elif st == "skipped_amber":
                counts["total_records_including_amber"] += rc
        return counts
