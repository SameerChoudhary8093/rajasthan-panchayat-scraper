"""Combined Excel workbook builder and Error / Missing-Record / Summary report generator."""

from __future__ import annotations

import csv
import logging
import re
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
import pandas as pd
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from progress_tracker import ProgressTracker, atomic_write_csv
from schema import (
    CLEANED_ROW1_HEADERS,
    CLEANED_ROW2_HEADERS,
    COMBINED_FIELD_KEYS,
    EXCEL_COLUMN_WIDTHS,
    HEADER_MERGE_RANGES_17,
    HEADER_MERGE_RANGES_19,
    PORTAL_FIELD_KEYS,
    RAW_ROW1_HEADERS,
    RAW_ROW2_HEADERS,
    clean_mobile_str,
    clean_text,
    parse_assets_or_raw,
    parse_int_or_raw,
)

EMAIL_REGEX = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
MOBILE_10_REGEX = re.compile(r"^[6-9]\d{9}$")



# ----------------------------------------------------------------------
# Structured "Contesting Candidates" sheet layout
# ----------------------------------------------------------------------
# The portal's "Sr. No." restarts at 1 for every Panchayat Samiti, so the
# combined sheet leads with a continuous overall serial number plus District
# and Panchayat Samiti, followed by the original 17 AMBER columns.
# (key, row-1 header, row-2 header, width, kind)
CANDIDATE_SHEET_LAYOUT = [
    ("overall_sno", "S. No.\n(Overall)", None, 9.0, "center"),
    ("district", "District", None, 18.0, "left"),
    ("panchayat_samiti", "Panchayat Samiti", None, 30.0, "left"),
    ("sr_no", "Sr. No.\n(within Panchayat Samiti)", None, 14.0, "center"),
    ("gram_panchayat_name", "Name of Gram Panchayat", None, 22.0, "left"),
    ("gram_panchayat_category", "Category of Gram Panchayat", None, 22.0, "left"),
    ("candidate_serial_no", "Contesting Candidate\nSerial No.", None, 14.0, "center"),
    ("candidate_name", "Name of Contesting Candidate", None, 26.0, "left"),
    ("father_husband_name", "Father/Husband of\nContesting Candidate", None, 26.0, "left"),
    ("gender", "Gender", None, 8.5, "center"),
    ("marital_status", "Marital Status", None, 12.0, "center"),
    ("candidate_category", "Category of Candidate", None, 18.0, "left"),
    ("education_status", "Education Status", None, 18.0, "left"),
    ("occupation", "Contesting Candidate Occupation", None, 32.0, "left"),
    ("age", "Age", None, 7.0, "center"),
    ("capital_assets", "Total Value of Capital Assets\n(Land, Building, Jewellery, etc.)", None, 26.0, "money"),
    ("children_before_1995", "Children", "Before 27.11.1995", 16.0, "center"),
    ("children_on_or_after_1995", None, "On or After 28.11.1995", 18.0, "center"),
    ("mobile_no", "Mobile No.", None, 14.0, "text"),
    ("email_address", "E-mail Address", None, 32.0, "left"),
]


def write_candidates_sheet(wb: "openpyxl.Workbook", title: str, records: List[Dict[str, Any]]) -> None:
    """Append a formatted, filterable candidate sheet to a write-only workbook."""
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    sub_header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    header_side = Side(style="thin", color="FFFFFF")
    header_border = Border(left=header_side, right=header_side, top=header_side, bottom=header_side)
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10)
    aligns = {
        "center": Alignment(horizontal="center", vertical="center"),
        "left": Alignment(horizontal="left", vertical="center"),
        "money": Alignment(horizontal="right", vertical="center"),
        "text": Alignment(horizontal="center", vertical="center"),
    }
    center_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws = wb.create_sheet(title=title)
    ws.freeze_panes = "E3"  # keep S. No., District, PS, Sr. No. visible while scrolling
    n_cols = len(CANDIDATE_SHEET_LAYOUT)
    last_col = get_column_letter(n_cols)

    for idx, (_, _, _, width, _) in enumerate(CANDIDATE_SHEET_LAYOUT, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width

    r1, r2 = [], []
    for key, h1, h2, _, _ in CANDIDATE_SHEET_LAYOUT:
        is_child = key.startswith("children_")
        c1 = WriteOnlyCell(ws, value=h1)
        c2 = WriteOnlyCell(ws, value=h2)
        for c in (c1, c2):
            c.font = header_font
            c.fill = header_fill
            c.alignment = center_wrap
            c.border = header_border
        if is_child:
            c2.fill = sub_header_fill
        r1.append(c1)
        r2.append(c2)
    ws.append(r1)
    ws.append(r2)

    child_cols = [i for i, row in enumerate(CANDIDATE_SHEET_LAYOUT, start=1) if row[0].startswith("children_")]
    for idx in range(1, n_cols + 1):
        col = get_column_letter(idx)
        if idx in child_cols:
            continue
        ws.merged_cells.ranges.add(f"{col}1:{col}2")
    ws.merged_cells.ranges.add(f"{get_column_letter(child_cols[0])}1:{get_column_letter(child_cols[-1])}1")

    for overall, rec in enumerate(records, start=1):
        row_cells = []
        for key, _, _, _, kind in CANDIDATE_SHEET_LAYOUT:
            val = overall if key == "overall_sno" else rec.get(key)
            c = WriteOnlyCell(ws, value=val)
            c.font = data_font
            c.alignment = aligns[kind]
            if kind == "money" and isinstance(val, (int, float)):
                c.number_format = "#,##0"
            elif kind == "text":
                c.number_format = "@"
            row_cells.append(c)
        ws.append(row_cells)

    # Excel filter dropdowns on the header so users can pick District / PS / GP
    ws.auto_filter.ref = f"A2:{last_col}{len(records) + 2}"


def load_amber_excel_untouched(
    amber_path: Path,
    district_name: str = "JAIPUR",
    ps_name: str = "AMBER PANCHAYAT SAMITI",
) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    """Read Sarpanch Result (AMBER) New.xlsx verbatim without altering source values."""
    wb = openpyxl.load_workbook(amber_path, read_only=True, data_only=True)
    ws = wb.active

    records: List[Dict[str, Any]] = []
    anomalies: List[Dict[str, Any]] = []

    for row_idx, row in enumerate(ws.iter_rows(min_row=3, max_col=17, values_only=True), start=3):
        if all(cell is None or str(cell).strip() == "" for cell in row):
            continue

        rec: Dict[str, Any] = {
            "sr_no": parse_int_or_raw(row[0]),
            "gram_panchayat_name": clean_text(row[1]) or None,
            "gram_panchayat_category": clean_text(row[2]) or None,
            "candidate_serial_no": parse_int_or_raw(row[3]),
            "candidate_name": clean_text(row[4]) or None,
            "father_husband_name": clean_text(row[5]) or None,
            "gender": clean_text(row[6]) or None,
            "marital_status": clean_text(row[7]) or None,
            "candidate_category": clean_text(row[8]) or None,
            "education_status": clean_text(row[9]) or None,
            "occupation": clean_text(row[10]) or None,
            "age": parse_int_or_raw(row[11]),
            "capital_assets": parse_assets_or_raw(row[12]),
            "children_before_1995": parse_int_or_raw(row[13]),
            "children_on_or_after_1995": parse_int_or_raw(row[14]),
            "mobile_no": clean_mobile_str(row[15]),
            "email_address": clean_text(row[16]) or None,
            "district": district_name,
            "panchayat_samiti": ps_name,
        }
        records.append(rec)

        missing_cols = [k for k in PORTAL_FIELD_KEYS if rec[k] is None or rec[k] == ""]
        if missing_cols:
            anomalies.append(
                {
                    "category": "AMBER_BLANK_FIELDS",
                    "severity": "INFO",
                    "district": district_name,
                    "panchayat_samiti": ps_name,
                    "gram_panchayat": rec["gram_panchayat_name"] or "",
                    "sr_no": rec["sr_no"] or "",
                    "candidate_name": rec["candidate_name"] or "",
                    "details": f"Original AMBER Excel row {row_idx} has blank fields: {', '.join(missing_cols)} (preserved untouched per approval)",
                }
            )
        if rec["email_address"] and not EMAIL_REGEX.match(str(rec["email_address"])):
            anomalies.append(
                {
                    "category": "MALFORMED_EMAIL_SOURCE",
                    "severity": "INFO",
                    "district": district_name,
                    "panchayat_samiti": ps_name,
                    "gram_panchayat": rec["gram_panchayat_name"] or "",
                    "sr_no": rec["sr_no"] or "",
                    "candidate_name": rec["candidate_name"] or "",
                    "details": f"Source email value preserved verbatim: {rec['email_address']!r}",
                }
            )

    wb.close()
    return records, anomalies


def load_scraped_ps_csv(csv_path: Path) -> List[Dict[str, Any]]:
    """Load one Panchayat Samiti CSV using Pandas (all strings) and convert numeric columns safely."""
    df = pd.read_csv(csv_path, dtype=str, keep_default_na=False)
    records: List[Dict[str, Any]] = []
    for raw in df.to_dict(orient="records"):
        rec: Dict[str, Any] = {
            "sr_no": parse_int_or_raw(raw.get("sr_no")),
            "gram_panchayat_name": clean_text(raw.get("gram_panchayat_name")) or None,
            "gram_panchayat_category": clean_text(raw.get("gram_panchayat_category")) or None,
            "candidate_serial_no": parse_int_or_raw(raw.get("candidate_serial_no")),
            "candidate_name": clean_text(raw.get("candidate_name")) or None,
            "father_husband_name": clean_text(raw.get("father_husband_name")) or None,
            "gender": clean_text(raw.get("gender")) or None,
            "marital_status": clean_text(raw.get("marital_status")) or None,
            "candidate_category": clean_text(raw.get("candidate_category")) or None,
            "education_status": clean_text(raw.get("education_status")) or None,
            "occupation": clean_text(raw.get("occupation")) or None,
            "age": parse_int_or_raw(raw.get("age")),
            "capital_assets": parse_assets_or_raw(raw.get("capital_assets")),
            "children_before_1995": parse_int_or_raw(raw.get("children_before_1995")),
            "children_on_or_after_1995": parse_int_or_raw(raw.get("children_on_or_after_1995")),
            "mobile_no": clean_mobile_str(raw.get("mobile_no")),
            "email_address": clean_text(raw.get("email_address")) or None,
            "district": clean_text(raw.get("district")) or None,
            "panchayat_samiti": clean_text(raw.get("panchayat_samiti")) or None,
        }
        records.append(rec)
    return records


def build_all_outputs(config: Dict[str, Any], base_dir: Path, logger: logging.Logger) -> Dict[str, Any]:
    """Combine AMBER + all completed PS CSVs into the final Excel workbook and audit reports."""
    paths = config["paths"]
    tracker = ProgressTracker(
        catalog_path=base_dir / paths["catalog_file"],
        progress_json_path=base_dir / paths["progress_json"],
        progress_csv_path=base_dir / paths["progress_csv"],
        skip_list=config.get("skip_panchayat_samitis", []),
    )

    amber_path = base_dir / config["amber_excel_path"]
    if not amber_path.exists():
        raise FileNotFoundError(f"AMBER reference Excel not found at: {amber_path}")

    amber_skip_cfg = (config.get("skip_panchayat_samitis") or [{}])[0]
    amber_district = amber_skip_cfg.get("district_name", "JAIPUR")
    amber_ps = amber_skip_cfg.get("ps_name", "AMBER PANCHAYAT SAMITI")

    logger.info("Loading original AMBER workbook untouched: %s", amber_path)
    amber_records, audit_issues = load_amber_excel_untouched(amber_path, amber_district, amber_ps)
    logger.info("Loaded %d records from AMBER Excel.", len(amber_records))

    # Build combined dataset in clean District -> Panchayat Samiti -> original Sr. No. order
    all_items = tracker.get_all_items()
    combined_records: List[Dict[str, Any]] = []

    # Data quality summary counters
    quality_stats = {
        "amber_records": len(amber_records),
        "scraped_records": 0,
        "total_records": 0,
        "total_unique_gps": 0,
        "rows_with_blank_email": 0,
        "rows_with_malformed_email": 0,
        "rows_with_nonstandard_mobile": 0,
        "rows_with_other_blank_fields": 0,
        "duplicate_serial_in_gp_groups": 0,
        "exact_duplicate_rows_in_ps": 0,
        "missing_gps_from_table_total": 0,
    }

    district_rollup: Dict[str, Dict[str, Any]] = defaultdict(
        lambda: {
            "total_ps": 0,
            "completed_ps": 0,
            "amber_ps": 0,
            "empty_ps": 0,
            "failed_ps": 0,
            "pending_ps": 0,
            "total_gps": 0,
            "total_records": 0,
        }
    )

    seen_ps_keys = set()
    amber_inserted = False

    for item in all_items:
        ps_key = item["ps_key"]
        if ps_key in seen_ps_keys:
            continue
        seen_ps_keys.add(ps_key)

        d_name = item["district_name"]
        ps_name = item["ps_name"]
        status = item.get("status", "pending")
        d_stat = district_rollup[d_name]
        d_stat["total_ps"] += 1

        if status == "skipped_amber":
            d_stat["amber_ps"] += 1
            amber_gps = {r["gram_panchayat_name"] for r in amber_records if r.get("gram_panchayat_name")}
            d_stat["total_gps"] += len(amber_gps)
            d_stat["total_records"] += len(amber_records)
            combined_records.extend(amber_records)
            amber_inserted = True
            continue

        if status == "completed":
            csv_rel = item.get("csv_path") or ""
            csv_full = base_dir / csv_rel
            if not csv_rel or not csv_full.exists():
                audit_issues.append(
                    {
                        "category": "MISSING_PS_CSV_FILE",
                        "severity": "ERROR",
                        "district": d_name,
                        "panchayat_samiti": ps_name,
                        "gram_panchayat": "",
                        "sr_no": "",
                        "candidate_name": "",
                        "details": f"PS marked completed in progress.json but CSV missing at {csv_rel}",
                    }
                )
                d_stat["failed_ps"] += 1
                continue

            ps_records = load_scraped_ps_csv(csv_full)
            d_stat["completed_ps"] += 1
            ps_gps = {r["gram_panchayat_name"] for r in ps_records if r.get("gram_panchayat_name")}
            d_stat["total_gps"] += len(ps_gps)
            d_stat["total_records"] += len(ps_records)
            quality_stats["scraped_records"] += len(ps_records)
            combined_records.extend(ps_records)

            # Audit missing GPs (present in PS dropdown but 0 contesting candidates in table, whitespace-normalized)
            ps_gps_lower = {clean_text(g).lower() for g in ps_gps if g}
            raw_missing_gps = item.get("missing_gps_in_table") or []
            missing_gps = [
                clean_text(g)
                for g in raw_missing_gps
                if clean_text(g) and clean_text(g).lower() not in ps_gps_lower
            ]
            if missing_gps:
                quality_stats["missing_gps_from_table_total"] += len(missing_gps)
                audit_issues.append(
                    {
                        "category": "GP_MISSING_IN_TABLE",
                        "severity": "WARNING",
                        "district": d_name,
                        "panchayat_samiti": ps_name,
                        "gram_panchayat": "; ".join(missing_gps),
                        "sr_no": "",
                        "candidate_name": "",
                        "details": f"{len(missing_gps)} GP(s) listed in dropdown had 0 candidates in Contesting Sarpanch table (possibly uncontested/deferred).",
                    }
                )

            # Audit duplicate Candidate Serial No. within same GP & exact duplicate rows inside PS
            gp_serial_counts = Counter((r["gram_panchayat_name"], r["candidate_serial_no"]) for r in ps_records)
            dup_serials = [(gp, sn, cnt) for (gp, sn), cnt in gp_serial_counts.items() if cnt > 1]
            if dup_serials:
                quality_stats["duplicate_serial_in_gp_groups"] += len(dup_serials)
                sample_str = ", ".join(f"{gp} (#{sn} x{cnt})" for gp, sn, cnt in dup_serials[:5])
                audit_issues.append(
                    {
                        "category": "DUPLICATE_CANDIDATE_SERIAL_IN_GP",
                        "severity": "INFO",
                        "district": d_name,
                        "panchayat_samiti": ps_name,
                        "gram_panchayat": dup_serials[0][0] or "",
                        "sr_no": "",
                        "candidate_name": "",
                        "details": f"{len(dup_serials)} GP/Serial group(s) share duplicate Candidate Serial No. on portal (preserved verbatim): {sample_str}",
                    }
                )

            row_sigs = Counter(tuple(r[k] for k in PORTAL_FIELD_KEYS[1:]) for r in ps_records)
            exact_dups = [(sig, cnt) for sig, cnt in row_sigs.items() if cnt > 1]
            if exact_dups:
                extra_cnt = sum(cnt - 1 for _, cnt in exact_dups)
                quality_stats["exact_duplicate_rows_in_ps"] += extra_cnt
                audit_issues.append(
                    {
                        "category": "EXACT_DUPLICATE_PORTAL_ROW",
                        "severity": "WARNING",
                        "district": d_name,
                        "panchayat_samiti": ps_name,
                        "gram_panchayat": str(exact_dups[0][0][0] or ""),
                        "sr_no": "",
                        "candidate_name": str(exact_dups[0][0][3] or ""),
                        "details": f"{extra_cnt} exact duplicate candidate row(s) returned by portal within same PS (preserved verbatim).",
                    }
                )

        elif status == "empty_on_portal":
            d_stat["empty_ps"] += 1
            audit_issues.append(
                {
                    "category": "EMPTY_ON_PORTAL",
                    "severity": "WARNING",
                    "district": d_name,
                    "panchayat_samiti": ps_name,
                    "gram_panchayat": "",
                    "sr_no": "",
                    "candidate_name": "",
                    "details": f"0 candidates published on portal (dropdown GPs={item.get('gp_dropdown_count')}). Portal message: {item.get('portal_message', '')}",
                }
            )
        elif status == "failed":
            d_stat["failed_ps"] += 1
            audit_issues.append(
                {
                    "category": "SCRAPE_FAILED",
                    "severity": "ERROR",
                    "district": d_name,
                    "panchayat_samiti": ps_name,
                    "gram_panchayat": "",
                    "sr_no": "",
                    "candidate_name": "",
                    "details": f"Failed after {item.get('attempts')} attempts: {item.get('error_message', '')}",
                }
            )
        else:
            d_stat["pending_ps"] += 1
            audit_issues.append(
                {
                    "category": "PENDING_NOT_SCRAPED",
                    "severity": "WARNING",
                    "district": d_name,
                    "panchayat_samiti": ps_name,
                    "gram_panchayat": "",
                    "sr_no": "",
                    "candidate_name": "",
                    "details": "Panchayat Samiti has not been scraped yet.",
                }
            )

    if not amber_inserted:
        combined_records = amber_records + combined_records

    # Record-level data quality metrics across the combined dataset (non-destructive)
    unique_gp_keys = set()
    for rec in combined_records:
        unique_gp_keys.add((rec.get("district"), rec.get("panchayat_samiti"), rec.get("gram_panchayat_name")))
        email = rec.get("email_address")
        if not email:
            quality_stats["rows_with_blank_email"] += 1
        elif not EMAIL_REGEX.match(str(email)):
            quality_stats["rows_with_malformed_email"] += 1

        mob = rec.get("mobile_no")
        if not mob or not MOBILE_10_REGEX.match(str(mob)) or str(mob) in ("9999999999", "0000000000"):
            quality_stats["rows_with_nonstandard_mobile"] += 1

        # Blank in any core field other than email
        other_blank = any(
            rec.get(k) is None or rec.get(k) == ""
            for k in PORTAL_FIELD_KEYS
            if k != "email_address"
        )
        if other_blank:
            quality_stats["rows_with_other_blank_fields"] += 1

    quality_stats["total_records"] = len(combined_records)
    quality_stats["total_unique_gps"] = len(unique_gp_keys)

    # Write combined Excel file using OpenPyXL write_only mode for speed & low memory
    excel_cfg = config.get("excel", {})
    use_cleaned = bool(excel_cfg.get("use_cleaned_amber_headers", True))
    include_dist_ps = bool(excel_cfg.get("include_district_and_ps_columns", True))
    combined_excel_path = base_dir / paths["combined_excel"]

    logger.info(
        "Writing combined Excel workbook (%d rows) to %s ...",
        len(combined_records),
        combined_excel_path,
    )
    _write_combined_excel(
        out_path=combined_excel_path,
        records=combined_records,
        all_items=all_items,
        district_rollup=district_rollup,
        tracker_counts=tracker.summary_counts(),
        quality_stats=quality_stats,
        use_cleaned_headers=use_cleaned,
        include_dist_ps=include_dist_ps,
    )

    # Write Error & Missing-Record CSV + Markdown Report
    error_csv_path = base_dir / paths["error_report_csv"]
    error_md_path = base_dir / paths["error_report_md"]
    summary_md_path = base_dir / paths["summary_report_md"]

    error_cols = [
        "category",
        "severity",
        "district",
        "panchayat_samiti",
        "gram_panchayat",
        "sr_no",
        "candidate_name",
        "details",
    ]
    atomic_write_csv(error_csv_path, error_cols, audit_issues)
    _write_error_report_md(error_md_path, audit_issues, quality_stats, tracker.summary_counts())
    _write_summary_report_md(
        summary_md_path,
        tracker.summary_counts(),
        quality_stats,
        district_rollup,
        combined_excel_path,
    )

    logger.info("All Excel and summary deliverables written successfully.")
    return {
        "counts": tracker.summary_counts(),
        "quality_stats": quality_stats,
        "audit_issues_count": len(audit_issues),
        "combined_excel": str(combined_excel_path),
    }


def _write_combined_excel(
    out_path: Path,
    records: List[Dict[str, Any]],
    all_items: List[Dict[str, Any]],
    district_rollup: Dict[str, Dict[str, Any]],
    tracker_counts: Dict[str, int],
    quality_stats: Dict[str, int],
    use_cleaned_headers: bool = True,
    include_dist_ps: bool = True,
) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = out_path.with_suffix(".tmp.xlsx")

    wb = openpyxl.Workbook(write_only=True)

    # Shared styles
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    sub_header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    summary_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    thin_side = Side(style="thin", color="D9D9D9")
    header_side = Side(style="thin", color="FFFFFF")
    cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    header_border = Border(left=header_side, right=header_side, top=header_side, bottom=header_side)

    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    bold_font = Font(name="Calibri", size=10, bold=True)
    data_font = Font(name="Calibri", size=10)
    center_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center")
    center_align = Alignment(horizontal="center", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")

    # ------------------------------------------------------------------
    # Sheet 1: Contesting Candidates (main dataset, continuous overall S. No.)
    # ------------------------------------------------------------------
    write_candidates_sheet(wb, "Contesting Candidates", records)

    # ------------------------------------------------------------------
    # Sheet 2: Summary (KPIs + District-wise Breakdown)
    # ------------------------------------------------------------------
    ws_sum = wb.create_sheet(title="Summary")
    ws_sum.freeze_panes = "A2"
    sum_widths = {"A": 34, "B": 16, "C": 16, "D": 16, "E": 16, "F": 14, "G": 14, "H": 18, "I": 20}
    for col, w in sum_widths.items():
        ws_sum.column_dimensions[col].width = w

    title_cell = WriteOnlyCell(ws_sum, value="Rajasthan Panchayat Samiti-Wise Contesting Sarpanch Extraction Summary")
    title_cell.font = Font(name="Calibri", size=13, bold=True, color="1F4E78")
    ws_sum.append([title_cell])
    ws_sum.append(["Generated At", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    ws_sum.append(["Source Portal", "https://sec.rajasthan.gov.in/GramPanchayatDetails.aspx (General Election: OCT-NOV 2026)"])
    ws_sum.append([])

    kpi_header = [WriteOnlyCell(ws_sum, value="Metric"), WriteOnlyCell(ws_sum, value="Count")]
    for c in kpi_header:
        c.font = header_font
        c.fill = header_fill
        c.alignment = center_align
    ws_sum.append(kpi_header)

    kpis = [
        ("Total Districts Discovered", len(district_rollup)),
        ("Total Panchayat Samitis Discovered", tracker_counts["total"]),
        ("Panchayat Samitis Completed (Scraped)", tracker_counts["completed"]),
        ("Panchayat Samitis Skipped (AMBER Provided)", tracker_counts["skipped_amber"]),
        ("Panchayat Samitis Completed + AMBER Combined", tracker_counts["completed"] + tracker_counts["skipped_amber"]),
        ("Panchayat Samitis Empty on Portal (0 Records)", tracker_counts["empty_on_portal"]),
        ("Panchayat Samitis Failed", tracker_counts["failed"]),
        ("Panchayat Samitis Pending", tracker_counts["pending"]),
        ("Total Unique Gram Panchayats (District+PS+GP)", quality_stats["total_unique_gps"]),
        ("Total Candidate Records (AMBER)", quality_stats["amber_records"]),
        ("Total Candidate Records (Scraped)", quality_stats["scraped_records"]),
        ("Total Candidate Records (Combined)", quality_stats["total_records"]),
    ]
    for label, val in kpis:
        c1 = WriteOnlyCell(ws_sum, value=label)
        c1.font = bold_font if "Combined" in label else data_font
        c1.border = cell_border
        c2 = WriteOnlyCell(ws_sum, value=val)
        c2.font = bold_font if "Combined" in label else data_font
        c2.number_format = "#,##0"
        c2.alignment = right_align
        c2.border = cell_border
        ws_sum.append([c1, c2])

    ws_sum.append([])
    dist_headers = [
        "District",
        "Total PSs",
        "Scraped OK",
        "AMBER (Skipped)",
        "Empty on Portal",
        "Failed",
        "Pending",
        "Gram Panchayats",
        "Total Candidates",
    ]
    dh_cells = []
    for h in dist_headers:
        c = WriteOnlyCell(ws_sum, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = center_wrap
        dh_cells.append(c)
    ws_sum.append(dh_cells)

    for d_name in sorted(district_rollup.keys()):
        st = district_rollup[d_name]
        vals = [
            d_name,
            st["total_ps"],
            st["completed_ps"],
            st["amber_ps"],
            st["empty_ps"],
            st["failed_ps"],
            st["pending_ps"],
            st["total_gps"],
            st["total_records"],
        ]
        row_c = []
        for idx, v in enumerate(vals):
            c = WriteOnlyCell(ws_sum, value=v)
            c.font = data_font
            c.border = cell_border
            if idx == 0:
                c.alignment = left_align
            else:
                c.alignment = right_align
                c.number_format = "#,##0"
            row_c.append(c)
        ws_sum.append(row_c)

    # Grand total row on Summary
    grand_vals = [
        "GRAND TOTAL",
        sum(v["total_ps"] for v in district_rollup.values()),
        sum(v["completed_ps"] for v in district_rollup.values()),
        sum(v["amber_ps"] for v in district_rollup.values()),
        sum(v["empty_ps"] for v in district_rollup.values()),
        sum(v["failed_ps"] for v in district_rollup.values()),
        sum(v["pending_ps"] for v in district_rollup.values()),
        sum(v["total_gps"] for v in district_rollup.values()),
        sum(v["total_records"] for v in district_rollup.values()),
    ]
    gt_cells = []
    for idx, v in enumerate(grand_vals):
        c = WriteOnlyCell(ws_sum, value=v)
        c.font = bold_font
        c.fill = summary_fill
        c.border = cell_border
        if idx > 0:
            c.alignment = right_align
            c.number_format = "#,##0"
        gt_cells.append(c)
    ws_sum.append(gt_cells)

    # ------------------------------------------------------------------
    # Sheet 3: PS Status (full audit table of all discovered PSs)
    # ------------------------------------------------------------------
    ws_ps = wb.create_sheet(title="PS Status")
    ws_ps.freeze_panes = "A2"
    ps_headers = [
        "District ID",
        "District",
        "PS ID",
        "Panchayat Samiti",
        "Status",
        "Attempts",
        "GPs in Dropdown",
        "GPs in Table",
        "Missing GPs Count",
        "Candidate Records",
        "Duration (s)",
        "Last Updated",
        "Portal Message / Error",
    ]
    ps_widths = [11, 22, 9, 32, 16, 10, 16, 14, 17, 17, 13, 24, 42]
    for idx, w in enumerate(ps_widths, start=1):
        ws_ps.column_dimensions[get_column_letter(idx)].width = w

    ps_h_cells = []
    for h in ps_headers:
        c = WriteOnlyCell(ws_ps, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = center_wrap
        ps_h_cells.append(c)
    ws_ps.append(ps_h_cells)

    for item in all_items:
        msg = item.get("error_message") or item.get("portal_message") or ""
        vals = [
            parse_int_or_raw(item.get("district_id")),
            item.get("district_name"),
            parse_int_or_raw(item.get("ps_id")),
            item.get("ps_name"),
            item.get("status"),
            item.get("attempts", 0),
            item.get("gp_dropdown_count"),
            item.get("gp_in_table_count"),
            item.get("missing_gps_count"),
            item.get("record_count", 0),
            item.get("duration_sec", 0.0),
            item.get("last_updated", ""),
            msg,
        ]
        r_cells = []
        for idx, v in enumerate(vals):
            c = WriteOnlyCell(ws_ps, value=v)
            c.font = data_font
            c.border = cell_border
            if idx in (0, 2, 4, 5, 6, 7, 8, 9, 10):
                c.alignment = center_align
            else:
                c.alignment = left_align
            r_cells.append(c)
        ws_ps.append(r_cells)

    wb.save(tmp_path)
    tmp_path.replace(out_path)


def _write_error_report_md(
    out_path: Path,
    audit_issues: List[Dict[str, Any]],
    quality_stats: Dict[str, int],
    counts: Dict[str, int],
) -> None:
    by_cat: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    for issue in audit_issues:
        by_cat[issue["category"]].append(issue)

    lines = [
        "# Error & Missing-Record Report",
        "",
        f"**Generated:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}  ",
        f"**Total Issues / Notices Logged:** {len(audit_issues)}",
        "",
        "## 1. Issue Breakdown by Category",
        "",
        "| Category | Severity | Count |",
        "|---|---|---:|",
    ]
    for cat in sorted(by_cat.keys()):
        sev = by_cat[cat][0]["severity"]
        lines.append(f"| `{cat}` | {sev} | {len(by_cat[cat])} |")

    lines.extend(
        [
            "",
            "## 2. Source Data Quality Audit (Preserved Verbatim)",
            "",
            "| Quality Check | Count | Notes |",
            "|---|---:|---|",
            f"| Rows with blank E-mail Address | {quality_stats['rows_with_blank_email']:,} | Common across SEC portal records; left blank verbatim |",
            f"| Rows with non-standard / malformed E-mail | {quality_stats['rows_with_malformed_email']:,} | Preserved verbatim (e.g., AMBER Sr. No. 38, literal `0`) |",
            f"| Rows with non-standard Mobile No. | {quality_stats['rows_with_nonstandard_mobile']:,} | Blank, `0`, `9999999999`, or non-10-digit values preserved verbatim |",
            f"| Rows with blank non-email fields | {quality_stats['rows_with_other_blank_fields']:,} | Includes AMBER Row 209 (preserved untouched) |",
            f"| GP / Candidate Serial No. duplicate groups on portal | {quality_stats['duplicate_serial_in_gp_groups']:,} | Authentic SEC portal duplicate serials; never dropped |",
            f"| Exact duplicate candidate rows within same PS | {quality_stats['exact_duplicate_rows_in_ps']:,} | Returned verbatim by SEC portal; flagged for review |",
            f"| Gram Panchayats in dropdown with 0 table rows | {quality_stats['missing_gps_from_table_total']:,} | Typically uncontested / deferred GPs |",
            "",
        ]
    )

    # Detailed sections for Failures and Empty-on-Portal PSs
    for section_title, cat_key in [
        ("3. Failed Panchayat Samitis", "SCRAPE_FAILED"),
        ("4. Empty-on-Portal Panchayat Samitis (0 Records Published)", "EMPTY_ON_PORTAL"),
        ("5. Pending Panchayat Samitis (Not Yet Scraped)", "PENDING_NOT_SCRAPED"),
    ]:
        entries = by_cat.get(cat_key, [])
        lines.append(f"## {section_title} ({len(entries)})")
        lines.append("")
        if not entries:
            lines.append("_None._")
            lines.append("")
            continue
        lines.append("| District | Panchayat Samiti | Details |")
        lines.append("|---|---|---|")
        for e in entries:
            details_esc = str(e["details"]).replace("|", "\\|")
            lines.append(f"| {e['district']} | {e['panchayat_samiti']} | {details_esc} |")
        lines.append("")

    # Missing GPs summary (top 50 in MD, complete list in CSV)
    gp_missing = by_cat.get("GP_MISSING_IN_TABLE", [])
    lines.append(f"## 6. Panchayat Samitis with Missing Gram Panchayats in Table ({len(gp_missing)})")
    lines.append("")
    if not gp_missing:
        lines.append("_None._")
    else:
        lines.append("_Full row-level list is available in `error_and_missing_report.csv`._")
        lines.append("")
        lines.append("| District | Panchayat Samiti | Missing GP(s) |")
        lines.append("|---|---|---|")
        for e in gp_missing[:50]:
            gps_esc = str(e["gram_panchayat"]).replace("|", "\\|")
            lines.append(f"| {e['district']} | {e['panchayat_samiti']} | {gps_esc} |")
        if len(gp_missing) > 50:
            lines.append(f"\n_...and {len(gp_missing) - 50} more Panchayat Samitis (see `error_and_missing_report.csv`)._")
    lines.append("")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text("\n".join(lines), encoding="utf-8")


def _write_summary_report_md(
    out_path: Path,
    counts: Dict[str, int],
    quality_stats: Dict[str, int],
    district_rollup: Dict[str, Dict[str, Any]],
    combined_excel_path: Path,
) -> None:
    completed_plus_amber = counts["completed"] + counts["skipped_amber"]
    pending_or_failed = counts["pending"] + counts["failed"]

    lines = [
        "# Rajasthan Panchayat Samiti Scraping Summary Report",
        "",
        f"**Generated:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}  ",
        f"**Source:** `https://sec.rajasthan.gov.in/GramPanchayatDetails.aspx` (`General Election` — `OCT-NOV 2026`, `Contesting Sarpanch`)  ",
        f"**Combined Excel File:** `{combined_excel_path.name}`",
        "",
        "## Overall Status",
        "",
        "| Metric | Count |",
        "|---|---:|",
        f"| Total Districts Discovered | **{len(district_rollup)}** |",
        f"| Total Panchayat Samitis Discovered on Portal | **{counts['total']}** |",
        f"| Panchayat Samitis Extracted (Scraped Successfully) | **{counts['completed']}** |",
        f"| Panchayat Samitis Skipped (AMBER — Original File Preserved) | **{counts['skipped_amber']}** |",
        f"| **Total Panchayat Samitis With Data Combined** | **{completed_plus_amber}** |",
        f"| Panchayat Samitis Empty on Portal (0 Records Published) | **{counts['empty_on_portal']}** |",
        f"| Panchayat Samitis Failed (Requires Retry) | **{counts['failed']}** |",
        f"| Panchayat Samitis Pending | **{counts['pending']}** |",
        f"| **Total Actionable Pending + Failed** | **{pending_or_failed}** |",
        f"| Total Unique Gram Panchayats Combined | **{quality_stats['total_unique_gps']:,}** |",
        f"| Candidate Records from AMBER Excel | **{quality_stats['amber_records']:,}** |",
        f"| Candidate Records Scraped from SEC Portal | **{quality_stats['scraped_records']:,}** |",
        f"| **Total Combined Candidate Records** | **{quality_stats['total_records']:,}** |",
        "",
        "## District-Wise Breakdown",
        "",
        "| District | Total PSs | Scraped OK | AMBER | Empty on Portal | Failed | Pending | GPs | Candidates |",
        "|---|---:|---:|---:|---:|---:|---:|---:|---:|",
    ]
    for d_name in sorted(district_rollup.keys()):
        st = district_rollup[d_name]
        lines.append(
            f"| {d_name} | {st['total_ps']} | {st['completed_ps']} | {st['amber_ps']} | "
            f"{st['empty_ps']} | {st['failed_ps']} | {st['pending_ps']} | {st['total_gps']:,} | {st['total_records']:,} |"
        )
    lines.append("")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text("\n".join(lines), encoding="utf-8")
