"""Canonical schema, headers, and lossless value helpers for Rajasthan SEC Sarpanch data."""

from __future__ import annotations

import re
from typing import Any, Optional

# 17 canonical internal column keys matching the portal table (columns 0..16)
PORTAL_FIELD_KEYS = [
    "sr_no",
    "gram_panchayat_name",
    "gram_panchayat_category",
    "candidate_serial_no",
    "candidate_name",
    "father_husband_name",
    "gender",
    "marital_status",
    "candidate_category",
    "education_status",
    "occupation",
    "age",
    "capital_assets",
    "children_before_1995",
    "children_on_or_after_1995",
    "mobile_no",
    "email_address",
]

# Full 19 internal column keys including District & Panchayat Samiti
COMBINED_FIELD_KEYS = PORTAL_FIELD_KEYS + [
    "district",
    "panchayat_samiti",
]

# Cleaned headers matching Sarpanch Result (AMBER) New.xlsx (Row 1 and Row 2)
CLEANED_ROW1_HEADERS = [
    "Sr. No.",
    "Name of Gram Panchayat",
    "Category of Gram Panchayat",
    "Contesting Candidate\nSerial No.",
    "Name of Contesting Candidate",
    "Father/Husband of\nContesting Candidate",
    "Gender",
    "Marital Status",
    "Category of Candidate",
    "Education Status",
    "Contesting Candidate Occupation",
    "Age",
    "Total Value of Capital Assets\n(Land, Building, Jewellery, etc.)",
    "Children",
    None,  # Merged with N1:O1
    "Mobile No.",
    "E-mail Address",
    "District",
    "Panchayat Samiti",
]

CLEANED_ROW2_HEADERS = [
    None,
    None,
    None,
    None,
    None,
    None,
    None,
    None,
    None,
    None,
    None,
    None,
    None,
    "Before 27.11.1995",
    "On or After 28.11.1995",
    None,
    None,
    None,
    None,
]

# Raw portal headers (kept for reference/config option)
RAW_ROW1_HEADERS = [
    "Sr. No.",
    "Name Of GramPanchayat",
    "Category Of GramPanchayat",
    "Contesting Candidate\nSerial No.",
    "Name Of Contesting Candidate",
    "Father/Husband Of\nContesting Candidate",
    "Gender",
    "Martial Status",
    "Category of Candidate",
    "Education Status",
    "Contesting Candidate Occupation",
    "Age",
    "Total Value of Capital Assets\n(Land,Building,Jewellary & etc.)",
    "Children",
    None,
    "Mobile No.",
    "E-mail Address",
    "District",
    "Panchayat Samiti",
]

RAW_ROW2_HEADERS = [
    None,
    None,
    None,
    None,
    None,
    None,
    None,
    None,
    None,
    None,
    None,
    None,
    None,
    "Before 27.11.1995",
    "On Or After 28.11.1995",
    None,
    None,
    None,
    None,
]

# Column widths inspired by Sarpanch Result (AMBER) New.xlsx + generous widths for District & PS
EXCEL_COLUMN_WIDTHS = {
    "A": 8.0,
    "B": 22.0,
    "C": 22.0,
    "D": 18.0,
    "E": 26.0,
    "F": 26.0,
    "G": 8.5,
    "H": 13.0,
    "I": 18.0,
    "J": 22.0,
    "K": 34.0,
    "L": 7.0,
    "M": 26.0,
    "N": 17.0,
    "O": 19.0,
    "P": 15.0,
    "Q": 32.0,
    "R": 22.0,
    "S": 32.0,
}

# Merged ranges for the 2-row header (19 columns: A..S)
HEADER_MERGE_RANGES_19 = [
    "A1:A2",
    "B1:B2",
    "C1:C2",
    "D1:D2",
    "E1:E2",
    "F1:F2",
    "G1:G2",
    "H1:H2",
    "I1:I2",
    "J1:J2",
    "K1:K2",
    "L1:L2",
    "M1:M2",
    "N1:O1",
    "P1:P2",
    "Q1:Q2",
    "R1:R2",
    "S1:S2",
]

HEADER_MERGE_RANGES_17 = HEADER_MERGE_RANGES_19[:-2]


def clean_text(value: Any) -> str:
    """Normalize non-breaking spaces and surrounding whitespace without altering content."""
    if value is None:
        return ""
    s = str(value).replace("\u00a0", " ")
    s = re.sub(r"[ \t\r\n]+", " ", s).strip()
    return s


def parse_int_or_raw(value: Any) -> Any:
    """Convert clean integer values to int; preserve non-integer raw source strings verbatim."""
    if value is None:
        return None
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        if value != value:  # NaN
            return None
        return int(value) if value.is_integer() else value
    s = clean_text(value)
    if s == "":
        return None
    if re.fullmatch(r"-?\d+", s):
        return int(s)
    if re.fullmatch(r"-?\d+\.0+", s):
        return int(float(s))
    return s


def parse_assets_or_raw(value: Any) -> Any:
    """Convert numeric asset amounts to int/float; preserve non-numeric raw values verbatim."""
    if value is None:
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        if value != value:  # NaN
            return None
        return int(value) if value.is_integer() else value
    s = clean_text(value)
    if s == "":
        return None
    no_commas = s.replace(",", "")
    if re.fullmatch(r"-?\d+", no_commas):
        return int(no_commas)
    if re.fullmatch(r"-?\d+\.\d+", no_commas):
        f = float(no_commas)
        return int(f) if f.is_integer() else f
    return s


def clean_mobile_str(value: Any) -> Optional[str]:
    """Preserve mobile numbers cleanly as text (avoiding Excel scientific notation)."""
    if value is None:
        return None
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if value != value:  # NaN
            return None
        if value.is_integer():
            return str(int(value))
        return str(value)
    s = clean_text(value)
    if s == "":
        return None
    # Handle float-like CSV serialization such as '9001350988.0'
    if re.fullmatch(r"\d+\.0+", s):
        return s.split(".")[0]
    return s


def sanitize_filename(name: str) -> str:
    """Convert a district/PS name to a safe filesystem slug."""
    slug = re.sub(r"[^A-Za-z0-9]+", "_", clean_text(name)).strip("_")
    return slug or "UNKNOWN"
