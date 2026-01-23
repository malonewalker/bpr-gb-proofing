#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import re
import sys
import unicodedata
from typing import Optional, Set, List, Tuple

import pandas as pd

# -----------------------------
# Utilities
# -----------------------------
NON_DIGIT = re.compile(r"\D+")
NON_ALNUM = re.compile(r"[^A-Za-z0-9]+")

def normalize_phone(p: str) -> str:
    return NON_DIGIT.sub("", str(p or ""))

def normalize_token(s: str) -> str:
    return NON_ALNUM.sub("", str(s or "")).upper()

def to_clean_str(val) -> str:
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except Exception:
        pass
    return str(val)

def norm_text(s: str) -> str:
    """Normalize unicode, unify quotes, collapse spaces, lowercase."""
    s = to_clean_str(s)
    s = unicodedata.normalize("NFKD", s)
    s = s.replace("’", "'").replace("‘", "'").replace("“", '"').replace("”", '"')
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s

def contains_phrase(text: str, expected_substring: str) -> bool:
    return norm_text(expected_substring) in norm_text(text)

def split_licenses(raw: str) -> Set[str]:
    """Tokenize licenses; treat 'Not Required' as empty."""
    txt = to_clean_str(raw)
    if is_blankish(txt) or "not required" in norm_text(txt):
        return set()
    parts = re.split(r"[,\|;/\s]+", txt)
    tokens = {normalize_token(p) for p in parts if p}
    return {t for t in tokens if t}

_BLANKISH = {"", "na", "n/a", "none", "null", "not applicable", "tbd", "-", "—", "–"}

def is_blankish(val) -> bool:
    s = to_clean_str(val).strip()
    if not s:
        return True
    low = s.lower()
    if low in _BLANKISH:
        return True
    if re.fullmatch(r"[-–—\s]+", s):
        return True
    return False

def ensure_columns(df: pd.DataFrame, cols: list) -> None:
    for c in cols:
        if c not in df.columns:
            df[c] = ""

def add_note(acc: list, msg: str):
    if msg and msg not in acc:
        acc.append(msg)

# Small helper to encode/decode structured error details as a string in a single cell
# Format per item: issue|expected|found
# Items are joined by '||'
def encode_error_items(items: List[Tuple[str, str, str]]) -> str:
    safe = []
    for issue, exp, fnd in items:
        # guard against accidental separators in text
        issue = issue.replace("||", "| |").replace("|", "¦")
        exp = exp.replace("||", "| |").replace("|", "¦")
        fnd = fnd.replace("||", "| |").replace("|", "¦")
        safe.append(f"{issue}|{exp}|{fnd}")
    return "||".join(safe)

def decode_error_items(s: str) -> List[Tuple[str, str, str]]:
    out: List[Tuple[str, str, str]] = []
    s = to_clean_str(s)
    if not s.strip():
        return out
    for item in s.split("||"):
        parts = item.split("|")
        if len(parts) == 3:
            # restore '¦' back to '|'
            out.append(tuple(p.replace("¦", "|") for p in parts))  # type: ignore
        elif len(parts) == 1:
            out.append((parts[0].replace("¦", "|"), "", ""))
    return out

# -----------------------------
# Core Checks
# -----------------------------

def run_checks(primary: pd.DataFrame, bbb: pd.DataFrame) -> pd.DataFrame:
    # Primary columns
    P_PHONE = "Phone"
    P_RATING = "Rating (out of 5)"
    P_STARS = "Five-Star (count)"
    P_LICENSES = "Trade License Numbers"
    P_VERIFIED = "Verified Block"
    P_PUBLISHED_NAME = "Published Name"

    # BBB columns
    B_BOOK_PHONE = "Book Phone Number"
    B_LICENSES = "Licenses"
    B_WC_STATUS = "WC Status"
    B_PUBLISHED_NAME = "PublishedName"

    ensure_columns(primary, [P_PHONE, P_RATING, P_STARS, P_LICENSES, P_VERIFIED, P_PUBLISHED_NAME])
    ensure_columns(bbb, [B_BOOK_PHONE, B_LICENSES, B_WC_STATUS, B_PUBLISHED_NAME])

    # Normalize for join
    primary["_phone_norm"] = primary[P_PHONE].fillna("").astype(str).map(normalize_phone)
    bbb["_book_phone_norm"] = bbb[B_BOOK_PHONE].fillna("").astype(str).map(normalize_phone)

    # Build BBB lookup by normalized phone
    bbb_lookup = {}
    for _, row in bbb.iterrows():
        phone_norm = row.get("_book_phone_norm", "")
        if not phone_norm:
            continue
        bbb_lic_raw = to_clean_str(row.get(B_LICENSES, ""))
        wc_status = to_clean_str(row.get(B_WC_STATUS, ""))
        bbb_published_name = to_clean_str(row.get(B_PUBLISHED_NAME, ""))
        bbb_lookup[phone_norm] = {
            "licenses_raw": bbb_lic_raw,
            "licenses_set": split_licenses(bbb_lic_raw),
            "wc_status": wc_status,
            "book_phone_raw": to_clean_str(row.get(B_BOOK_PHONE, "")),
            "published_name": bbb_published_name,
        }

    notes_internal, notes_compare = [], []
    errors_detail_col: List[str] = []

    def five_star_missing(stars_val: str) -> bool:
        # Treat "1000" as the sentinel for missing in the PDF export
        s = to_clean_str(stars_val).strip()
        return (s == "") or (s == "1000")

    for _, row in primary.iterrows():
        row_notes_internal, row_notes_compare = [], []
        row_error_items: List[Tuple[str, str, str]] = []

        phone = to_clean_str(row.get(P_PHONE, ""))
        rating = to_clean_str(row.get(P_RATING, ""))
        stars = to_clean_str(row.get(P_STARS, ""))
        licenses_raw = to_clean_str(row.get(P_LICENSES, ""))
        verified_block = to_clean_str(row.get(P_VERIFIED, ""))
        published_name = to_clean_str(row.get(P_PUBLISHED_NAME, ""))

        # --- Internal checks ---
        if normalize_phone(phone) == "":
            issue = "missing phone"
            add_note(row_notes_internal, issue)
            # Expected vs Found for phone discrepancy
            row_error_items.append((issue, "non-empty phone", phone))

        if (rating.strip() == "") or five_star_missing(stars):
            add_note(row_notes_internal, "missing qual data")

        if "zzz" in norm_text(licenses_raw):
            add_note(row_notes_internal, "license number missing")

        # --- Cross-file checks vs BBB ---
        phone_norm = normalize_phone(phone)
        if phone_norm and phone_norm in bbb_lookup:
            bbb_info = bbb_lookup[phone_norm]
            bbb_lic_raw = bbb_info["licenses_raw"]
            bbb_lic_set = bbb_info["licenses_set"]
            wc_status = bbb_info["wc_status"]
            bbb_published_name = bbb_info["published_name"]

            # Check Published Name - STRICT exact match
            if published_name != bbb_published_name:
                issue = "published name mismatch"
                add_note(row_notes_compare, issue)
                row_error_items.append((issue, bbb_published_name, published_name))

            primary_lic_set = split_licenses(licenses_raw)

            # License number mismatch (both sides have tokens but no overlap)
            if primary_lic_set and bbb_lic_set and primary_lic_set.isdisjoint(bbb_lic_set):
                issue = "license number mismatch"
                add_note(row_notes_compare, issue)
                # Expected (BBB) vs Found (Primary)
                row_error_items.append((issue, bbb_lic_raw, licenses_raw))

            # Workers' Comp phrase expected if WC Status != exempt
            if norm_text(wc_status) != "exempt":
                if "verified workers' comp" not in norm_text(verified_block):
                    add_note(
                        row_notes_compare,
                        "expected 'Verified Workers' Comp' in Verified Block"
                    )
                    row_error_items.append((
                        "expected 'Verified Workers' Comp' in Verified Block",
                        "Verified Workers' Comp",
                        verified_block
                    ))

            # Trade License(s) expectation from BBB truth
            bbb_has_license = (not is_blankish(bbb_lic_raw)) and ("not required" not in norm_text(bbb_lic_raw))

            if bbb_has_license:
                # Expect Verified Trade License(s)
                if not (
                    contains_phrase(verified_block, "Verified Trade License(s)")
                    or contains_phrase(verified_block, "Verified Trade License")
                ):
                    add_note(row_notes_compare, "missing license?; review checkboxes")
                    row_error_items.append((
                        "missing license?; review checkboxes",
                        "Verified Trade License(s)",
                        verified_block
                    ))
            else:
                # Expect Not Required
                if not (
                    contains_phrase(verified_block, "Trade License(s) Not Required")
                    or contains_phrase(verified_block, "Trade License Not Required")
                    or contains_phrase(verified_block, "Trade Licenses Not Required")
                ):
                    add_note(row_notes_compare, "expected 'Trade License(s) Not Required' in Verified Block")
                    row_error_items.append((
                        "expected 'Trade License(s) Not Required' in Verified Block",
                        "Trade License(s) Not Required",
                        verified_block
                    ))

                # Optional: flag if primary has license tokens but BBB says none
                if primary_lic_set:
                    add_note(row_notes_compare, "license in PDF but missing in BBB")
                    row_error_items.append((
                        "license in PDF but missing in BBB",
                        "(none in BBB / Not Required)",
                        licenses_raw
                    ))
        else:
            issue = "no BBB match by phone"
            add_note(row_notes_compare, issue)
            row_error_items.append((issue, "BBB record matching phone", phone))

        notes_internal.append("; ".join(row_notes_internal))
        notes_compare.append("; ".join(row_notes_compare))
        errors_detail_col.append(encode_error_items(row_error_items))

    primary["Notes_Internal"] = notes_internal
    primary["Notes_Compare"] = notes_compare

    # --- NEW: Combined ERRORS column ---
    def _combine_errors(a, b):
        a = a.strip() if isinstance(a, str) else ""
        b = b.strip() if isinstance(b, str) else ""
        if a and b:
            return f"{a}; {b}"
        return a or b

    primary["ERRORS"] = [
        _combine_errors(i, c) for i, c in zip(primary["Notes_Internal"], primary["Notes_Compare"])
    ]

    # Helper column used by the ERRORS tab to populate Expected/Found rows
    primary["ERRORS_DETAIL"] = errors_detail_col

    primary.drop(columns=["_phone_norm"], inplace=True, errors="ignore")
    return primary

def build_errors_tab(primary: pd.DataFrame) -> pd.DataFrame:
    """
    Create the ERRORS sheet with columns:
      Sheet, Row, Key, Issue, Expected, Found, Page

    - Sheet: literal "Profiles"
    - Row: Excel row number on the 'Profiles' sheet (index + 2)
    - Key: "Category/Published Name + Number"
    - Issue: one issue per row (exploded)
    - Expected / Found: populated for license & phone discrepancies (and some text expectations)
    - Page: from 'Page'
    """
    ensure_columns(primary, ["Category", "Published Name + Number", "Page", "ERRORS", "ERRORS_DETAIL"])

    df = primary.reset_index(drop=False).copy()
    df.rename(columns={"index": "__row_index"}, inplace=True)

    rows = []
    for _, r in df.iterrows():
        key = f"{to_clean_str(r.get('Category',''))}/{to_clean_str(r.get('Published Name + Number',''))}".strip("/")
        page = to_clean_str(r.get("Page", ""))
        excel_row = int(r["__row_index"]) + 2

        detail_items = decode_error_items(to_clean_str(r.get("ERRORS_DETAIL", "")))
        if detail_items:
            for issue, exp, fnd in detail_items:
                issue_clean = issue.strip()
                if not issue_clean:
                    continue
                rows.append({
                    "Sheet": "Profiles",
                    "Row": excel_row,
                    "Key": key,
                    "Issue": issue_clean,
                    "Expected": exp,
                    "Found": fnd,
                    "Page": page,
                })
        else:
            # Fallback: if no structured items, use the concatenated ERRORS text in one row
            issue_text = to_clean_str(r.get("ERRORS", "")).strip()
            if issue_text:
                rows.append({
                    "Sheet": "Profiles",
                    "Row": excel_row,
                    "Key": key,
                    "Issue": issue_text,
                    "Expected": "",
                    "Found": "",
                    "Page": page,
                })

    errors_df = pd.DataFrame(rows).reset_index(drop=True)
    return errors_df

# -----------------------------
# Streamlit-friendly wrapper
# -----------------------------
def run_pipeline(primary_df: pd.DataFrame, bbb_df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Streamlit-friendly entry point.

    - primary_df: Profiles dataframe (from proofing pipeline)
    - bbb_df: BBB reference dataframe
    Returns:
      (checked_profiles_df, errors_tab_df)
    """
    # Work on copies so original dfs passed in aren't mutated unintentionally
    primary_copy = primary_df.copy()
    bbb_copy = bbb_df.copy()

    checked = run_checks(primary_copy, bbb_copy)
    errors_tab = build_errors_tab(checked)

    return checked, errors_tab

# -----------------------------
# Main
# -----------------------------
def infer_output_path(primary_path: str) -> str:
    base, _ = os.path.splitext(primary_path)
    return f"{base}_checked.xlsx"

def load_table(path: str) -> pd.DataFrame:
    """
    Load the BBB / reference table from a filesystem path.

    Supports:
      - .xlsx / .xls (first sheet) via openpyxl
      - .csv / .txt via pandas.read_csv

    If the file has an Excel-looking extension but isn't a real Excel file
    (e.g., CSV saved as .xlsx), we fall back to CSV parsing.
    """
    if not path or not isinstance(path, str):
        raise ValueError(f"load_table expected a filesystem path (str), got: {repr(path)}")

    if not os.path.exists(path):
        raise FileNotFoundError(f"BBB / reference file not found at path: {path}")

    ext = os.path.splitext(path)[1].lower()

    # ---------- Excel-looking extensions ----------
    if ext in [".xlsx", ".xls"]:
        try:
            # First, try as real Excel
            return pd.read_excel(path, dtype=str, engine="openpyxl")
        except Exception as e:
            # This covers "File is not a zip file", "Excel file format cannot be determined", etc.
            print(f"[WARN] Failed to read {path!r} as Excel ({ext}): {e}. Trying CSV parser instead...")
            # Try CSV fallback (common case: CSV content with .xlsx extension)
            try:
                try:
                    return pd.read_csv(
                        path,
                        dtype=str,
                        keep_default_na=False,
                        na_values=[""],
                        encoding="utf-8",
                    )
                except UnicodeDecodeError:
                    return pd.read_csv(
                        path,
                        dtype=str,
                        keep_default_na=False,
                        na_values=[""],
                        encoding="latin-1",
                    )
            except Exception as e2:
                # If *both* Excel and CSV parsing fail, give a clear combined error
                raise RuntimeError(
                    f"Failed to read file at {path!r} ({ext}) "
                    f"as Excel (error: {e}) and as CSV (error: {e2})"
                )

    # ---------- Plain CSV / TXT ----------
    elif ext in [".csv", ".txt"]:
        try:
            try:
                return pd.read_csv(
                    path,
                    dtype=str,
                    keep_default_na=False,
                    na_values=[""],
                    encoding="utf-8",
                )
            except UnicodeDecodeError:
                return pd.read_csv(
                    path,
                    dtype=str,
                    keep_default_na=False,
                    na_values=[""],
                    encoding="latin-1",
                )
        except Exception as e:
            raise RuntimeError(f"Failed to read delimited file at {path!r} ({ext}): {e}")

    # ---------- Anything else ----------
    else:
        raise ValueError(f"Unsupported file extension for BBB table: {ext} (path={path!r})")
