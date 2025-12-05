#!/usr/bin/env python3
"""
Best Pick PDF profile extractor (half-page aware)

- Starts processing at the Table of Contents.
- Each body page can contain one or two half-page company profiles.
- Extracts header info (published name/phone), company quote (unchanged),
  ratings (incl. right-margin star counts), and the Verified Block with Trade License numbers.
- REMOVED: text extraction for these sections:
    Services Offered, Warranty, Company History, Distinctions,
    Employee Information, Additional Information, Services Not Offered, Minimum Job.

Output CSV — columns for those removed sections are no longer emitted.
'Category' equals the PAGE HEADER (top of each page).
"""

import re
import argparse
import tempfile
import pdfplumber
import pandas as pd
from pathlib import Path
from typing import List, Dict, Tuple, Optional

try:
    import tkinter as tk
    from tkinter import filedialog
except Exception:
    tk = None
    filedialog = None

# =========================
# Regexes & constants
# =========================
PHONE_RE = re.compile(r"\b(\d{3})[–-](\d{3})[–-](\d{4})\b")
INT_RE = re.compile(r"\b(\d{1,3}(?:,\d{3})*|\d+)\b")
PCT_RE = re.compile(r"(\d{1,3}(?:\.\d+)?)\s*%")

SECTION_LABELS = [
    "Services Offered","Warranty","Company History","Distinctions",
    "Employee Information","Additional Information","Services Not Offered","Minimum Job",
    "Homeowner Satisfaction Results","Rating:","Verified","Trade License",
    "Trade License(s) Not Required",
    "Best Pick Guaranteed","Scan QR","Best Pick Reports recommends:"
]

# Markers for license extraction in Verified block
START_TL_RE = re.compile(r"verified\s+trade\s+license\(s\)\s*:?", re.I)
END_GLI_RE  = re.compile(r"verified\s+general\s+liability", re.I)
ANY_VERIFIED_RE = re.compile(r"^\s*verified\b", re.I)

# =========================
# Utilities
# =========================
def extract_words_trimmed(page, top_margin=20, bottom_margin=20):
    """Words in reading flow, excluding headers/footers."""
    words = page.extract_words(
        use_text_flow=True,
        x_tolerance=2,
        y_tolerance=3,
        keep_blank_chars=False
    )
    if not words:
        return []
    page_h = page.height
    return [
        w
        for w in words
        if (w["top"] > top_margin and w["bottom"] < page_h - bottom_margin)
    ]


def group_lines_by_y(words, y_tol=3.0):
    """Group words into lines; return dicts with y, x0, x1, text, words."""
    if not words:
        return []
    words_sorted = sorted(words, key=lambda w: (w["top"], w["x0"]))
    lines = []
    current = []
    current_y = None

    for w in words_sorted:
        if current_y is None or abs(w["top"] - current_y) <= y_tol:
            current.append(w)
            current_y = w["top"] if current_y is None else (current_y + w["top"]) / 2.0
        else:
            current.sort(key=lambda ww: ww["x0"])
            text = " ".join(ww["text"] for ww in current)
            x0 = min(ww["x0"] for ww in current)
            x1 = max(ww["x1"] for ww in current)
            lines.append(
                {
                    "y": current[0]["top"],
                    "x0": x0,
                    "x1": x1,
                    "text": text,
                    "words": current[:],
                }
            )
            current = [w]
            current_y = w["top"]

    if current:
        current.sort(key=lambda ww: ww["x0"])
        text = " ".join(ww["text"] for ww in current)
        x0 = min(ww["x0"] for ww in current)
        x1 = max(ww["x1"] for ww in current)
        lines.append(
            {
                "y": current[0]["top"],
                "x0": x0,
                "x1": x1,
                "text": text,
                "words": current[:],
            }
        )
    return lines


def detect_two_profiles_split_y(words, page_height, band=0.08):
    """Split into two half-page regions if a low-density band exists near midline."""
    if not words:
        return None
    mid_y = page_height / 2.0
    near_mid = [
        w
        for w in words
        if (mid_y - page_height * band) <= w["top"] <= (mid_y + page_height * band)
    ]
    above = [w for w in words if w["top"] < mid_y - page_height * band]
    below = [w for w in words if w["top"] > mid_y + page_height * band]

    if (
        len(near_mid)
        <= max(2, 0.15 * max(len(above), len(below)))
        and len(above) > 10
        and len(below) > 10
    ):
        lines = sorted({w["top"] for w in words})
        gaps = []
        for i in range(len(lines) - 1):
            y0, y1 = lines[i], lines[i + 1]
            if (mid_y - page_height * 0.15) <= (y0 + y1) / 2.0 <= (mid_y + page_height * 0.15):
                gaps.append((y1 - y0, (y0 + y1) / 2.0))
        if gaps:
            gaps.sort(reverse=True)
            return gaps[0][1]
        return mid_y
    return None


def detect_gutter_x(words, bins=24):
    """Find the column gutter by histogram valley of word x centers."""
    if not words:
        return None
    xs = sorted(0.5 * (w["x0"] + w["x1"]) for w in words)
    xmin, xmax = xs[0], xs[-1]
    width = xmax - xmin if xmax > xmin else 1.0
    counts = [0] * bins
    edges = [xmin + i * width / bins for i in range(bins + 1)]

    def bi(x):
        if x <= xmin:
            return 0
        if x >= xmax:
            return bins - 1
        return min(bins - 1, int((x - xmin) / width * bins))

    for x in xs:
        counts[bi(x)] += 1

    left_band = int(bins * 0.30)
    right_band = int(bins * 0.70)
    if right_band <= left_band:
        left_band, right_band = 1, bins - 2

    valley_idx = min(range(left_band, right_band), key=lambda i: counts[i])
    return edges[valley_idx]


def split_words_into_columns(words, gutter_x):
    left = [w for w in words if 0.5 * (w["x0"] + w["x1"]) <= gutter_x]
    right = [w for w in words if 0.5 * (w["x0"] + w["x1"]) > gutter_x]
    return left, right


def find_published_name_left(lines_left):
    """First meaningful line in the left column."""
    for ln in lines_left[:8]:
        t = ln["text"].strip()
        if (
            t
            and "TABLE OF CONTENTS" not in t.upper()
            and "BEST PICK REPORTS" not in t.upper()
        ):
            return t, ln["y"]
    return "", None


def find_phone_top_right(lines_right):
    for ln in lines_right[:10]:
        m = PHONE_RE.search(ln["text"])
        if m:
            return f"{m.group(1)}-{m.group(2)}-{m.group(3)}", ln["y"]
    return "", None


def find_first_section_y(lines_all):
    for ln in lines_all:
        if any(lbl.lower() in ln["text"].lower() for lbl in SECTION_LABELS):
            return ln["y"]
    return None


def find_published_name_and_number(lines_all, page_width, y_after, y_before, min_span_ratio=0.65):
    """First line between y_after and y_before that spans most of the width."""
    if y_after is None or y_before is None:
        return ""
    if y_before < y_after:
        y_after, y_before = y_before, y_after
    for ln in lines_all:
        if y_after - 2 <= ln["y"] <= y_before + 2:
            span = (ln["x1"] - ln["x0"]) / page_width
            if span >= min_span_ratio and len(ln["text"].strip()) >= 8:
                return ln["text"].strip()
    return ""


def lines_to_plain(lines):
    return [ln["text"] for ln in lines]


def subsection_between(lines_text, start_kw, end_kws):
    """Extract text between a start label and the next of several end labels (inclusive of the start line)."""
    start_idx = next(
        (i for i, ln in enumerate(lines_text) if start_kw.lower() in ln.lower()),
        None,
    )
    if start_idx is None:
        return ""
    end_idx = next(
        (
            j
            for j in range(start_idx + 1, len(lines_text))
            if any(ek.lower() in lines_text[j].lower() for ek in end_kws)
        ),
        len(lines_text),
    )
    return "\n".join(s.strip() for s in lines_text[start_idx:end_idx])


def extract_category(lines: List[str]) -> str:
    """Legacy category heuristic (unused for final Category; kept for fallback)."""
    for i, ln in enumerate(lines):
        if "Best Pick Reports recommends:" in ln:
            for nxt in lines[i + 1 :]:
                if nxt.strip():
                    return nxt.strip()
    for ln in lines[:6]:
        if (
            re.search(r"[A-Z][a-z]+(?:\s+[&A-Za-z][A-Za-z]+)+", ln)
            and len(ln.strip()) <= 45
            and "BEST PICK REPORTS" not in ln.upper()
            and "TABLE OF CONTENTS" not in ln.upper()
        ):
            return ln.strip()
    return ""


def extract_rating(text: str) -> Optional[float]:
    m = re.search(r"Rating:\s*([0-9.]+)\s*out of\s*5", text, re.I)
    return float(m.group(1)) if m else None


# ---------- Verified + Trade License numbers ----------
_SECTION_GUARDS = [
    "Services Offered","Warranty","Company History","Distinctions",
    "Employee Information","Services Not Offered","Minimum Job",
    "Homeowner Satisfaction Results","Rating:","Verified","Trade License",
    "Trade License(s) Not Required",
    "Best Pick Guaranteed","Scan QR","Best Pick Reports recommends:"
]


def _build_verified_block(all_lines: List[str]) -> str:
    raw = [
        s
        for s in all_lines
        if (
            "verified" in s.lower()
            or "best pick guaranteed" in s.lower()
            or "trade license(s) not required" in s.lower()
        )
    ]
    seen = set()
    return "\n".join([s for s in raw if not (s in seen or seen.add(s))])


def _extract_between_markers(lines: List[str]) -> str:
    """
    - Find 'Verified Trade License(s)'
    - Capture all non-empty lines after it
    - Stop at 'Verified General Liability Insurance' (or next 'Verified …').
    """
    start_idx = next((i for i, s in enumerate(lines) if START_TL_RE.search(s)), None)
    if start_idx is None:
        return ""
    end_idx = next(
        (j for j in range(start_idx + 1, len(lines)) if END_GLI_RE.search(lines[j])),
        None,
    )
    if end_idx is None:
        end_idx = next(
            (j for j in range(start_idx + 1, len(lines)) if ANY_VERIFIED_RE.search(lines[j])),
            len(lines),
        )
    payload = []
    for s in lines[start_idx + 1 : end_idx]:
        t = s.strip()
        if not t:
            break
        payload.append(t)
    return " ".join(payload).strip()


def _extract_verified_and_licenses(left_lines: List[str], right_lines: List[str]) -> Tuple[str, str]:
    all_lines = left_lines + right_lines
    verified_block = _build_verified_block(all_lines)
    tl_numbers = _extract_between_markers(left_lines) or _extract_between_markers(right_lines)
    return verified_block, tl_numbers


# =========================
# Ratings: STRICT COUNTS from right margin (no percentages)
# =========================
def parse_star_counts_right_margin(lines_all: List[dict], page_width: float) -> Dict[str, Optional[int]]:
    """
    Pull five review counts (5★→1★) by geometry only.
    """
    if not lines_all:
        return {}

    lines_sorted = sorted(lines_all, key=lambda ln: ln["y"])

    # bounds
    start_idx = next(
        (
            i
            for i, ln in enumerate(lines_sorted)
            if "homeowner satisfaction results" in ln["text"].lower()
        ),
        None,
    )
    if start_idx is None:
        return {}

    END_KWS = [
        "Rating:",
        "Verified",
        "Trade License",
        "Trade License(s) Not Required",
        "Best Pick Guaranteed",
        "Scan QR",
        "Services Offered",
        "Best Pick Reports recommends:",
        "Services Not Offered",
        "Minimum Job",
    ]
    end_y = None
    for j in range(start_idx + 1, len(lines_sorted)):
        if any(kw.lower() in lines_sorted[j]["text"].lower() for kw in END_KWS):
            end_y = lines_sorted[j]["y"]
            break

    # Collect per-line rightmost non-% integer tokens
    perline_candidates = []  # (y, x0, val)
    for ln in lines_sorted[start_idx + 1 :]:
        if end_y is not None and ln["y"] >= end_y:
            break
        rightmost = None
        for w in ln["words"]:
            t = w["text"]
            if "%" in t:
                continue
            if INT_RE.fullmatch(t):
                if rightmost is None or w["x0"] > rightmost[0]:
                    rightmost = (w["x0"], t)
        if rightmost:
            x0, t = rightmost
            try:
                val = int(t.replace(",", ""))
            except Exception:  # pragma: no cover
                continue
            perline_candidates.append((ln["y"], x0, val))

    if not perline_candidates:
        return {}

    # Keep only those near the right margin via a dynamic threshold
    max_x0 = max(x0 for _, x0, _ in perline_candidates)
    RIGHT_MARGIN_PAD = 72.0  # pts (~1 inch)
    dynamic_thresh = max(0.65 * page_width, max_x0 - RIGHT_MARGIN_PAD)
    right_side = [
        (y, x0, v) for (y, x0, v) in perline_candidates if x0 >= dynamic_thresh
    ]
    if len(right_side) < 5:
        # relax a bit if needed
        dynamic_thresh = max(0.60 * page_width, max_x0 - 100.0)
        right_side = [
            (y, x0, v) for (y, x0, v) in perline_candidates if x0 >= dynamic_thresh
        ]

    if not right_side:
        return {}

    right_side.sort(key=lambda t: t[0])  # by y

    # Choose the tightest run of 5 vertically
    if len(right_side) >= 5:
        best_window = None
        best_span = None
        for i in range(len(right_side) - 4):
            ys = [right_side[i + k][0] for k in range(5)]
            span = ys[-1] - ys[0]
            if best_span is None or span < best_span:
                best_span = span
                best_window = right_side[i : i + 5]
        chosen = best_window
    else:
        chosen = right_side  # fewer than 5 available

    # Map top→bottom to star buckets
    result: Dict[str, Optional[int]] = {}
    chosen.sort(key=lambda t: t[0])
    vals = [v for (_, _, v) in chosen]
    # pad to 5 with None
    while len(vals) < 5:
        vals.append(None)
    result["five_star_count"] = vals[0]
    result["four_star_count"] = vals[1]
    result["three_star_count"] = vals[2]
    result["two_star_count"] = vals[3]
    result["one_star_count"] = vals[4]
    return result


# =========================
# Page header (Category) detection
# =========================
def get_page_header_text(page) -> str:
    """Topmost line (after margins) is used as the PAGE HEADER -> Category."""
    words = extract_words_trimmed(page)
    if not words:
        return ""
    lines = group_lines_by_y(words)
    if not lines:
        return ""
    top_line = min(lines, key=lambda ln: ln["y"])
    return top_line["text"].strip()


# =========================
# Profile extraction
# =========================
def extract_profile_from_region(region_words, page, page_header_text: str):
    if not region_words:
        return None

    gutter_x = detect_gutter_x(region_words) or (page.width / 2.0)
    left_w, right_w = split_words_into_columns(region_words, gutter_x)

    lines_all = group_lines_by_y(region_words)
    lines_left = group_lines_by_y(left_w)
    lines_right = group_lines_by_y(right_w)

    # 1) Published Name (top-left) & Phone (top-right)
    published_name, y_pub = find_published_name_left(lines_left)
    phone, y_phone = find_phone_top_right(lines_right)

    # 2) First section anchor (used for quote window only)
    first_section_y = find_first_section_y(lines_all)

    # 3) Company Quote (kept)
    y_after = max([y for y in [y_pub, y_phone] if y is not None], default=None)
    y_before = (
        first_section_y if first_section_y is not None else (y_after + 60 if y_after else None)
    )
    published_name_and_number = find_published_name_and_number(
        lines_all, page.width, y_after, y_before
    )

    # 4) Per-column plain text
    left_plain = lines_to_plain(lines_left)
    right_plain = lines_to_plain(lines_right)
    joined_for_sections = left_plain + right_plain

    # 5) Ratings block (kept for totals/% only; star counts come from geometry)
    ratings_block = subsection_between(
        joined_for_sections,
        "Homeowner Satisfaction Results",
        [
            "Rating:",
            "Verified",
            "Trade License",
            "Trade License(s) Not Required",
            "Best Pick Guaranteed",
            "Scan QR",
            "Services Offered",
            "Best Pick Reports recommends:",
            "Services Not Offered",
            "Minimum Job",
        ],
    )

    reviews_total = None
    recommendation_rate_pct = None
    if ratings_block:
        for ln in [l.strip() for l in ratings_block.splitlines() if l.strip()]:
            if reviews_total is None and re.search(
                r"(reviews\s*(surveyed|total|count)|total\s*reviews)", ln, re.I
            ):
                m = INT_RE.search(ln)
                if m:
                    try:
                        reviews_total = int(m.group(1).replace(",", ""))
                    except Exception:
                        pass
            if recommendation_rate_pct is None and re.search(
                r"(recommend|would\s*recommend|recommendation\s*rate)", ln, re.I
            ):
                m = PCT_RE.search(ln)
                if m:
                    try:
                        recommendation_rate_pct = float(m.group(1))
                    except Exception:
                        pass

    # STRICT counts from right margin (no mixing with %)
    star_counts = parse_star_counts_right_margin(lines_all, page.width)

    # Verified block + license numbers
    verified_block, trade_license_numbers = _extract_verified_and_licenses(
        left_plain, right_plain
    )

    return {
        "Page": None,
        "Phone": phone,
        "Category": page_header_text,
        "Published Name + Number": published_name_and_number,
        "Rating (out of 5)": extract_rating("\n".join(joined_for_sections))
        if joined_for_sections
        else None,
        "Reviews Surveyed": reviews_total,
        "Recommendation Rate (%)": recommendation_rate_pct,
        "Five-Star (count)": star_counts.get("five_star_count"),
        "Four-Star (count)": star_counts.get("four_star_count"),
        "Three-Star (count)": star_counts.get("three_star_count"),
        "Two-Star (count)": star_counts.get("two_star_count"),
        "One-Star (count)": star_counts.get("one_star_count"),
        # Removed section text fields
        "Verified Block": verified_block,
        "Trade License Numbers": trade_license_numbers,
    }


# =========================
# TOC detection
# =========================
def find_first_toc_page_index(pdf) -> int:
    for idx, page in enumerate(pdf.pages):
        words = extract_words_trimmed(page)
        if not words:
            continue
        snippet = " ".join(w["text"] for w in words[:200])
        if "table of contents" in snippet.lower():
            return idx
    return 0


# =========================
# Main processing
# =========================
def process_pdf(pdf_path: Path) -> pd.DataFrame:
    rows = []
    with pdfplumber.open(pdf_path) as pdf:
        start_idx = find_first_toc_page_index(pdf)
        for page_number, page in enumerate(pdf.pages[start_idx:], start=start_idx + 1):
            words = extract_words_trimmed(page)
            if not words:
                continue

            page_header_text = get_page_header_text(page)

            split_y = detect_two_profiles_split_y(words, page.height)
            regions = (
                [(0, page.height)]
                if split_y is None
                else [(0, split_y), (split_y, page.height)]
            )

            for (y0, y1) in regions:
                region_words = [w for w in words if y0 <= w["top"] <= y1]
                if not region_words:
                    continue

                preview = " ".join(w["text"] for w in region_words[:300]).lower()
                if "services offered" not in preview:
                    continue

                rec = extract_profile_from_region(region_words, page, page_header_text)
                if not rec:
                    continue
                rec["Page"] = page_number
                rows.append(rec)

    df = pd.DataFrame(rows)

    # ---- Cast numeric columns ----
    int_cols = [
        "Five-Star (count)",
        "Four-Star (count)",
        "Three-Star (count)",
        "Two-Star (count)",
        "One-Star (count)",
        "Reviews Surveyed",
        "Page",
    ]
    for col in int_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").astype("Int64")

    if "Rating (out of 5)" in df.columns:
        df["Rating (out of 5)"] = pd.to_numeric(
            df["Rating (out of 5)"], errors="coerce"
        )

    if "Recommendation Rate (%)" in df.columns:
        df["Recommendation Rate (%)"] = pd.to_numeric(
            df["Recommendation Rate (%)"], errors="coerce"
        )

    cols = [
        "Page",
        "Phone",
        "Category",
        "Published Name + Number",
        "Rating (out of 5)",
        "Reviews Surveyed",
        "Recommendation Rate (%)",
        "Five-Star (count)",   # H
        "Four-Star (count)",   # I
        "Three-Star (count)",  # J
        "Two-Star (count)",    # K
        "One-Star (count)",    # L
        "Verified Block",      # T
        "Trade License Numbers",  # U
    ]
    for c in cols:
        if c not in df.columns:
            df[c] = None
    if df.empty:
        return pd.DataFrame(columns=cols)
    return df[cols]


def pick_pdf_file_via_dialog() -> Path:
    """
    Open a native file dialog to select a single PDF.
    Exits if no selection is made or the environment cannot show a dialog.
    """
    if tk is None or filedialog is None:
        raise SystemExit(
            "This script requires a GUI file dialog, but tkinter is unavailable."
        )

    root = tk.Tk()
    root.withdraw()
    root.update_idletasks()
    try:
        file_path = filedialog.askopenfilename(
            title="Select Best Pick PDF",
            filetypes=[("PDF files", "*.pdf")],
            multiple=False,
        )
    finally:
        root.destroy()

    if not file_path:
        raise SystemExit("No file selected. Exiting.")

    p = Path(file_path)
    if not p.exists() or p.suffix.lower() != ".pdf":
        raise SystemExit("Invalid selection (must be an existing .pdf file). Exiting.")
    return p


# =========================
# Streamlit-friendly wrapper
# =========================
def run_pipeline(pdf_bytes: bytes, expected_order_df: Optional[pd.DataFrame] = None) -> pd.DataFrame:
    """
    Streamlit-friendly entry point.

    - Accepts PDF bytes (from file_uploader)
    - Optionally accepts expected_order_df (unused, kept for API parity)
    - Returns the Profiles DataFrame from process_pdf
    """
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        tmp.write(pdf_bytes)
        tmp_path = Path(tmp.name)

    try:
        df_profiles = process_pdf(tmp_path)
    finally:
        try:
            tmp_path.unlink()
        except OSError:
            pass

    return df_profiles


# =========================
# CLI entry point
# =========================
def main():
    parser = argparse.ArgumentParser(
        description=(
            "Extract one row per half-page Company Profile from a Best Pick PDF "
            "(starting at TOC). Category equals page header."
        )
    )
    # No positional PDF arg — we always use the dialog now
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="Output CSV path (default: <selected-pdf-stem>_profiles.csv)",
    )
    args = parser.parse_args()

    pdf_path = pick_pdf_file_via_dialog()
    df = process_pdf(pdf_path)

    # Choose output name (default comes from the selected PDF name)
    output_csv = args.output or f"{pdf_path.stem}_profiles.csv"

    # Write CSV
    df.to_csv(
        output_csv,
        index=False,
        encoding="utf-8-sig",
        na_rep="",
    )

    print(f"Profile extraction complete! {len(df)} profiles saved to {output_csv}")


if __name__ == "__main__":
    main()
