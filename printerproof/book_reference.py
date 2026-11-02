"""Published Google Sheet reference for BPR cover, spine, and back-cover checks."""

from io import BytesIO
from pathlib import Path
import re
from urllib.parse import urlsplit, urlunsplit
from urllib.request import Request, urlopen

import pandas as pd


REFERENCE_URL = (
    "https://docs.google.com/spreadsheets/d/e/"
    "2PACX-1vRpD6cgG8ilq5RjIwqWXcKvyIufwXq4VwKulwDYvzfPdbuAa2GC2CBnFcUVw230_0cqf2WTMaFFqDoS/"
    "pub?output=xlsx"
)
BOOK_NAME_COLUMN = "BOOK NAME  (Front + Back Cover Name)"
SUBHEADER_COLUMN = "BOOK SUB-HEADER"
QR_COLUMN = "Back Cover QR Bitly"
REQUIRED_COLUMNS = {BOOK_NAME_COLUMN, SUBHEADER_COLUMN, QR_COLUMN}


def _clean(value):
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def _key(value):
    return _clean(value).casefold()


def _read_table(path):
    path = Path(path)
    if path.suffix.lower() in (".csv", ".txt"):
        return pd.read_csv(path, dtype=str, encoding="utf-8-sig")
    return pd.read_excel(path, dtype=str)


def book_lookup_from_dataframe(frame):
    """Extract the unique text after the first hyphen in the BBB Book column."""
    if "Book" not in frame.columns:
        raise ValueError("BBB/reference input is missing the 'Book' column.")
    values = []
    for raw in frame["Book"].dropna():
        value = _clean(raw)
        if not value:
            continue
        lookup = value.split("-", 1)[1].strip() if "-" in value else value
        if _key(lookup) not in {_key(existing) for existing in values}:
            values.append(lookup)
    if not values:
        raise ValueError("BBB/reference input contains no usable Book value.")
    if len(values) != 1:
        raise ValueError(f"BBB/reference input contains multiple Book values: {values}")
    return values[0]


def book_lookup_from_file(path):
    return book_lookup_from_dataframe(_read_table(path))


def load_published_workbook(content):
    excel = pd.ExcelFile(BytesIO(content))
    sheet = next((name for name in excel.sheet_names if name.casefold() == "bpr"), None)
    if not sheet:
        raise ValueError("Published book reference workbook has no 'BPR' tab.")
    frame = pd.read_excel(excel, sheet_name=sheet, dtype=str)
    missing = REQUIRED_COLUMNS - set(frame.columns)
    if missing:
        raise ValueError(f"Published BPR reference is missing columns: {sorted(missing)}")
    return frame


def download_published_reference(url=REFERENCE_URL, timeout=30):
    """Download a fresh workbook for this run; no application cache is used."""
    request = Request(url, headers={"User-Agent": "BPR-Guidebook-Proofing/1.0"})
    with urlopen(request, timeout=timeout) as response:
        return load_published_workbook(response.read())


def match_book_reference(frame, lookup):
    matches = frame[frame[BOOK_NAME_COLUMN].map(_key) == _key(lookup)]
    if matches.empty:
        raise ValueError(f"No published BPR reference row matches Book value {lookup!r}.")

    relevant = []
    for _, row in matches.iterrows():
        item = {
            "book_name": _clean(row[BOOK_NAME_COLUMN]),
            "book_subheader": _clean(row[SUBHEADER_COLUMN]),
            "back_cover_qr_bitly": _clean(row[QR_COLUMN]),
        }
        if item not in relevant:
            relevant.append(item)
    if len(relevant) != 1:
        variants = [
            {"subheader": item["book_subheader"], "qr": item["back_cover_qr_bitly"]}
            for item in relevant
        ]
        raise ValueError(
            f"Published BPR reference has ambiguous rows for {lookup!r}: {variants}"
        )
    result = relevant[0]
    result.update({"lookup": lookup, "source_url": REFERENCE_URL})
    return result


def fetch_book_reference(bbb_path):
    lookup = book_lookup_from_file(bbb_path)
    return match_book_reference(download_published_reference(), lookup)


def _contains(text, expected):
    return _key(expected) in _key(text)


def _normalized_url(value):
    value = _clean(value)
    if not value:
        return ""
    parts = urlsplit(value)
    scheme = parts.scheme.casefold() or "https"
    host = parts.netloc.casefold()
    path = parts.path.rstrip("/")
    return urlunsplit((scheme, host, path, "", ""))


def check_book_reference(layout_pages, qr_result, reference):
    """Compare spine/front/back text and back-cover QR with the matched row."""
    last_page = max(layout_pages) if layout_pages else None
    book_name = reference["book_name"]
    subheader = reference["book_subheader"]

    checks = []

    def text_check(name, expected, page):
        found = layout_pages.get(page, "") if page else ""
        checks.append({
            "name": name,
            "status": "PASS" if _contains(found, expected) else "FIX",
            "expected": expected,
            "found": _clean(found)[:500],
            "page": page,
        })

    text_check("Spine book name", book_name, 1)
    text_check("Front cover book name", book_name, 2)
    if subheader:
        text_check("Front cover book sub-header", subheader, 2)
    else:
        checks.append({
            "name": "Front cover book sub-header",
            "status": "Not applicable",
            "expected": "(blank in published reference)",
            "found": "",
            "page": 2,
        })
    text_check("Back cover book name", book_name, last_page)

    expected_qr = reference["back_cover_qr_bitly"]
    if qr_result.get("skipped"):
        qr_status, decoded = "DID NOT RUN", []
    else:
        decoded = [
            destination
            for destination, pages in qr_result.get("destinations", {}).items()
            if last_page in [int(page) for page in pages]
        ]
        qr_status = (
            "PASS"
            if expected_qr and _normalized_url(expected_qr) in {_normalized_url(x) for x in decoded}
            else "FIX"
        )
    checks.append({
        "name": "Back cover QR Bitly",
        "status": qr_status,
        "expected": expected_qr,
        "found": ", ".join(decoded),
        "page": last_page,
    })

    return {
        "lookup": reference["lookup"],
        "source_url": reference["source_url"],
        "expected": {
            "book_name": book_name,
            "book_subheader": subheader,
            "back_cover_qr_bitly": expected_qr,
        },
        "checks": checks,
    }
