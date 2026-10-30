"""pyspellchecker English spell-checking for printer-proof PDF text.

Production runs use Poppler geometry-aware text blocks so columns and designed
text areas retain their own reading flow.
"""

from collections import Counter, defaultdict
from difflib import SequenceMatcher
from functools import lru_cache
from pathlib import Path
import re
import xml.etree.ElementTree as ET


WORD_RE = re.compile(r"(?<![\w@])([A-Za-z]+(?:['’][A-Za-z]+)*)(?![\w@])")
TOKEN_RE = re.compile(r"[A-Za-z][A-Za-z’'\-]*[A-Za-z]")
URL_RE = re.compile(r"[@/]|\.(com|net|org)")
CLITIC_RE = re.compile(r"'(s|ll|re|ve|d|t|m)$|'$")
CONTROL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
SOFT_WRAP_RE = re.compile(r"\b([A-Za-z]{2,})-\s+([A-Za-z]{2,})\b")

# Publishing and BPR terms that are valid but absent from a general dictionary.
ALLOWLIST = {
    "bitly", "guidebook", "guidebooks", "homeowner", "homeowners",
    "licensee", "licensees", "licensure", "qr", "salesforce", "utm", "utms",
}

# BBB file columns whose words are known-good (company names, categories).
REFERENCE_COLUMNS = ("Account Name", "PublishedName", "Category")

# Candidate split fragments for checking a lost hyphen against layout text.
# Suffix-like fragments don't count, otherwise 'warrenty' could be interpreted
# as 'warren' + 'ty'.
SUFFIXES = {"ed", "ing", "er", "ers", "ly", "ty", "es", "ion", "al"}
HYPHENS = r"[-‐‑‒–—]"


def _allowlist_words(entries):
    """Normalize entries and include components of hyphenated/slashed phrases."""
    words = set()
    for entry in entries:
        entry = entry.strip()
        if not entry or entry.startswith("#"):
            continue
        words.add(entry.lower().replace("’", "'"))
        words.update(
            match.group(1).lower().replace("’", "'")
            for match in WORD_RE.finditer(entry)
        )
    return words


@lru_cache(maxsize=1)
def load_allowlist():
    """Load the maintained project allow-list alongside this module."""
    path = Path(__file__).with_name("spellcheck_allowlist.txt")
    entries = path.read_text(encoding="utf-8-sig").splitlines() if path.is_file() else []
    return frozenset(_allowlist_words([*ALLOWLIST, *entries]))


def reference_words(ref_path):
    """Words from the BBB file's name/category columns; empty if it can't be read."""
    if not ref_path:
        return set()
    import pandas as pd

    try:
        if Path(ref_path).suffix.lower() in (".csv", ".txt"):
            df = pd.read_csv(ref_path, dtype=str, encoding="utf-8-sig")
        else:
            df = pd.read_excel(ref_path, dtype=str)
    except Exception:
        return set()
    cols = [c for c in REFERENCE_COLUMNS if c in df.columns]
    return _allowlist_words(df[cols].fillna("").to_numpy().ravel()) if cols else set()


def bbox_text_blocks(xml_text):
    """Return {PDF page: [text blocks]} from Poppler -bbox-layout XHTML."""
    # Poppler can preserve embedded control glyphs (for example backspace) in
    # word text even though XML 1.0 forbids them.
    root = ET.fromstring(CONTROL_RE.sub("", xml_text))
    pages = {}
    for page_number, page in enumerate(root.findall(".//{*}page"), start=1):
        blocks = []
        for block in page.findall(".//{*}block"):
            lines = []
            for line in block.findall("./{*}line"):
                words = [word.text or "" for word in line.findall("./{*}word")]
                if words:
                    lines.append(" ".join(words))
            if lines:
                blocks.append("\n".join(lines))
        pages[page_number] = blocks
    return pages


def _dictionary():
    from spellchecker import SpellChecker

    return SpellChecker(language="en")


def _context(flat, start, end, radius=45):
    left = max(0, start - radius)
    right = min(len(flat), end + radius)
    return ("..." if left else "") + flat[left:right].strip() + ("..." if right < len(flat) else "")


def _repair_soft_wraps(text, known):
    """Rejoin discretionary PDF wraps only when the result is a known word."""
    def replace(match):
        joined = match.group(1) + match.group(2)
        return joined if known(joined.lower()) else match.group(0)

    return SOFT_WRAP_RE.sub(replace, text)


def _suggestions(dictionary, word):
    candidates = dictionary.candidates(word) or set()
    candidates = {candidate for candidate in candidates if candidate != word}
    return sorted(candidates, key=lambda candidate: (-dictionary.word_frequency[candidate], candidate))[:3]


def _twin(word, count, seen, dictionary):
    """A near-identical name used more often in the same book ('Sulfrigde' vs 'Sulfridge')."""
    lower = word.lower()
    for other, n in seen.most_common():
        if n <= count:
            break
        if (other != word and other[:2] == word[:2] and len(other) >= 5
                and abs(len(other) - len(word)) <= 1 and other.lower() not in dictionary
                and SequenceMatcher(None, other.lower(), lower).ratio() >= 0.85):
            return other
    return None


def check_pages(pages, dictionary=None, allowlist=None, layout_pages=None, line_mode=False):
    """Return one Cobb-style finding per distinct unknown word.

    pages      {pdf page number: text}, or {page: [geometry-aware text blocks]}
    allowlist  extra known-good words for this run, e.g. reference_words(bbb_file)
    layout_pages  matching layout-preserved text used to verify printed hyphens
    line_mode  check each extracted text line independently instead of flattening a page
    """
    dictionary = dictionary or _dictionary()
    allowed = set(load_allowlist())
    allowed.update(_allowlist_words(allowlist or []))

    def known(word):
        return word in allowed or word in dictionary

    def common(part):
        if len(part) < 2 or part in SUFFIXES:
            return False
        return part in allowed or part in dictionary

    def printed_with_hyphen(word, layout_text):
        """True only when the same page proves the concatenation had a hyphen."""
        for index in range(2, len(word) - 1):
            left, right = word[:index], word[index:]
            if not (common(left) and common(right)):
                continue
            pattern = rf"\b{re.escape(left)}\s*{HYPHENS}\s*{re.escape(right)}\b"
            if re.search(pattern, layout_text or "", re.IGNORECASE):
                return True
        return False

    occurrences = defaultdict(list)
    seen = Counter()                     # every word in the book, for the consistency check
    for page, page_text in pages.items():
        text_blocks = page_text if isinstance(page_text, (list, tuple)) else [page_text]
        for text in text_blocks:
            cleaned = CONTROL_RE.sub(" ", text or "")
            cleaned = _repair_soft_wraps(cleaned, known)
            segments = cleaned.splitlines() if line_mode else [cleaned]
            for segment in segments:
                flat = re.sub(r"\s+", " ", segment).strip()
                if not flat:
                    continue
                for match in TOKEN_RE.finditer(flat):
                    token = match.group(0).replace("’", "'")
                    if URL_RE.search(flat[max(0, match.start() - 1):match.end() + 5]):
                        continue
                    if known(token.lower()):     # doesn't, I'm, you'll, side-by-side ...
                        seen[token] += 1
                        continue
                    for word in token.split("-"):
                        word = CLITIC_RE.sub("", word)
                        if len(word) < 3 or word.isupper():
                            continue
                        seen[word] += 1
                        lower = word.lower()
                        layout_text = (layout_pages or {}).get(page, "")
                        if known(lower) or printed_with_hyphen(lower, layout_text):
                            continue
                        occurrences[word].append({
                            "page": int(page),
                            "context": _context(flat, match.start(), match.end()),
                        })

    findings = []
    for word, items in occurrences.items():
        lower = word.lower()
        count = len(items)
        proper = word[:1].isupper()
        suggestions = _suggestions(dictionary, lower)
        twin = _twin(word, count, seen, dictionary) if proper else None

        if twin:
            issue = f"Spelling: Inconsistent spelling -- '{twin}' appears {seen[twin]}x elsewhere"
            priority, expected = "P2", twin
        elif not proper and suggestions and count <= 2:
            issue = "Spelling: Likely misspelling"
            priority, expected = "P2", ", ".join(suggestions)
        elif not proper:
            issue = "Spelling: Joined words -- hyphen lost at a line break, or missing in copy?"
            priority, expected = "P4", ""
        else:
            # Unknown proper names/brands are expected in these books. Hide both
            # one-off and consistently used names unless a near-identical spelling
            # elsewhere makes the difference actionable (the P2 twin case above).
            continue

        context = items[0]["context"]
        if count > 1:
            context += f"  (+{count - 1} more)"
        findings.append({
            "Sheet": "Book",
            "Row": "",
            "Key": word,
            "Issue": issue,
            "Expected": expected,
            "Found": context,
            "Page": ", ".join(f"pdf {page}" for page in sorted({item["page"] for item in items})),
            "Source": "PrinterProof / pyspellchecker English",
            "Priority": priority,
        })

    priority_order = {"P2": 0, "P3": 1, "P4": 2}
    findings.sort(key=lambda item: (priority_order[item["Priority"]], item["Key"].lower()))
    return findings
