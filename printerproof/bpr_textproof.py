#!/usr/bin/env python3
"""
Best Pick Reports — PRE-PRINTER (text-layer) proof QA

    python3 bpr_textproof.py Austin_Fall_2026_Book_v2.pdf --out ./results/austin

For PDFs exported straight from InDesign, which carry a live text layer. If the
file is a flattened raster proof this script says so and stops -- use
bpr_proof.py instead. Do not "fall back" to OCR here: the whole point of this
pipeline is that nothing is inferred.
"""

import argparse
import collections
import glob
import json
import os
import re
import math
import statistics as st
import subprocess
import sys
import uuid

from .spellcheck import bbox_text_blocks, check_pages

import cv2
import numpy as np

CONFIG = {
    "dpi_qr": 300,
    "qr_tile_win": 520,
    "qr_tile_step": 130,
    "qr_tile_pad": 50,

    "kern_min_word_len": 5,
    "kern_min_words_per_line": 4,
    "kern_min_bucket": 40,
    "kern_page_flag": 0.96,
    "kern_line_flag": 0.88,

    "dup_overlap_flag": 0.50,

    "licence_band": 22,
    "licence_x_tol": 40,
    "licence_outlier_max_count": 2,
    "licence_outlier_max_share": 0.20,

    "directory_min_page_entries": 6,
    "directory_page_share": 0.15,
    "folio_band": (700, 800),
    "folio_edge": 0.25,
    "folio_height_tol": 0.12,
}

PHONE = re.compile(r'\b(\d{3})[-–—](\d{3})[-–—](\d{4})\b')
TENURE = re.compile(r'(\d{1,2})(st|nd|rd|th) year as a Best Pick', re.I)
CALL_LINE = re.compile(r'or call\s+(\d{3}[-–—]\d{3}[-–—]\d{4})\s+for service', re.I)
QR_BLOCK = re.compile(r'Scan QR Code', re.I)
QR_EDU = re.compile(r'Scan for\s+additional', re.I | re.S)
QR_OTHER = re.compile(r'Scan for Free Access|scan the\s+QR code', re.I | re.S)
LISTING = re.compile(r'^\s*■?\s*(.+?)\s*(\d{3}-\d{3}-\d{4})\s*$', re.M)
LEADER = re.compile(r'\s*[.]\s*(?:[.]\s*){2,}\s*(\d{1,3})(?=\s|$)')


def sh(cmd):
    return subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )


def _tmp(workdir, stem, ext=""):
    d = os.path.join(workdir, "tmp")
    os.makedirs(d, exist_ok=True)
    return os.path.join(d, f"{stem}-{os.getpid()}-{uuid.uuid4().hex[:8]}{ext}")


def page_count(pdf):
    return int(re.search(r'^Pages:\s+(\d+)', sh(["pdfinfo", pdf]).stdout, re.M).group(1))


def has_text_layer(pdf):
    out = sh(["pdffonts", pdf]).stdout.strip().split("\n")
    if len(out) <= 2:
        return False
    probe = sh(["pdftotext", "-f", "1", "-l", "8", pdf, "-"]).stdout
    return len(re.sub(r'\s', '', probe)) > 200


def text_pages(pdf, n, layout=True):
    out = {}
    for i in range(1, n + 1):
        cmd = ["pdftotext"] + (["-layout"] if layout else []) + \
              ["-f", str(i), "-l", str(i), pdf, "-"]
        out[i] = sh(cmd).stdout
    return out


# ------------------------------------------------------------------ geometry
def word_boxes(pdf, workdir):
    p = _tmp(workdir, "bbox", ".xml")
    sh(["pdftotext", "-bbox-layout", pdf, p])
    d = open(p, encoding="utf-8", errors="replace").read()
    os.remove(p)
    text_blocks = bbox_text_blocks(d)
    pages = re.findall(r'<page width="([\d.]+)" height="([\d.]+)">(.*?)</page>', d, re.S)
    LINE = re.compile(r'<line [^>]*>(.*?)</line>', re.S)
    WORD = re.compile(r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">(.*?)</word>')
    out = []
    for pi, (w, h, body) in enumerate(pages, start=1):
        lines = []
        for lm in LINE.findall(body):
            ws = [(float(a), float(b), float(c), float(dd),
                   re.sub(r'&[a-z]+;', 'x', e)) for a, b, c, dd, e in WORD.findall(lm)]
            lines.append(ws)
        out.append({"page": pi, "w": float(w), "h": float(h), "lines": lines})
    return out, text_blocks


def inventory(geo):
    y0, y1 = CONFIG["folio_band"]
    cand = {}
    for g in geo:
        words = [w for ln in g["lines"] for w in ln]
        rows = []
        for x0, yy0, x1, yy1, t in words:
            if not (t.isdigit() and len(t) <= 3 and y0 < yy0 < y1):
                continue
            sd = ("L" if x1 < g["w"] * CONFIG["folio_edge"]
                  else ("R" if x0 > g["w"] * (1 - CONFIG["folio_edge"]) else "C"))
            if sd == "C":
                continue
            rows.append((round(yy1 - yy0, 1), int(t), sd))
        cand[g["page"]] = rows

    heights = collections.Counter(h for rows in cand.values() for h, _, _ in rows)
    modal = heights.most_common(1)[0][0] if heights else None
    tol = CONFIG["folio_height_tol"]

    inv = []
    for g in geo:
        words = [w for ln in g["lines"] for w in ln]
        indb = None
        for k, wd in enumerate(words):
            if wd[4] == "Book.indb" and k + 1 < len(words) and words[k + 1][4].isdigit():
                indb = int(words[k + 1][4])
        folio = side = None
        if modal:
            keep = [c for c in cand[g["page"]] if abs(c[0] - modal) / modal <= tol]
            if keep:
                keep.sort(key=lambda c: -c[0])
                folio, side = keep[0][1], keep[0][2]
        inv.append({"pdf": g["page"], "indb": indb, "folio": folio, "side": side,
                    "folio_h": modal})
    return inv


def check_folios(inv):
    body = [r for r in inv if r["folio"] is not None]
    if not body:
        return {"ok": 0, "issues": [["no folios found"]], "first": None, "last": None, "offset": None, "unnumbered": []}
    offsets = collections.Counter(r["pdf"] - r["folio"] for r in body)
    offset = offsets.most_common(1)[0][0]
    issues, ok = [], 0
    prev = None
    for r in body:
        exp_side = "L" if r["folio"] % 2 == 0 else "R"
        bad = []
        if r["pdf"] - r["folio"] != offset:
            bad.append(f"offset {r['pdf'] - r['folio']} != book offset {offset}")
        if r["side"] != exp_side:
            bad.append(f"folio {r['folio']} on {r['side']}, expected {exp_side}")
        if prev is not None and r["folio"] != prev + 1:
            bad.append(f"sequence break: {prev} -> {r['folio']}")
        prev = r["folio"]
        if bad:
            issues.append([r["pdf"], r["folio"]] + bad)
        else:
            ok += 1
    return {"ok": ok, "issues": issues, "first": body[0]["folio"],
            "last": body[-1]["folio"], "offset": offset,
            "unnumbered": [r["pdf"] for r in inv if r["folio"] is None]}


def check_assembly(inv):
    seen = collections.defaultdict(list)
    for r in inv:
        if r["indb"] is not None:
            seen[r["indb"]].append(r["pdf"])
    dup = {k: v for k, v in seen.items() if len(v) > 1}
    stamps = [r["indb"] for r in inv if r["indb"] is not None]
    return {"physical_pages": len(inv),
            "max_stamp": max(stamps) if stamps else None,
            "duplicate_stamps": dup,
            "divisible_by_4": len(inv) % 4 == 0}


# -------------------------------------------------------------------- content
def category_pages(lay):
    cats = {}
    for i, t in lay.items():
        if "Best Pick Reports recommends" not in t:
            continue
        for l in [x.strip() for x in t.split("\n")]:
            if (l and "recommends" not in l and not l.startswith("Austin Fall")
                    and not re.match(r'^\w+ \w+ \d{4} Book\.indb', l)
                    and not l.isdigit() and "TRADE LICENSE" not in l):
                cats[i] = l
                break
    return cats


def collect_listings(lay, cats):
    listing = collections.OrderedDict()
    for i, cat in cats.items():
        rows = LISTING.findall(lay[i])
        tens = [int(m.group(1)) for m in TENURE.finditer(lay[i])]
        out = []
        for k, (name, ph) in enumerate(rows):
            out.append({"company": name.strip(" ■"), "phone": ph,
                        "tenure": tens[k] if k < len(tens) else None})
        listing[cat] = {"page": i, "rows": out}
    return listing


def check_phones(raw):
    parsed = expected = 0
    bad = []
    for i, t in raw.items():
        flat = re.sub(r'\s+', ' ', t)
        expected += len(QR_BLOCK.findall(flat))
        for m in CALL_LINE.finditer(flat):
            parsed += 1
            prev = PHONE.findall(flat[:m.start()])
            if prev:
                hdr = "-".join(prev[-1])
                call = m.group(1).replace("–", "-").replace("—", "-")
                if hdr != call:
                    bad.append([i, hdr, call])
    return {"parsed": parsed, "expected": expected, "mismatches": bad}


def check_ratings(raw):
    checked, fails, unread = 0, [], []
    for i, t in raw.items():
        runs, cur = [], []
        for l in [x.strip() for x in t.split("\n")]:
            if re.fullmatch(r'\d{1,3}(,\d{3})*|\d{1,6}', l):
                cur.append(int(l.replace(",", "")))
            elif l == "":
                continue
            else:
                if len(cur) >= 5:
                    runs.append(cur[:5])
                cur = []
        if len(cur) >= 5:
            runs.append(cur[:5])
        printed = [float(x) for x in re.findall(r'Rating:\s*([\d.]+)\s*out of 5', t)]
        if len(runs) != len(printed):
            if printed:
                unread.append([i, len(runs), len(printed)])
            continue
        for c, p in zip(runs, printed):
            tot = sum(c)
            if not tot:
                unread.append([i, "zero-total histogram", c])
                continue
            avg = (c[0]*5 + c[1]*4 + c[2]*3 + c[3]*2 + c[4]*1) / tot
            checked += 1
            if abs(math.floor(avg * 10 + 0.5) / 10 - p) > 0.051:
                fails.append([i, c, round(avg, 3), p])
    return {"checked": checked, "mismatches": fails, "unreadable": unread}


def check_tenure(listing, edition):
    order, over, cross = [], [], []
    comp = collections.defaultdict(dict)
    for cat, d in listing.items():
        tens = [r["tenure"] for r in d["rows"] if r["tenure"]]
        if tens != sorted(tens, reverse=True):
            order.append([cat, d["page"], tens])
        for r in d["rows"]:
            if r["tenure"]:
                if edition and r["tenure"] > edition:
                    over.append([cat, r["company"], r["tenure"], edition])
                comp[r["company"]][cat] = r["tenure"]
    for name, d in comp.items():
        if len(d) > 1 and len(set(d.values())) > 1:
            cross.append([name, d])
    claims = [r["tenure"] for d in listing.values() for r in d["rows"] if r["tenure"]]
    return {"order_issues": order, "over_ceiling": over, "cross_category": cross,
            "max_claim": max(claims) if claims else None, "edition": edition,
            "multi_category_companies": sum(1 for d in comp.values() if len(d) > 1)}


def check_directory(lay, raw, listing, inv):
    body = {}
    for cat, d in listing.items():
        for r in d["rows"]:
            body[r["phone"]] = (cat, r["company"])
    cand = [r["pdf"] for r in inv if r["folio"] is None and r["pdf"] > len(inv) * 0.7]
    per_page = {}
    for p in cand:
        per_page[p] = {"-".join(x) for x in PHONE.findall(raw.get(p, ""))}
    floor = max(CONFIG["directory_min_page_entries"],
                len(body) * CONFIG["directory_page_share"])
    sheet_pages = sorted(p for p in cand if len(per_page[p]) >= floor)
    sheet = set()
    for p in sheet_pages:
        sheet |= per_page[p]
    sheet.discard("877-906-7575")
    sheet_text = re.sub(r'\W', '', "".join(raw.get(p, "") for p in sheet_pages)).lower()
    cats_missing = [c for c in listing
                    if re.sub(r'\W', '', c).lower() not in sheet_text]
    return {"sheet_pages": sheet_pages,
            "sheet_page": sheet_pages[0] if sheet_pages else None,
            "entries_per_sheet_page": {p: len(per_page[p]) for p in sheet_pages},
            "sheet_count": len(sheet),
            "listing_count": len(body),
            "missing_from_sheet": sorted(set(body) - sheet),
            "orphans_on_sheet": sorted(sheet - set(body)),
            "categories_missing_from_sheet": cats_missing}


def check_toc(lay, listing, raw, inv):
    starts = {}
    for cat, d in listing.items():
        folio = next((r["folio"] for r in inv if r["pdf"] == d["page"]), None)
        if folio:
            starts[re.sub(r'\W', '', cat).lower()] = folio
    ALIAS = {"frequentlyaskedquestions": ["faqs"]}
    for p, t in sorted(lay.items()):
        head = None
        for l in [x.strip() for x in t.split("\n")]:
            if l and not re.match(r'^.*Book\.indb\s+\d+', l) and not l.isdigit():
                head = l.rstrip(". ")
                break
        if not head or head.endswith(":"):
            continue
        for name in ("Additional Information", "Homeowner Protection Tips",
                     "Seasonal Maintenance Checklist", "Frequently Asked Questions",
                     "Quick Reference Sheet"):
            if re.sub(r'\W', '', head).lower() != re.sub(r'\W', '', name).lower():
                continue
            k = re.sub(r'\W', '', name).lower()
            f = next((r["folio"] for r in inv if r["pdf"] == p), None)
            if f:
                starts.setdefault(k, f)
                for al in ALIAS.get(k, []):
                    starts.setdefault(al, f)
    qrs = next((x["pdf"] for x in inv if x["folio"] is None and x["pdf"] > len(inv) * 0.7), None)
    if qrs:
        prev = next((x["folio"] for x in inv if x["pdf"] == qrs - 1), None)
        if prev:
            starts.setdefault("quickreferencesheet", prev + 1)

    def parse(t):
        t = re.sub(r'(Damage)\s*\n\s*(Restoration)', r'\1 \2', t)
        ents = []
        for line in t.split("\n"):
            pos = 0
            for m in LEADER.finditer(line):
                lab = re.sub(r'\s+', ' ', line[pos:m.start()]).strip(" .")
                pos = m.end()
                if lab and "Book.indb" not in lab:
                    ents.append((lab, int(m.group(1))))
        return ents

    results = {}
    for p, t in lay.items():
        ents = parse(t)
        if len(ents) < 10:
            continue
        ok, bad, unres = 0, [], []
        for lab, num in ents:
            k = re.sub(r'\W', '', lab).lower()
            a = starts.get(k)
            if a is None:
                if num in starts.values():
                    ok += 1
                else:
                    unres.append([lab, num])
                continue
            if a != num:
                bad.append([lab, num, a])
            else:
                ok += 1
        results[p] = {"entries": len(ents), "correct": ok,
                      "mismatched": bad, "unresolved": unres}
    return results


def check_characters(raw):
    straight, nested, glyphs = [], [], []
    dashes = collections.Counter()
    for i, t in raw.items():
        flat = t.replace("\n", " ")
        for m in re.finditer(r".{0,32}\w'\w.{0,28}", flat):
            straight.append([i, re.sub(r'\s+', ' ', m.group(0))])
        for m in re.finditer(r'[“"][‘’\'][^”"]{0,70}[‘’\'][”"]', flat):
            nested.append([i, re.sub(r'\s+', ' ', m.group(0))])
        for m in re.finditer(r'.{0,25}[-].{0,25}', flat):
            glyphs.append([i, re.sub(r'\s+', ' ', m.group(0))])
        for ch in t:
            if ch in '-‐‑‒–—­−':
                dashes[ch] += 1
    return {"straight_apostrophes": straight, "nested_quotes": nested,
            "private_use_glyphs": glyphs,
            "dash_census": {f"U+{ord(k):04X}": v for k, v in dashes.items()}}


def check_text_hygiene(raw):
    ordinals, doubled, dups = [], [], []
    for i, t in raw.items():
        flat = re.sub(r'\s+', ' ', t)
        for m in re.finditer(r'\b(\d{1,3})(st|nd|rd|th)\b', flat, re.I):
            n, s = int(m.group(1)), m.group(2).lower()
            exp = "th" if 11 <= n % 100 <= 13 else {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
            if s != exp:
                ordinals.append([i, m.group(0)])
        for line in t.split("\n"):
            for m in re.finditer(r'\b([A-Za-z]{3,}) +\1\b', line):
                doubled.append([i, m.group(0)])

    def ws(t):
        return {w for w in re.sub(r'[^a-z0-9 ]', ' ', t.lower()).split() if len(w) > 3}
    ks = sorted(raw)
    for a, b in zip(ks, ks[1:]):
        x, y = ws(raw[a]), ws(raw[b])
        if x and y:
            j = len(x & y) / len(x | y)
            if j > CONFIG["dup_overlap_flag"]:
                dups.append([a, b, round(j, 2)])
    return {"ordinals": ordinals, "doubled_words": doubled, "duplicate_pages": dups}


LIC_LABEL = re.compile(r'^Verified(?!.*nsurance)[A-Za-z]{0,28}'
                       r'(License\(s\)|Licence\(s\)|Licenses|License|Registration)$', re.I)


def _page_lines(g):
    out = []
    for ln in g["lines"]:
        if ln:
            out.append((min(w[1] for w in ln), min(w[0] for w in ln),
                        " ".join(w[4] for w in ln)))
    out.sort()
    return out


def check_licences(geo):
    pairs, missing = [], []
    band, xtol = CONFIG["licence_band"], CONFIG["licence_x_tol"]
    for g in geo:
        lines = _page_lines(g)
        for y, x, txt in lines:
            if not LIC_LABEL.match(re.sub(r'\s+', '', txt)):
                continue
            val = None
            for y2, x2, t2 in lines:
                if not (0 < y2 - y < band and abs(x2 - x) < xtol):
                    continue
                if not re.search(r'\d', t2):
                    continue
                if re.match(r'^(Verified|Best Pick|Scan|or call|Rating|Trade License)', t2):
                    continue
                val = t2.strip()
                break
            if val:
                pairs.append([g["page"], val])
            else:
                missing.append([g["page"], txt.strip()])

    lic = collections.defaultdict(list)
    for page, val in pairs:
        head = val.split(",")[0].strip()
        m = re.match(r'^([A-Za-z]{1,6})[- ]?([\d-]{3,})', head)
        cls = m.group(1).upper() if m else "NUM"
        lic[cls].append([page, val, len(re.sub(r'\D', '', head))])

    outliers, spread = [], {}
    for pre, rows in lic.items():
        uniq = {}
        for page, s_, ln in rows:
            uniq.setdefault(s_, (page, ln))
        lens = collections.Counter(v[1] for v in uniq.values())
        spread[pre] = dict(sorted(lens.items()))
        if len(lens) > 1:
            top = lens.most_common()
            if len(top) > 1 and top[0][1] == top[1][1]:
                continue
            common = top[0][0]
            for s_, (page, ln) in uniq.items():
                if ln == common:
                    continue
                share = lens[ln] / len(uniq)
                if lens[ln] <= CONFIG["licence_outlier_max_count"] and \
                        share <= CONFIG["licence_outlier_max_share"]:
                    outliers.append([pre, page, s_, ln, common])
    total = sum(len({r[1] for r in v}) for v in lic.values())
    for pre, rows in lic.items():
        n = len({r[1] for r in rows})
        if total >= 10 and n / total < 0.05:
            outliers.append([pre, rows[0][0], sorted({r[1] for r in rows}),
                             f"{n} of {total} licences in this book", ""])
    return {"classes": {k: sorted({r[1] for r in v}) for k, v in lic.items()},
            "length_spread": spread,
            "values_read": len(pairs),
            "labels_without_value": missing,
            "format_outliers": outliers}


def check_tracking(geo):
    rows = []
    for g in geo:
        for li, ln in enumerate(g["lines"]):
            if len(ln) < 4:
                continue
            for x0, y0, x1, y1, t in ln:
                if not re.fullmatch(r"[A-Za-z][A-Za-z’'\-]{%d,}" % (CONFIG["kern_min_word_len"] - 1), t):
                    continue
                w, h = x1 - x0, y1 - y0
                if h < 4 or w < 4:
                    continue
                rows.append((g["page"], li, t, round(h, 1), (w / len(t)) / h))
    b = collections.defaultdict(list)
    for _, _, _, h, r in rows:
        b[h].append(r)
    med = {h: st.median(v) for h, v in b.items() if len(v) >= CONFIG["kern_min_bucket"]}
    lines = collections.defaultdict(list)
    for p, li, t, h, r in rows:
        if h in med:
            lines[(p, li)].append(r / med[h])
    lscore = {k: st.median(v) for k, v in lines.items()
              if len(v) >= CONFIG["kern_min_words_per_line"]}
    pg = collections.defaultdict(list)
    for (p, _), s in lscore.items():
        pg[p].append(s)
    pscore = {p: st.median(v) for p, v in pg.items() if len(v) >= 5}
    return {"words": len(rows), "lines": len(lscore),
            "median": round(st.median(list(lscore.values())), 4) if lscore else None,
            "stdev": round(st.pstdev(list(lscore.values())), 4) if lscore else None,
            "flagged_pages": sorted([round(s, 3), p] for p, s in pscore.items()
                                    if s < CONFIG["kern_page_flag"]),
            "flagged_lines": sorted([round(s, 3), list(k)] for k, s in lscore.items()
                                    if s < CONFIG["kern_line_flag"])[:20]}


# ------------------------------------------------------------------------ QR
def _decode_full(img, det):
    ok, dec, _, _ = det.detectAndDecodeMulti(img)
    return {d for d in dec if d} if ok else set()


def _decode_tiles(img, det):
    win, step, pad = CONFIG["qr_tile_win"], CONFIG["qr_tile_step"], CONFIG["qr_tile_pad"]
    h, w = img.shape[:2]
    hits = set()
    for y in range(0, max(1, h - win // 2), step):
        for x in range(0, max(1, w - win // 2), step):
            t = img[y:y + win, x:x + win]
            if t.shape[0] < 80 or t.shape[1] < 80:
                continue
            t = cv2.copyMakeBorder(t, pad, pad, pad, pad, cv2.BORDER_CONSTANT, value=(255, 255, 255))
            d, _, _ = det.detectAndDecode(t)
            if d:
                hits.add(d)
    return hits


def check_qr(pdf, raw, workdir, verbose=True):
    det = cv2.QRCodeDetector()
    found, per_page, rescued = {}, {}, []
    for i, t in raw.items():
        flat = re.sub(r'\s+', ' ', t)
        exp = len(QR_BLOCK.findall(flat)) + len(QR_EDU.findall(flat)) + len(QR_OTHER.findall(flat))
        if exp == 0:
            continue
        pre = _tmp(workdir, f"qr{i}")
        sh(["pdftoppm", "-png", "-r", str(CONFIG["dpi_qr"]), "-f", str(i), "-l", str(i), pdf, pre])
        f = sorted(glob.glob(pre + "*.png"))
        if not f:
            continue
        img = cv2.imread(f[0])
        hits = _decode_full(img, det)
        if len(hits) < exp:
            more = _decode_tiles(img, det)
            if more - hits:
                rescued.append([i, sorted(more - hits)])
            hits |= more
        for x in f:
            os.remove(x)
        per_page[i] = {"expected": exp, "decoded": sorted(hits)}
        for h in hits:
            found.setdefault(h, []).append(i)
        if verbose:
            print(f"  QR page {i}: {len(hits)}/{exp}", end="\r", flush=True)
    if verbose:
        print()
    exp_tot = sum(v["expected"] for v in per_page.values())
    got_tot = sum(len(v["decoded"]) for v in per_page.values())
    return {"expected": exp_tot, "decoded": got_tot, "unique": len(found),
            "duplicates": {k: v for k, v in found.items() if len(v) > 1},
            "short_by_page": {k: v for k, v in per_page.items()
                              if len(v["decoded"]) < v["expected"]},
            "rescued_by_tiling": rescued,
            "destinations": {k: v for k, v in found.items()}}


# ----------------------------------------------------------------------- run
def run(pdf, outdir, skip_qr=False, spell_words=None):
    os.makedirs(outdir, exist_ok=True)
    workdir = os.path.join(outdir, "_work")
    os.makedirs(os.path.join(workdir, "tmp"), exist_ok=True)
    name = os.path.basename(pdf)
    n = page_count(pdf)
    print(f"{name}: {n} pages")

    if not has_text_layer(pdf):
        print("  !! NO TEXT LAYER — this is a flattened raster proof.")
        sys.exit(2)
    print("  text layer present — extraction is exact, no OCR")

    lay = text_pages(pdf, n, layout=True)
    raw = text_pages(pdf, n, layout=False)
    geo, text_blocks = word_boxes(pdf, workdir)
    inv = inventory(geo)

    cats = category_pages(lay)
    listing = collect_listings(lay, cats)

    edition = None
    for i in (1, 2, 3):
        m = re.search(r'(\d{1,2})(?:st|nd|rd|th)\s+Annual', raw.get(i, ""))
        if m:
            edition = int(m.group(1))
            break
    metro = None
    first = [l.strip() for l in lay.get(1, "").split("\n") if l.strip()]
    if first:
        metro = first[0]

    r = {
        "pdf": name, "pages": n, "edition": edition, "cover_metro": metro,
        "assembly": check_assembly(inv),
        "folios": check_folios(inv),
        "inventory": inv,
        "categories": len(listing),
        "companies": sum(len(d["rows"]) for d in listing.values()),
        "listing": {k: v["rows"] for k, v in listing.items()},
        "phones": check_phones(raw),
        "ratings": check_ratings(raw),
        "tenure": check_tenure(listing, edition),
        "directory": check_directory(lay, raw, listing, inv),
        "toc": check_toc(lay, listing, raw, inv),
        "characters": check_characters(raw),
        "hygiene": check_text_hygiene(raw),
        "licences": check_licences(geo),
        "tracking": check_tracking(geo),
        # Geometry-aware blocks preserve each column/design area's reading flow
        # and allow wrapped words to rejoin without mixing adjacent columns.
        "spellcheck": check_pages(text_blocks, allowlist=spell_words),
        "spellcheck_blocks": sum(len(blocks) for blocks in text_blocks.values()),
    }
    print("Decoding QR codes..." if not skip_qr else "Skipping QR pass")
    r["qr"] = {"skipped": True} if skip_qr else check_qr(pdf, raw, workdir)

    with open(os.path.join(outdir, "results.json"), "w") as f:
        json.dump(r, f, indent=1)
    report(r)
    return r


def report(r):
    W = 74
    print("\n" + "=" * W)
    print(f"RESULTS — {r['pdf']}  ({r['pages']}pp, {r['cover_metro']}, "
          f"{r['edition']}th Annual)")
    print("=" * W)
    a = r["assembly"]
    print(f"  Assembly            : {a['physical_pages']} physical pages, "
          f"InDesign stamps to {a['max_stamp']}, "
          f"{'divisible by 4' if a['divisible_by_4'] else 'NOT divisible by 4'}")
    if a["duplicate_stamps"]:
        print(f"    !! pages sharing a stamp: {a['duplicate_stamps']}")
    f = r["folios"]
    print(f"  Folios              : {f['first']}-{f['last']}, {f['ok']} correct, "
          f"{len(f['issues'])} issues, {len(f.get('unnumbered', []))} unnumbered pages")
    for x in f["issues"][:6]:
        print(f"    {x}")
    for p, t in r["toc"].items():
        print(f"  TOC (pdf p{p})       : {t['correct']}/{t['entries']} correct, "
              f"{len(t['mismatched'])} mismatched, {len(t['unresolved'])} unresolved")
        for x in t["mismatched"][:6]:
            print(f"    MISMATCH {x}")
    ph = r["phones"]
    print(f"  Phone integrity     : {len(ph['mismatches'])} mismatches across "
          f"{ph['parsed']} of {ph['expected']} call-lines")
    ra = r["ratings"]
    print(f"  Ratings vs histogram: {len(ra['mismatches'])} mismatches over "
          f"{ra['checked']} recomputed, {len(ra['unreadable'])} unread")
    for x in ra["mismatches"][:6]:
        print(f"    MISMATCH {x}")
    d = r["directory"]
    print(f"  Directory sheet     : {d['sheet_count']} on sheet vs "
          f"{d['listing_count']} in listings")
    if d["missing_from_sheet"]:
        print(f"    missing from sheet: {d['missing_from_sheet']}")
    if d["orphans_on_sheet"]:
        print(f"    ORPHANS on sheet  : {d['orphans_on_sheet']}")
    t = r["tenure"]
    print(f"  Tenure              : {len(t['order_issues'])} order issues, "
          f"max claim {t['max_claim']} vs edition {t['edition']}, "
          f"{len(t['over_ceiling'])} over ceiling")
    print(f"    cross-category differences: {len(t['cross_category'])} of "
          f"{t['multi_category_companies']} multi-category companies")
    c = r["characters"]
    print(f"  Characters          : {len(c['straight_apostrophes'])} straight "
          f"apostrophes, {len(c['nested_quotes'])} doubled quote marks, "
          f"{len(c['private_use_glyphs'])} symbol-font glyphs")
    h = r["hygiene"]
    print(f"  Hygiene             : {len(h['ordinals'])} bad ordinals, "
          f"{len(h['doubled_words'])} doubled words, "
          f"{len(h['duplicate_pages'])} duplicate-page flags")
    li = r["licences"]
    print(f"  Licences            : {li['values_read']} values read, classes "
          f"{ {k: len(v) for k, v in li['classes'].items()} }")
    if li["labels_without_value"]:
        print(f"    !! {len(li['labels_without_value'])} licence labels print with NO number: "
              f"{li['labels_without_value'][:6]}")
    for x in li["format_outliers"]:
        print(f"    OUTLIER {x}")
    k = r["tracking"]
    print(f"  Tracking            : median {k['median']} (stdev {k['stdev']}) over "
          f"{k['lines']} lines / {k['words']} words; "
          f"{len(k['flagged_pages'])} pages below {CONFIG['kern_page_flag']}")
    for x in k["flagged_pages"][:5]:
        print(f"    {x}")
    q = r["qr"]
    if not q.get("skipped"):
        print(f"  QR codes            : {q['decoded']}/{q['expected']} decoded, "
              f"{q['unique']} unique destinations")
        if q["duplicates"]:
            print(f"    !! DUPLICATE short links: {q['duplicates']}")
        if q["short_by_page"]:
            print(f"    !! pages still short: {list(q['short_by_page'])}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Best Pick Reports pre-printer proof QA")
    ap.add_argument("pdf")
    ap.add_argument("--out", default="./textproof_results")
    ap.add_argument("--skip-qr", action="store_true")
    a = ap.parse_args()
    if not os.path.exists(a.pdf):
        sys.exit(f"not found: {a.pdf}")
    run(a.pdf, a.out, a.skip_qr)
