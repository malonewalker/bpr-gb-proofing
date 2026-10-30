#!/usr/bin/env python3
"""
printerproof_adapter.py — merge printer-proof results into the proofing apps' output.

    python3 printerproof_adapter.py results.json --app <app workbook> --brand BPR|FSR --out merged.xlsx

Writes one workbook:
  Errors    — the apps' existing schema (Sheet, Row, Key, Issue, Expected, Found, Page, Source)
              plus a Priority column, with the app's own rows kept and printer-proof rows added.
  Coverage  — one row per check: what ran, how much it covered, PASS / FIX / VERIFY /
              NEEDS MANUAL CHECK / N/A. A check that did not run is listed, never silent.
  (the app's other tabs are copied through untouched)

Priority scheme is the one the team already uses: P1 blocks sign-off, P2 fix before print,
P3 confirm against source data, P4 informational.
"""
import argparse
import json
import re

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

COLS = ["Sheet", "Row", "Key", "Issue", "Expected", "Found", "Page", "Source", "Priority"]
FILL = {"P1": "FCE4D6", "P2": "FFF2CC", "P3": "DDEBF7", "P4": "E2EFDA",
        "FIX": "F8CBAD", "VERIFY": "FFE699", "PASS": "E2EFDA", "PASS (partial)": "E2EFDA",
        "NEEDS MANUAL CHECK": "FCE4D6", "Not applicable": "EEEEEE", "DID NOT RUN": "FCE4D6"}


_CTRL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _clean(v):
    return _CTRL.sub("", v) if isinstance(v, str) else v


def row(key, issue, expected="", found="", page=None, prio="P3", sheet="Book"):
    key, issue, expected, found = map(_clean, (key, issue, expected, found))
    return {"Sheet": sheet, "Row": "", "Key": key, "Issue": issue, "Expected": expected,
            "Found": found, "Page": page, "Source": "PrinterProof", "Priority": prio}


def folio_of(r, pdf_page):
    for x in r.get("inventory", []):
        if x["pdf"] == pdf_page:
            return x["folio"]
    return None


def printerproof_rows(r, brand):
    """Translate results.json into Errors rows. Every row is a candidate the proofer
    confirms on the page; nothing here is auto-'fixed'."""
    E, C = [], []   # errors, coverage

    def cov(check, status, detail):
        C.append({"Check": check, "Status": status, "Coverage / detail": detail})

    # -- assembly
    a = r["assembly"]
    off = r["folios"].get("offset") or 2
    spine = max(0, off - 2)            # offset 3 = one spine panel in the file, offset 4 = dual cover, etc.
    body_pages = a["physical_pages"] - spine
    div4 = body_pages % 4 == 0
    if not div4:
        E.append(row("Book", "Page count not divisible by 4 (after removing spine/extra cover panels)",
                     "multiple of 4", f"{body_pages} ({a['physical_pages']} physical, {spine} panel(s))", prio="P1"))
    cov("Page count divisible by 4", "PASS" if div4 else "FIX",
        f"{a['physical_pages']} physical pages, folio offset {off} => {spine} spine/cover panel(s), {body_pages} imposed pages")
    if a["duplicate_stamps"]:
        E.append(row("Book", "Pages share an InDesign stamp (assembled from more than one source?)",
                     "", str(a["duplicate_stamps"]), prio="P3"))

    # -- folios / page order
    f = r["folios"]
    for iss in f["issues"]:
        if len(iss) >= 2 and isinstance(iss[0], int):
            E.append(row(f"PDF page {iss[0]}", "Page number out of sequence: " + "; ".join(iss[2:]),
                         "", f"prints {iss[1]}", page=iss[1], prio="P1"))
    cov("Page order (folios sequential, even-left / odd-right)",
        "PASS" if not f["issues"] else "FIX",
        f"folios {f['first']}-{f['last']}, {f['ok']} correct, {len(f['issues'])} issues, "
        f"{len(f.get('unnumbered', []))} unnumbered pages")

    # -- TOC
    toc = r["toc"]
    if brand == "FSR":
        toc = {}
    if toc and any(t["entries"] for t in toc.values()):
        for p, t in toc.items():
            for lab, num, actual in t["mismatched"]:
                E.append(row(lab, "TOC page number does not match where the section starts",
                             actual, num, page=folio_of(r, int(p)), prio="P2", sheet=f"TOC (pdf p{p})"))
            resolved = t["correct"] + len(t["mismatched"])
            status = "FIX" if t["mismatched"] else ("PASS" if resolved == t["entries"] else "PASS (partial)")
            cov(f"TOC vs actual section start (pdf p{p})", status,
                f"{t['correct']}/{t['entries']} correct, {len(t['mismatched'])} mismatched, "
                f"{len(t['unresolved'])} unresolved")
    else:
        cov("TOC vs actual section start", "Not applicable" if brand == "FSR" else "DID NOT RUN",
            "FSR TOC layout not parsed by this script — covered by the FSR app's TOC extraction")

    # -- phones (BPR profile header vs 'or call' line)
    ph = r["phones"]
    if ph["expected"] and ph["parsed"]:
        for p, hdr, call in ph["mismatches"]:
            E.append(row(f"PDF page {p}", "Profile header phone differs from 'or call' line", hdr, call,
                         page=folio_of(r, p), prio="P2"))
        cov("Profile phone vs 'or call' line", "FIX" if ph["mismatches"] else
            ("PASS" if ph["parsed"] == ph["expected"] else "PASS (partial)"),
            f"{len(ph['mismatches'])} mismatches across {ph['parsed']} of {ph['expected']} call-lines")
    else:
        cov("Profile phone vs 'or call' line", "Not applicable", "BPR-only layout — FSR phones are covered by the FSR app (header vs CSV)")

    # -- ratings
    ra = r["ratings"]
    if ra["checked"] or ra["unreadable"]:
        for p, hist, avg, printed in ra["mismatches"]:
            E.append(row(f"PDF page {p}", "Printed rating does not match star histogram",
                         f"{avg} from {hist}", printed, page=folio_of(r, p), prio="P2"))
        cov("Rating recomputed from histogram", "FIX" if ra["mismatches"] else "PASS",
            f"{ra['checked']} recomputed, {len(ra['mismatches'])} mismatches, {len(ra['unreadable'])} unread")
    else:
        cov("Rating recomputed from histogram", "Not applicable", "BPR-only layout (FSR profiles carry no histogram)")

    # -- directory / Quick Reference Sheet
    d = r["directory"]
    if d["listing_count"]:
        for phn in d["missing_from_sheet"]:
            E.append(row(phn, "Listed company missing from Quick Reference Sheet", "on sheet", "missing", prio="P2"))
        for phn in d["orphans_on_sheet"]:
            E.append(row(phn, "Quick Reference Sheet entry has no category listing", "in listings", "orphan", prio="P2"))
        for c in d["categories_missing_from_sheet"]:
            E.append(row(c, "Category heading missing from Quick Reference Sheet", "present", "missing", prio="P2"))
        ok = not (d["missing_from_sheet"] or d["orphans_on_sheet"] or d["categories_missing_from_sheet"])
        cov("Quick Reference Sheet vs listings (both directions)", "PASS" if ok else "FIX",
            f"{d['sheet_count']} on sheet vs {d['listing_count']} in listings; sheet pages {d['sheet_pages']}")
    else:
        cov("Quick Reference Sheet vs listings", "Not applicable", "BPR-only")

    # -- tenure
    t = r["tenure"]
    if t["max_claim"] is not None:
        for cat, p, tens in t["order_issues"]:
            E.append(row(cat, "Tenure not in descending order within category", "descending", str(tens),
                         page=folio_of(r, p), prio="P3"))
        # ceiling: carried open wave-wide — report on Coverage, not as per-book findings
        cov("Tenure order within category", "FIX" if t["order_issues"] else "PASS",
            f"{len(t['order_issues'])} order issues")
        cov("Tenure vs edition ceiling", "VERIFY" if t["over_ceiling"] else "PASS",
            f"max claim {t['max_claim']} vs edition {t['edition']}; {len(t['over_ceiling'])} over ceiling "
            f"— carried open wave-wide; closes against the CSV 'Number of Years' column")
        cov("Tenure differs across categories", "PASS",
            f"{len(t['cross_category'])} of {t['multi_category_companies']} multi-category companies differ "
            f"— BY DESIGN (signing date is per category; retired as a flag)")
    else:
        cov("Tenure checks", "Not applicable", "BPR-only ('Nth year as a Best Pick')")

    # -- licences
    li = r["licences"]
    if li["values_read"]:
        for p, label in li["labels_without_value"]:
            E.append(row(f"PDF page {p}", f"'{label}' prints with no licence number", "a number", "none",
                         page=folio_of(r, p), prio="P2"))
        for o in li["format_outliers"]:
            if isinstance(o[2], list):
                E.append(row(o[0], f"Licence class rare in this book ({o[3]}) — check state/class",
                             "", ", ".join(o[2]), page=folio_of(r, o[1]), prio="P3"))
            else:
                E.append(row(o[2], f"Licence length unusual for class {o[0]} — possible dropped digit",
                             f"{o[4]} digits (class majority)", f"{o[3]} digits", page=folio_of(r, o[1]), prio="P3"))
        cov("Licence labels paired with values; format outliers",
            "FIX" if li["labels_without_value"] else ("VERIFY" if li["format_outliers"] else "PASS"),
            f"{li['values_read']} values read, classes {{{', '.join(f'{k}: {len(v)}' for k, v in li['classes'].items())}}}; "
            f"{len(li['labels_without_value'])} labels without value, {len(li['format_outliers'])} outliers")
    else:
        cov("Licence labels paired with values", "Not applicable" if brand == "FSR" else "DID NOT RUN",
            "FSR prints 'License(s):' inline — covered by the FSR app's licence extraction")

    # -- QR
    q = r["qr"]
    if q.get("skipped"):
        cov("QR codes decoded", "DID NOT RUN", "QR pass skipped")
    else:
        for p, v in q["short_by_page"].items():
            E.append(row(f"PDF page {p}", "QR code(s) on page could not be decoded", v["expected"],
                         len(v["decoded"]), page=folio_of(r, int(p)), prio="P1"))
        for dest, pages in q["duplicates"].items():
            E.append(row(dest, "Same QR destination appears on more than one page", "unique", str(pages), prio="P1"))
        cov("QR codes decoded", "FIX" if (q["short_by_page"] or q["duplicates"]) else "PASS",
            f"{q['decoded']}/{q['expected']} decoded, {q['unique']} unique destinations")
    cov("Back-cover QR destination / UTMs (metro + book)", "NEEDS MANUAL CHECK",
        "decoded short link recorded; compare to the Back Cover bitlys sheet (lookup to be added)")

    # -- characters / hygiene
    c = r["characters"]
    byword = {}
    for p, ctx in c["straight_apostrophes"]:
        m = re.search(r"[A-Za-z]+'[A-Za-z]+", ctx)
        wd = m.group(0) if m else ctx.strip()[:30]
        byword.setdefault(wd, set()).add(folio_of(r, p) or f"pdf {p}")
    for wd, pages in byword.items():
        E.append(row(wd, f"Straight apostrophe in '{wd}' — fix the source database field, not the layout",
                     wd.replace("'", "’"), wd, page=", ".join(str(x) for x in sorted(pages, key=str)), prio="P3"))
    for p, ctx in c["nested_quotes"]:
        E.append(row(ctx.strip(), "Doubled quote marks", "", "", page=folio_of(r, p), prio="P3"))
    cov("Straight apostrophes / doubled quotes", "VERIFY" if (c["straight_apostrophes"] or c["nested_quotes"]) else "PASS",
        f"{len(c['straight_apostrophes'])} straight apostrophes in {len(byword)} distinct word(s) {sorted(byword)}, {len(c['nested_quotes'])} doubled quotes "
        f"(symbol-font glyphs: {len(c['private_use_glyphs'])}, icon fonts — informational)")
    h = r["hygiene"]
    for p, s in h["ordinals"]:
        E.append(row(s, "Wrong ordinal suffix", "", s, page=folio_of(r, p), prio="P2"))
    for p, s in h["doubled_words"]:
        E.append(row(s, "Doubled word", "", s, page=folio_of(r, p), prio="P3"))
    for a_, b_, j in h["duplicate_pages"]:
        E.append(row(f"PDF pages {a_}/{b_}", "Two consecutive pages share most of their text — duplicate page?",
                     "", f"{int(j*100)}% overlap", page=folio_of(r, a_), prio="P1"))
    cov("Ordinals, doubled words, duplicate pages",
        "FIX" if (h["ordinals"] or h["duplicate_pages"]) else ("VERIFY" if h["doubled_words"] else "PASS"),
        f"{len(h['ordinals'])} ordinals, {len(h['doubled_words'])} doubled words, {len(h['duplicate_pages'])} duplicate-page flags")

    # -- tracking
    k = r["tracking"]
    for s, p in k["flagged_pages"]:
        E.append(row(f"PDF page {p}", "Text tracking tighter than book baseline — look at the page",
                     f">= {0.96}", s, page=folio_of(r, p), prio="P4"))
    cov("Tracking / kerning (text-layer glyph spacing)", "VERIFY" if k["flagged_pages"] else "PASS",
        f"median {k['median']} over {k['lines']} lines; {len(k['flagged_pages'])} pages below threshold "
        f"(expect front-matter / mixed-font pages to flag; confirm by eye)")

    # -- standing manual items
    cov("Front/back covers present, metro spelled right, edition year/season", "NEEDS MANUAL CHECK",
        f"cover reads '{r.get('cover_metro')}', edition {r.get('edition')} — placed artwork needs eyes; "
        f"text can be matched to Book Cover Titles sheet once that lookup is wired in")
    cov("Placed artwork on front matter (website mockup, guarantee photo)", "NEEDS MANUAL CHECK",
        "raster art is invisible to text extraction — render-and-inspect, or compare across the wave")
    cov("Salesforce Production / Specialty Notes applied", "NEEDS MANUAL CHECK", "human judgment")
    return E, C


def app_rows(app_path, brand):
    """Pull the app's own findings into the same schema."""
    if not app_path:
        return [], {}
    xl = pd.ExcelFile(app_path)
    rows, passthrough = [], {}
    if brand == "BPR" and "Errors" in xl.sheet_names:
        e = pd.read_excel(xl, "Errors")
        for _, x in e.iterrows():
            rows.append({"Sheet": x.get("Sheet", ""), "Row": x.get("Row", ""), "Key": x.get("Key", ""),
                         "Issue": x.get("Issue", ""), "Expected": x.get("Expected", ""), "Found": x.get("Found", ""),
                         "Page": x.get("Page", ""), "Source": f"BPR app / {x.get('Source', '')}", "Priority": "P3"})
        for s in xl.sheet_names:
            if s != "Errors":
                passthrough[s] = pd.read_excel(xl, s)
    elif brand == "FSR":
        s0 = xl.sheet_names[0]
        e = pd.read_excel(xl, s0)
        for _, x in e.iterrows():
            summ = str(x.get("Error Summary", ""))
            if summ in ("", "No Issues", "nan"):
                continue
            for part in summ.split(";"):
                part = part.strip()
                exp = found = ""
                if "Phone" in part:
                    exp, found = x.get("Phone_csv", ""), x.get("Phone_comb", "")
                elif "Order" in part:
                    exp, found = x.get("Order_csv", ""), x.get("Order_comb", "")
                elif "License" in part:
                    exp, found = x.get("Licenses_csv", ""), x.get("Licenses_comb", "")
                rows.append({"Sheet": "Comparison Results", "Row": "", "Key": x.get("__key__", ""),
                             "Issue": part, "Expected": exp, "Found": found, "Page": x.get("Page_comb", ""),
                             "Source": "FSR app", "Priority": "P3"})
        passthrough[s0] = e
    return rows, passthrough


def write(out, errors, coverage, passthrough, app_cov_line, spellcheck=None):
    errors = pd.DataFrame(errors, columns=COLS)
    order = {"P1": 0, "P2": 1, "P3": 2, "P4": 3}
    errors["_o"] = errors["Priority"].map(order).fillna(9)
    errors = errors.sort_values(["_o", "Source", "Page"], kind="mergesort").drop(columns="_o")
    cov = pd.DataFrame([app_cov_line] + coverage) if app_cov_line else pd.DataFrame(coverage)
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        errors.to_excel(w, sheet_name="Errors", index=False)
        cov.to_excel(w, sheet_name="Coverage", index=False)
        pd.DataFrame(spellcheck or [], columns=COLS).to_excel(
            w, sheet_name="Spellcheck", index=False
        )
        for name, df in passthrough.items():
            if name == "Spellcheck":
                continue
            df.to_excel(w, sheet_name=name[:31], index=False)
    wb = load_workbook(out)
    # --- Read me tab, first
    rm = wb.create_sheet("Read me", 0)
    lines = [
        ("How to read this workbook", ""),
        ("", ""),
        ("Errors", "One row per thing to look at. The app's own rows (Source = BPR app / FSR app) and the printer-proof rows (Source = PrinterProof) share the same columns."),
        ("Coverage", "One row per check that ran, with what it covered. PASS means it ran and found nothing. Not applicable means the check does not exist for this brand. DID NOT RUN means it should have and did not."),
        ("Spellcheck", "pyspellchecker English candidates from geometry-aware PDF text blocks. Columns and designed text areas are checked independently; suggestions and surrounding text are shown in the Cobb review format."),
        ("Other tabs", "Copied through from the app unchanged (Profiles, Cat Listings, TOC, Comparison Results)."),
        ("", ""),
        ("Columns on Errors", ""),
        ("Key", "What the row is about: a company, a licence number, a word, or a PDF page."),
        ("Issue", "What was found, in plain words."),
        ("Expected / Found", "What it should be vs what the book prints, where that applies."),
        ("Page", "Printed page number (the folio). 'pdf N' means an unnumbered page, counted from the front of the file."),
        ("Priority", "P1 blocks sign-off. P2 fix before print. P3 confirm against Salesforce or the source sheet. P4 informational, glance at it."),
        ("", ""),
        ("Every row is a candidate", "Confirm it on the page before treating it as a defect. A row with Priority P3 usually closes against the Salesforce export."),
    ]
    for i, (a, b) in enumerate(lines, start=1):
        rm.cell(i, 1, a); rm.cell(i, 2, b)
        rm.cell(i, 1).font = Font(bold=True)
        rm.cell(i, 2).alignment = Alignment(wrap_text=True, vertical="top")
    rm.column_dimensions["A"].width = 26; rm.column_dimensions["B"].width = 110
    hdr = PatternFill("solid", start_color="1F4E79")
    for ws in (wb["Errors"], wb["Coverage"], wb["Spellcheck"]):
        ws.freeze_panes = "A2"
        for c in ws[1]:
            c.fill = hdr; c.font = Font(bold=True, color="FFFFFF")
        for col in range(1, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(col)].width = 18
    ws = wb["Errors"]
    for col, wdt in {3: 44, 4: 58, 5: 22, 6: 22}.items():
        ws.column_dimensions[get_column_letter(col)].width = wdt
    for rr in range(2, ws.max_row + 1):
        pr = ws.cell(rr, 9).value
        if pr in FILL:
            ws.cell(rr, 9).fill = PatternFill("solid", start_color=FILL[pr])
        for cc in (3, 4):
            ws.cell(rr, cc).alignment = Alignment(wrap_text=True, vertical="top")
    ws = wb["Coverage"]
    ws.column_dimensions["A"].width = 58; ws.column_dimensions["B"].width = 20; ws.column_dimensions["C"].width = 110
    for rr in range(2, ws.max_row + 1):
        stv = ws.cell(rr, 2).value
        if stv in FILL:
            ws.cell(rr, 2).fill = PatternFill("solid", start_color=FILL[stv])
        ws.cell(rr, 3).alignment = Alignment(wrap_text=True, vertical="top")
    ws = wb["Spellcheck"]
    for col, wdt in {3: 30, 4: 62, 5: 30, 6: 75, 7: 30, 8: 30}.items():
        ws.column_dimensions[get_column_letter(col)].width = wdt
    for rr in range(2, ws.max_row + 1):
        pr = ws.cell(rr, 9).value
        if pr in FILL:
            ws.cell(rr, 9).fill = PatternFill("solid", start_color=FILL[pr])
        for cc in (4, 5, 6, 7):
            ws.cell(rr, cc).alignment = Alignment(wrap_text=True, vertical="top")
    wb.save(out)


def build(results_path, out_path, brand="BPR", app_workbook=None, app_note=""):
    """Library entry point used by the Streamlit app. Returns out_path."""
    r = json.load(open(results_path))
    E, C = printerproof_rows(r, brand)
    spellcheck = r.get("spellcheck", [])
    p2 = sum(item.get("Priority") == "P2" for item in spellcheck)
    p3 = sum(item.get("Priority") == "P3" for item in spellcheck)
    p4 = sum(item.get("Priority") == "P4" for item in spellcheck)
    C.append({
        "Check": "Spelling (geometry blocks + pyspellchecker + BBB/project allow-lists)",
        "Status": "VERIFY" if spellcheck else "PASS",
        "Coverage / detail": (
            f"{r.get('pages', 0)} pages and {r.get('spellcheck_blocks', 0)} text blocks scanned; "
            f"{len(spellcheck)} distinct words flagged: "
            f"{p2} likely typo(s) (P2), {p3} name/brand candidate(s) (P3), "
            f"{p4} informational/allow-list candidate(s) (P4). See Spellcheck tab."
        ),
    })
    A, P = app_rows(app_workbook, brand)
    app_line = None
    if app_workbook:
        app_line = {"Check": f"{brand} proofing app (company presence/order, phone, licence vs Salesforce CSV)",
                    "Status": "FIX" if A else "PASS",
                    "Coverage / detail": app_note or f"{len(A)} findings from the app"}
    write(out_path, A + E, C, P, app_line, spellcheck=spellcheck)
    return out_path


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("results")
    ap.add_argument("--app", default=None)
    ap.add_argument("--brand", choices=["BPR", "FSR"], required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--app-note", default="")
    a = ap.parse_args()
    r = json.load(open(a.results))
    E, C = printerproof_rows(r, a.brand)
    A, P = app_rows(a.app, a.brand)
    app_line = None
    if a.app:
        app_line = {"Check": f"{a.brand} proofing app (company presence/order, phone, licence vs Salesforce CSV)",
                    "Status": "FIX" if A else "PASS",
                    "Coverage / detail": a.app_note or f"{len(A)} findings from the app"}
    write(a.out, A + E, C, P, app_line)
    print(f"{a.out}: {len(E)} printer-proof rows + {len(A)} app rows; {len(C)} coverage lines")
