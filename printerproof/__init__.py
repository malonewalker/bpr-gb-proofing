"""
printerproof — the pre-printer scan, packaged for the BPR Guidebook Proofing app.

    from printerproof import run_printer_proof
    out = run_printer_proof(pdf_path, out_dir, app_workbook=None, brand="BPR")

Runs bpr_textproof on the PDF (page order, page count, QR codes, ratings vs
histograms, Quick Reference Sheet, licence pairing, apostrophes, tracking) and
writes one workbook in the app's Errors layout plus a Coverage tab. If the
app's own combined workbook is passed in, its Errors rows are merged in so the
proofer reads one file.

Nothing here needs network access or an AI service. Dependencies beyond the
app's existing ones: opencv-python-headless + numpy (QR decoding) and the
poppler command-line tools (pdftotext / pdftoppm / pdfinfo / pdffonts), which
Streamlit Cloud installs from packages.txt.
"""
import os
from pathlib import Path

from . import bpr_textproof
from . import adapter
from . import spellcheck


def run_printer_proof(pdf_path, out_dir, app_workbook=None, brand="BPR", app_note="", skip_qr=False,
                      ref_path=None):
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    results_dir = out_dir / "results"
    bpr_textproof.run(str(pdf_path), str(results_dir), skip_qr=skip_qr,
                      spell_words=spellcheck.reference_words(ref_path))
    results_json = results_dir / "results.json"
    stem = Path(pdf_path).stem
    out_xlsx = out_dir / f"{stem}_proof_merged.xlsx"
    adapter.build(str(results_json), str(out_xlsx), brand=brand,
                  app_workbook=str(app_workbook) if app_workbook else None, app_note=app_note)
    return str(out_xlsx)
