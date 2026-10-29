#!/usr/bin/env python3
"""
Streamlit UI for BPR GB Combined Pipeline

- Upload:
    1) GB PDF
    2) Single reference Excel/CSV used by all scripts
- Runs the full pipeline in bpr_pipeline.run_pipeline(...)
- Offers the combined Excel workbook as a download
"""

import tempfile
from pathlib import Path

import streamlit as st

import bpr_pipeline  # <-- this is the file with run_pipeline()
from printerproof import run_printer_proof


def save_uploaded_file(uploaded_file, dest_path: Path):
    """
    Save a Streamlit UploadedFile to dest_path.
    """
    with open(dest_path, "wb") as f:
        f.write(uploaded_file.read())


def run_combined_proofing_ui(pdf_file, ref_file):
    """Run both proofing scripts and return their combined workbook."""
    if pdf_file is None or ref_file is None:
        st.error(
            "Please upload both the Book PDF and BBB file before running proofing."
        )
        return

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)
        pdf_stem = Path(pdf_file.name).stem
        pdf_path = tmpdir_path / f"{pdf_stem}.pdf"
        save_uploaded_file(pdf_file, pdf_path)

        ref_suffix = Path(ref_file.name).suffix.lower()
        ref_path = tmpdir_path / f"reference{ref_suffix}"
        save_uploaded_file(ref_file, ref_path)

        app_workbook = tmpdir_path / f"{pdf_stem}_BPR_Combined.xlsx"
        app_note = ""
        with st.spinner("Running BPR proofing script (vs BBB file)..."):
            try:
                bpr_pipeline.run_pipeline(
                    pdf_path=str(pdf_path),
                    ref_excel_path=str(ref_path),
                    save_path=str(app_workbook),
                )
            except Exception as e:
                st.error(
                    "The BPR proofing script failed; the printer-proof scan will "
                    "still run and record that failure in its workbook."
                )
                st.exception(e)
                app_workbook = None
                app_note = "BPR proofing script FAILED on this run — see error above"

        with st.spinner(
            "Running printer-proof scan (page order, QR codes, ratings, ...). "
            "About 1 minute per 40 pages."
        ):
            try:
                out_path = run_printer_proof(
                    pdf_path=str(pdf_path),
                    out_dir=str(tmpdir_path / "printerproof"),
                    app_workbook=str(app_workbook) if app_workbook else None,
                    brand="BPR",
                    app_note=app_note,
                )
            except Exception as e:
                st.error("The printer-proof scan failed.")
                st.exception(e)
                return

        with open(out_path, "rb") as f:
            data = f.read()

    st.success(
        "Printer-proof scan complete. Open the Read me tab first, then Errors and Coverage."
    )
    st.download_button(
        label="Download Proof Workbook",
        data=data,
        file_name=Path(out_path).name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def main():
    st.set_page_config(page_title="BPR Guidebook Proofing", layout="wide")
    st.title("BPR Guidebook Proofing")

    st.markdown(
        """
        This app runs the full BPR guidebook script using **one PDF** and **one Excel/CSV**:

        1. Extracts & processes the **Book PDF**  
        2. Uses the **single reference Excel/CSV** for:
           - Expected order / Listings checks  
           - BBB / validation logic  
        3. Combines everything into **one Excel workbook** with a unified **Errors** tab.

        The same run also performs the **printer-proof scan** on the PDF —
        page order and page count, QR codes, ratings vs star histograms, Quick Reference Sheet,
        licence pairing, apostrophes and text tracking — and adds the results to the same
        **Errors** tab plus a **Coverage** tab. Both uploaded files are required before the
        combined proofing run can start.
        """
    )

    st.header("Step 1 — Upload Inputs")

    col1, col2 = st.columns(2)

    with col1:
        pdf_file = st.file_uploader("Book PDF", type=["pdf"], key="pdf_upload")

    with col2:
        ref_file = st.file_uploader(
            "BBB File",
            type=["xlsx", "xls", "csv", "txt"],
            key="ref_upload",
        )

    st.markdown("---")
    inputs_ready = pdf_file is not None and ref_file is not None
    run_button = st.button(
        "Run Both Proofing Scripts",
        disabled=not inputs_ready,
        help=(
            "Upload both the Book PDF and BBB file to enable proofing."
            if not inputs_ready
            else "Run BPR validation and the printer-proof scan."
        ),
    )

    if not inputs_ready:
        st.caption("Upload both required files to enable the proofing run.")

    if run_button:
        run_combined_proofing_ui(pdf_file, ref_file)


if __name__ == "__main__":
    main()
