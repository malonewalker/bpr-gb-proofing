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


def run_printer_proof_ui(pdf_file, ref_file):
    """Run the printer-proof scan, optionally merging the existing BPR checks."""
    if pdf_file is None:
        st.error("Please upload a PDF before running the printer-proof scan.")
        return

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)
        pdf_stem = Path(pdf_file.name).stem
        pdf_path = tmpdir_path / f"{pdf_stem}.pdf"
        save_uploaded_file(pdf_file, pdf_path)

        app_workbook = None
        app_note = ""
        if ref_file is not None:
            ref_path = tmpdir_path / "reference.xlsx"
            save_uploaded_file(ref_file, ref_path)
            app_workbook = tmpdir_path / f"{pdf_stem}_BPR_Combined.xlsx"
            with st.spinner("Running BPR proofing script (vs BBB file)..."):
                try:
                    bpr_pipeline.run_pipeline(
                        pdf_path=str(pdf_path),
                        ref_excel_path=str(ref_path),
                        save_path=str(app_workbook),
                    )
                except Exception as e:
                    st.error(
                        "The BPR proofing script failed; running the printer-proof "
                        "scan on its own."
                    )
                    st.exception(e)
                    app_workbook = None
                    app_note = (
                        "BPR proofing script FAILED on this run — see error above"
                    )
        else:
            st.info(
                "No BBB file uploaded: running the printer-proof scan only. "
                "Upload the BBB file too to get both sets of checks in one workbook."
            )

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

        **Printer-proof scan** (second button) runs the pre-printer checks on the same PDF —
        page order and page count, QR codes, ratings vs star histograms, Quick Reference Sheet,
        licence pairing, apostrophes and text tracking — and adds the results to the same
        **Errors** tab plus a **Coverage** tab. The BBB file is optional for this scan; with it,
        the app's own checks run first and the two are merged.
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
    btn_col1, btn_col2 = st.columns(2)
    with btn_col1:
        run_button = st.button("Run Proofing Script")
    with btn_col2:
        pp_button = st.button("Run Printer-Proof Scan")

    if pp_button:
        run_printer_proof_ui(pdf_file, ref_file)
        return

    if run_button:
        # Basic validation
        if pdf_file is None or ref_file is None:
            st.error("Please upload **both** a PDF and a reference Excel/CSV file before running the pipeline.")
            return

        # Use a temporary directory for all intermediate files
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir_path = Path(tmpdir)

            # Construct paths for the inputs and output
            pdf_path = tmpdir_path / "input.pdf"
            ref_path = tmpdir_path / "reference.xlsx"  # can be CSV too; name is just for convenience

            # Use the PDF's name to build a friendly output filename
            pdf_stem = Path(pdf_file.name).stem
            out_path = tmpdir_path / f"{pdf_stem}_BPR_Combined.xlsx"

            # Save uploads
            save_uploaded_file(pdf_file, pdf_path)
            save_uploaded_file(ref_file, ref_path)

            # Run the pipeline
            with st.spinner("Running BPR GB pipeline..."):
                try:
                    # NOTE:
                    # Update bpr_pipeline.run_pipeline to accept this new signature:
                    #   run_pipeline(pdf_path: str, ref_excel_path: str, save_path: str) -> str
                    final_path_str = bpr_pipeline.run_pipeline(
                        pdf_path=str(pdf_path),
                        ref_excel_path=str(ref_path),
                        save_path=str(out_path),
                    )
                except Exception as e:
                    st.error("The pipeline failed. Check the logs for details.")
                    st.exception(e)
                    return

            # Read the output workbook into bytes for download
            final_path = Path(final_path_str)
            if not final_path.is_file():
                st.error("Pipeline finished but the expected output file was not found.")
                return

            with open(final_path, "rb") as f:
                data = f.read()

            st.success("Pipeline complete! Download the combined workbook below.")

            st.download_button(
                label="Download Combined Workbook",
                data=data,
                file_name=final_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


if __name__ == "__main__":
    main()
