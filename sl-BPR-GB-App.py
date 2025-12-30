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


def save_uploaded_file(uploaded_file, dest_path: Path):
    """
    Save a Streamlit UploadedFile to dest_path.
    """
    with open(dest_path, "wb") as f:
        f.write(uploaded_file.read())


def main():
    st.set_page_config(page_title="BPR Book Proofing", layout="wide")
    st.title("BPR Book Proofing")

    st.markdown(
        """
        This app runs the full BPR GB pipeline using **one PDF** and **one Excel/CSV**:

        1. Extracts & processes the **Book PDF**  
        2. Uses the **single reference Excel/CSV** for:
           - Expected order / Listings checks  
           - BBB / validation logic  
        3. Combines everything into **one Excel workbook** with a unified **Errors** tab.
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
    run_button = st.button("Run Full Pipeline")

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
