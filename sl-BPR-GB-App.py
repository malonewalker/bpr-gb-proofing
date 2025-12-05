#!/usr/bin/env python3
"""
Streamlit UI for BPR GB Combined Pipeline

- Upload:
    1) GB PDF
    2) Expected-order Excel/CSV for BPRproofing
    3) BBB Excel/CSV for newvalidate
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
    st.set_page_config(page_title="BPR GB Proofing — Full Pipeline", layout="wide")
    st.title("BPR GB Proofing — Full Pipeline")

    st.markdown(
        """
        This app runs the full BPR GB pipeline:

        1. **BPRproofing** (TOC / Listings split using the expected-order file)  
        2. **proofing** (Profiles extraction from the PDF)  
        3. **newvalidate** (Profiles vs BBB validation)  
        4. Combines everything into **one Excel workbook** with a unified **Errors** tab.
        """
    )

    st.header("Step 1 — Upload Inputs")

    col1, col2, col3 = st.columns(3)

    with col1:
        pdf_file = st.file_uploader("Book PDF", type=["pdf"], key="pdf_upload")

    with col2:
        bpr_ref_file = st.file_uploader(
            "Expected-order (BPRproofing) Excel/CSV",
            type=["xlsx", "xls", "csv", "txt"],
            key="bpr_ref_upload",
        )

    with col3:
        bbb_file = st.file_uploader(
            "BBB reference Excel/CSV",
            type=["xlsx", "xls", "csv", "txt"],
            key="bbb_upload",
        )

    st.markdown("---")
    run_button = st.button("Run Full Pipeline")

    if run_button:
        # Basic validation
        if pdf_file is None or bpr_ref_file is None or bbb_file is None:
            st.error("Please upload **all three** files before running the pipeline.")
            return

        # Use a temporary directory for all intermediate files
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir_path = Path(tmpdir)

            # Construct paths for the inputs and output
            pdf_path = tmpdir_path / "input.pdf"
            bpr_path = tmpdir_path / "bpr_expected_order.xlsx"
            bbb_path = tmpdir_path / "bbb_reference.xlsx"

            # Use the PDF's name to build a friendly output filename
            pdf_stem = Path(pdf_file.name).stem
            out_path = tmpdir_path / f"{pdf_stem}_BPR_Combined.xlsx"

            # Save uploads
            save_uploaded_file(pdf_file, pdf_path)
            save_uploaded_file(bpr_ref_file, bpr_path)
            save_uploaded_file(bbb_file, bbb_path)

            # Run the pipeline
            with st.spinner("Running BPR GB pipeline... this may take a bit."):
                try:
                    final_path_str = bpr_pipeline.run_pipeline(
                        pdf_path=str(pdf_path),
                        bpr_csv_path=str(bpr_path),
                        bbb_path=str(bbb_path),
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
