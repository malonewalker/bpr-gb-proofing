import unittest

import pandas as pd

from printerproof.book_reference import (
    BOOK_NAME_COLUMN,
    QR_COLUMN,
    SUBHEADER_COLUMN,
    book_lookup_from_dataframe,
    check_book_reference,
    match_book_reference,
)


class BookReferenceTests(unittest.TestCase):
    def test_extracts_text_after_first_hyphen_from_bbb_book(self):
        frame = pd.DataFrame({"Book": ["Fall 2026 - Cobb / Marietta"]})

        self.assertEqual(book_lookup_from_dataframe(frame), "Cobb / Marietta")

    def test_matches_reference_case_insensitively(self):
        frame = pd.DataFrame([{
            BOOK_NAME_COLUMN: "Cobb / Marietta",
            SUBHEADER_COLUMN: "",
            QR_COLUMN: "https://bit.ly/example",
        }])

        result = match_book_reference(frame, "cobb / marietta")
        self.assertEqual(result["book_name"], "Cobb / Marietta")

    def test_rejects_ambiguous_reference_variants(self):
        frame = pd.DataFrame([
            {BOOK_NAME_COLUMN: "Northwest Suburbs", SUBHEADER_COLUMN: "Area One", QR_COLUMN: "https://bit.ly/one"},
            {BOOK_NAME_COLUMN: "Northwest Suburbs", SUBHEADER_COLUMN: "Area Two", QR_COLUMN: "https://bit.ly/two"},
        ])

        with self.assertRaisesRegex(ValueError, "ambiguous"):
            match_book_reference(frame, "Northwest Suburbs")

    def test_checks_spine_front_back_and_qr(self):
        reference = {
            "lookup": "Cobb / Marietta",
            "source_url": "https://example.test/reference.xlsx",
            "book_name": "Cobb / Marietta",
            "book_subheader": "",
            "back_cover_qr_bitly": "https://bit.ly/example",
        }
        layout = {
            1: "BEST PICK REPORTS | Cobb / Marietta",
            2: "Cobb / Marietta\nQuality Home Services Guide",
            3: "Back matter\nCobb / Marietta",
        }
        qr = {"destinations": {"https://bit.ly/example": [3]}}

        result = check_book_reference(layout, qr, reference)
        statuses = {item["name"]: item["status"] for item in result["checks"]}
        self.assertEqual(statuses["Spine book name"], "PASS")
        self.assertEqual(statuses["Front cover book name"], "PASS")
        self.assertEqual(statuses["Front cover book sub-header"], "Not applicable")
        self.assertEqual(statuses["Back cover book name"], "PASS")
        self.assertEqual(statuses["Back cover QR Bitly"], "PASS")


if __name__ == "__main__":
    unittest.main()
