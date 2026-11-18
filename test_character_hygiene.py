import json
from pathlib import Path
import unittest

from printerproof.adapter import printerproof_rows
from printerproof.bpr_textproof import check_characters


class CharacterHygieneTests(unittest.TestCase):
    def test_character_scan_keeps_glyph_and_dash_statistics(self):
        text = 'The company\'s owner said, "Great work." The owner’s reply was “Thanks.” — \ue000'
        result = check_characters({1: text})

        self.assertEqual(set(result), {"private_use_glyphs", "dash_census"})
        self.assertEqual(len(result["private_use_glyphs"]), 1)
        self.assertEqual(result["dash_census"], {"U+2014": 1})

    def test_workbook_ignores_punctuation_findings_in_older_results(self):
        path = Path(__file__).parent / "test-data/printerproof-test-output/run/results/results.json"
        results = json.loads(path.read_text(encoding="utf-8"))
        results["characters"].update({
            "smart_apostrophes": [[2, "The company’s owner"]],
            "smart_quotation_marks": [[3, "The owner said “Thanks.”"]],
        })
        results["hygiene"]["ordinals"] = [[2, "21th"]]

        errors, coverage = printerproof_rows(results, "BPR")

        issues = [row["Issue"] for row in errors]
        checks = [row["Check"] for row in coverage]
        self.assertIn("Wrong ordinal suffix", issues)
        self.assertIn("Ordinals, doubled words, duplicate pages", checks)
        self.assertFalse(any("apostrophe" in issue or "quotation marks" in issue for issue in issues))
        self.assertFalse(any("apostrophe" in check or "quotation marks" in check for check in checks))


if __name__ == "__main__":
    unittest.main()
