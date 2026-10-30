import unittest

from printerproof.bpr_textproof import check_ratings


def rating_text(printed):
    # Weighted average: (5*5 + 4*4 + 3*3 + 2*2 + 1*1) / 15 = 3.666...
    return {1: f"5\n4\n3\n2\n1\nRating: {printed} out of 5\n"}


class RatingRoundingTests(unittest.TestCase):
    def test_conventional_rounded_rating_passes(self):
        result = check_ratings(rating_text("3.7"))

        self.assertEqual(result["checked"], 1)
        self.assertEqual(result["mismatches"], [])

    def test_truncated_rating_passes(self):
        result = check_ratings(rating_text("3.6"))

        self.assertEqual(result["checked"], 1)
        self.assertEqual(result["mismatches"], [])

    def test_value_matching_neither_is_flagged_with_both_options(self):
        result = check_ratings(rating_text("3.5"))

        self.assertEqual(len(result["mismatches"]), 1)
        self.assertIn("3.7 rounded or 3.6 truncated", result["mismatches"][0][2])


if __name__ == "__main__":
    unittest.main()
