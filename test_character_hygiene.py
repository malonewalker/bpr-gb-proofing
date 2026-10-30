import unittest

from printerproof.bpr_textproof import check_characters


class CharacterHygieneTests(unittest.TestCase):
    def test_straight_apostrophe_and_quotes_pass(self):
        result = check_characters({1: 'The company\'s owner said, "Great work."'})

        self.assertEqual(result["smart_apostrophes"], [])
        self.assertEqual(result["smart_quotation_marks"], [])

    def test_smart_apostrophe_is_flagged(self):
        result = check_characters({2: "The company’s owner."})

        self.assertEqual(len(result["smart_apostrophes"]), 1)

    def test_smart_quotation_pair_is_flagged(self):
        result = check_characters({3: "The owner said, “Great work.”"})

        self.assertEqual(len(result["smart_quotation_marks"]), 1)


if __name__ == "__main__":
    unittest.main()
