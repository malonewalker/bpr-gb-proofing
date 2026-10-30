import unittest

from printerproof.bpr_textproof import check_licences


def geometry(value):
    return [{
        "page": 7,
        "w": 600.0,
        "h": 800.0,
        "lines": [
            [(50.0, 100.0, 180.0, 112.0, "Verified Trade License(s)")],
            [(50.0, 115.0, 180.0, 127.0, value)],
        ],
    }]


class LicenceCharacterTests(unittest.TestCase):
    def test_question_plus_and_caret_are_flagged(self):
        result = check_licences(geometry("ABC+123?^"))

        self.assertEqual(result["invalid_characters"], [[7, "ABC+123?^", "+ ? ^"]])

    def test_normal_license_has_no_character_flag(self):
        result = check_licences(geometry("ABC-123456"))

        self.assertEqual(result["invalid_characters"], [])


if __name__ == "__main__":
    unittest.main()
