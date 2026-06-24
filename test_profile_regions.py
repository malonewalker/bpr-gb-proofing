import sys
import types
import unittest


try:
    from sl_proofing import create_regions_from_profile_starts, find_profile_start_positions
except ModuleNotFoundError as exc:
    if exc.name != "pdfplumber":
        raise
    sys.modules["pdfplumber"] = types.ModuleType("pdfplumber")
    from sl_proofing import create_regions_from_profile_starts, find_profile_start_positions


def words_for_line(text, top, x0=40):
    words = []
    x = x0
    for token in text.split():
        width = max(len(token) * 5, 8)
        words.append({"text": token, "top": top, "bottom": top + 8, "x0": x, "x1": x + width})
        x += width + 4
    return words


class ProfileRegionTests(unittest.TestCase):
    def test_profile_starts_ignore_qr_call_lines_and_keep_badges_with_prior_company(self):
        words = []
        words.extend(words_for_line("Atlanta Basement Systems 770-284-8356", 40))
        words.extend(words_for_line("Services Offered: foundation repair", 70))
        words.extend(words_for_line("Scan QR Code to visit us online", 250))
        words.extend(words_for_line("or call 770-284-8356 for service", 265))
        words.extend(words_for_line("Rating: 4.9 out of 5", 300))
        words.extend(words_for_line("Trade License(s) Not Required", 315))
        words.extend(words_for_line("Verified General Liability Insurance", 330))
        words.extend(words_for_line("Verified Workers' Comp", 345))
        words.extend(words_for_line("Best Pick Guaranteed", 360))
        words.extend(words_for_line("'58 Foundations & Waterproofing 678-919-3915", 390))
        words.extend(words_for_line("Services Offered: basement waterproofing", 420))

        starts = find_profile_start_positions(words)
        regions = create_regions_from_profile_starts(starts, page_height=700)

        self.assertEqual([40, 390], starts)
        self.assertEqual(2, len(regions))
        self.assertLessEqual(regions[0][0], 40)
        self.assertGreater(regions[0][1], 360)
        self.assertLess(regions[0][1], 390)
        self.assertLessEqual(regions[1][0], 390)


if __name__ == "__main__":
    unittest.main()
