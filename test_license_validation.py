import unittest

import pandas as pd

from sl_newvalidate import run_checks


class LicenseValidationTests(unittest.TestCase):
    def test_trade_license_not_required_badge_with_pdf_spacing(self):
        primary = pd.DataFrame(
            [
                {
                    "Phone": "770-284-8356",
                    "Rating (out of 5)": "4.9",
                    "Five-Star (count)": "271",
                    "Trade License Numbers": "",
                    "Verified Block": (
                        "\u2004Trade License(s) Not Required\n"
                        "\u2004Verified General Liability Insurance\u2003\n"
                        "\u2004Verified Workers' Comp \u2003\u2003\n"
                        "\u2004Best Pick Guaranteed\u2003"
                    ),
                }
            ]
        )
        bbb = pd.DataFrame(
            [
                {
                    "Book Phone Number": "770-284-8356",
                    "Licenses": "Not Required",
                    "WC Status": "Active",
                }
            ]
        )

        checked = run_checks(primary, bbb)

        self.assertNotIn(
            "expected 'Trade License(s) Not Required' in Verified Block",
            checked.loc[0, "Notes_Compare"],
        )

    def test_trade_license_not_required_badge_glued_to_rating_text(self):
        primary = pd.DataFrame(
            [
                {
                    "Phone": "678-919-3915",
                    "Rating (out of 5)": "4.9",
                    "Five-Star (count)": "204",
                    "Trade License Numbers": "",
                    "Verified Block": (
                        "Rating: 4.9 out of 5 Trade License(s) Not Required "
                        "Verified General Liability Insurance "
                        "Verified Workers' Comp Best Pick Guaranteed"
                    ),
                }
            ]
        )
        bbb = pd.DataFrame(
            [
                {
                    "Book Phone Number": "678-919-3915",
                    "Licenses": "Not Required",
                    "WC Status": "Active",
                }
            ]
        )

        checked = run_checks(primary, bbb)

        self.assertEqual("", checked.loc[0, "Notes_Compare"])


if __name__ == "__main__":
    unittest.main()
