import sys
import types
import unittest


try:
    from sl_bprproofing import parse_pairs_split_on_numbers
except ModuleNotFoundError as exc:
    if exc.name != "PyPDF2":
        raise
    pypdf2_stub = types.ModuleType("PyPDF2")
    pypdf2_stub.PdfReader = object
    sys.modules["PyPDF2"] = pypdf2_stub
    from sl_bprproofing import parse_pairs_split_on_numbers


FAIRFAX_RAW_TOC = """FREE  Local Reference for Homeowners!Fairfax County
Air Conditioning & Heating  ...... 5
Animal Removal  ............ 136
Appliance Repair  .............. 9
Basement Remodeling  ....... 11
Bathroom & Kitchen
Remodeling  ................. 80
Bathtub & Shower
Conversions  ................ 14
Cabinet Refacing  ............ 17
Carpet, Upholstery & Rug
Cleaning  .................... 19
Chimneys & Fireplaces  ....... 23
Concrete Leveling & Raising  ...25
Crawl Space
Encapsulation & Repair  ....... 28
Deck Building &
Maintenance  ................ 30
Door & Window
Replacement  ............... 138
Drain/Sewer Cleaning
& Repair  .................... 34
Drainage Systems  ........... 37
Dryer Vent Cleaning  .......... 39
Duct Cleaning  ............... 41
Electricians  ................. 44
Fences  ..................... 48
Fire Damage Restoration  ....129
Fireplaces & Chimneys  ....... 23
Flooring  .................... 50Foundation Repair  ........... 53
Garage Doors  ............... 57
Gutter & Gutter Guard
Installation  .................. 61
Gutter Cleaning  .............. 65
Handyman Services . . . . . . . . . . 68
Heating & Air Conditioning  ...... 5
Home Remodeling &
Additions  ................... 71
House Cleaning Services  ..... 75
Insulation  ................... 77
Kitchen & Bathroom
Remodeling  ................. 80
Landscaping  ................ 84
Lawn Maintenance  .......... 86
Lawn Treatment  ............. 88
Mold Removal  .............. 129
Movers  ..................... 90
Outdoor Kitchens & Living
Spaces  ..................... 93
Painters  .................... 96
Pest & Termite Control  ...... 100
Plumbers  .................. 103
Pressure Washing &
Window Cleaning  ........... 142
Roof Cleaning  .............. 107
Roofers  ................... 109
Sewer/Drain Cleaning
& Repair  .................... 34Shower & Bathtub
Conversions  ................ 14
Siding  ..................... 113
Solar Panel Installation  ...... 117
Sprinkler Systems  .......... 119
Termite & Pest Control  ...... 100
Tile Installation  ............. 121
Tree Services  .............. 124
Upholstery, Rug & Carpet
Cleaning  .................... 19
Water Treatment &
Filtration  ................... 127
Water, Mold, Fire & Storm
Damage Restoration  ........ 129
Waterproofing  .............. 132
Wildlife Removal  ............ 136
Window & Door
Replacement  ............... 138
Window Cleaning &
Pressure Washing  .......... 142
Additional Information  ...... 145
Homeowner Protection
Tips ....................... 147
Seasonal Maintenance
Checklist  .................. 148
FAQs  ...................... 150
Quick Reference Sheet  ...... 151
Best of all,
IT'S FREE!PRSRT STD
US POSTAGE PAID
Best Pick Reports, LLCBest Pick Reports
7700 Irvine Center Dr, Suite 430
Irvine, CA 92618
"""


class TocParserTests(unittest.TestCase):
    def test_fairfax_raw_toc_wraps_glued_entries_and_footer(self):
        pairs = parse_pairs_split_on_numbers(FAIRFAX_RAW_TOC)
        pair_set = set(pairs)

        self.assertEqual(64, len(pairs))
        self.assertIn(("Bathroom & Kitchen Remodeling", 80), pair_set)
        self.assertIn(("Drain/Sewer Cleaning & Repair", 34), pair_set)
        self.assertIn(("Shower & Bathtub Conversions", 14), pair_set)
        self.assertIn(("Foundation Repair", 53), pair_set)
        self.assertIn(("Window Cleaning & Pressure Washing", 142), pair_set)
        self.assertNotIn(("PRSRT STD 7700 Irvine Center Dr, Suite", 430), pair_set)


if __name__ == "__main__":
    unittest.main()
