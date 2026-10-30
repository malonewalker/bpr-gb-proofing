import unittest

from printerproof.spellcheck import bbox_text_blocks, _dictionary, check_pages, load_allowlist


class SpellcheckTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.dictionary = _dictionary()

    def check(self, pages, allowlist=None, layout_pages=None):
        return {
            item["Key"]: item
            for item in check_pages(pages, self.dictionary, allowlist, layout_pages)
        }

    def test_loads_project_allowlist_case_insensitively(self):
        self.assertIn("abbasov", load_allowlist())
        self.assertIn("tankless", load_allowlist())
        self.assertEqual(list(self.check({1: "Abbasov installed a tankless heater."})), [])

    def test_lowercase_typo_is_p2_with_suggestions(self):
        found = self.check({40: "Members receive a five-year warrenty on all service repairs."})
        self.assertEqual(found["warrenty"]["Priority"], "P2")
        self.assertIn("warranty", found["warrenty"]["Expected"])
        self.assertEqual(found["warrenty"]["Page"], "pdf 40")

    def test_groups_words_across_pages(self):
        found = self.check({1: "Calls and inquirie welcome.", 2: "More inquirie here."})
        self.assertEqual(found["inquirie"]["Page"], "pdf 1, pdf 2")
        self.assertIn("(+1 more)", found["inquirie"]["Found"])

    def test_name_spelled_two_ways_is_p2(self):
        found = self.check({
            1: "Josh Sulfridge | General Manager",
            2: "Josh Sulfridge leads the team.",
            3: "Josh Sulfrigde | General Manager",
        })
        self.assertEqual(found["Sulfrigde"]["Priority"], "P2")
        self.assertEqual(found["Sulfrigde"]["Expected"], "Sulfridge")

    def test_unknown_name_is_hidden(self):
        found = self.check({1: "Founded by Quintavious Brell in 1990."})
        self.assertNotIn("Quintavious", found)

    def test_consistently_used_unknown_name_is_hidden(self):
        found = self.check({
            1: "Quintavious founded the company.",
            2: "Quintavious leads the team.",
            3: "Call Quintavious for service.",
        })
        self.assertNotIn("Quintavious", found)

    def test_reference_words_are_known(self):
        found = self.check({1: "Founded by Quintavious Brell in 1990."}, allowlist={"Quintavious Brell"})
        self.assertEqual(found, {})

    def test_lost_line_break_hyphen_is_not_flagged(self):
        # pdftotext reading order drops the hyphen in 'ten-\nyear' -> 'tenyear'
        raw = {1: "a tenyear warranty from a familyand veteran-owned company"}
        layout = {1: "a ten-\nyear warranty from a family-\nand veteran-owned company"}
        self.assertEqual(self.check(raw, layout_pages=layout), {})

    def test_trade_compound_with_lost_hyphen_is_not_flagged(self):
        raw = {1: "They pump a mortarbased material into the voids."}
        layout = {1: "They pump a mortar-based material into the voids."}
        self.assertEqual(self.check(raw, layout_pages=layout), {})

    def test_joined_dictionary_words_without_print_evidence_are_flagged(self):
        raw = {1: "They pump a mortarbased material into the voids."}
        found = self.check(raw, layout_pages={1: "They pump a mortarbased material."})
        self.assertIn("mortarbased", found)

    def test_line_mode_checks_layout_lines_independently(self):
        pages = {1: "First line is correct.\nSecond line has warrenty."}
        findings = check_pages(pages, self.dictionary, line_mode=True)
        found = {item["Key"]: item for item in findings}

        self.assertIn("warrenty", found)
        self.assertNotIn("correctSecond", found)

    def test_geometry_blocks_rejoin_wraps_without_crossing_columns(self):
        blocks = {
            129: [
                "The rigorous nature of our qualifica-\ntion process.",
                "If no agreeable solu-\ntion can be achieved.",
            ]
        }
        findings = check_pages(blocks, self.dictionary)

        self.assertEqual(findings, [])

    def test_parses_bbox_layout_into_separate_blocks(self):
        xml = """<html xmlns="http://www.w3.org/1999/xhtml"><body><doc>
        <page><flow><block><line><word>Left</word><word>column</word></line></block>
        <block><line><word>Right</word><word>column</word></line></block></flow></page>
        </doc></body></html>"""

        self.assertEqual(bbox_text_blocks(xml), {1: ["Left column", "Right column"]})

    def test_bbox_parser_removes_illegal_pdf_control_glyphs(self):
        xml = """<html><body><page><flow><block><line>
        <word>Conditioning\x08</word></line></block></flow></page></body></html>"""

        self.assertEqual(bbox_text_blocks(xml), {1: ["Conditioning"]})

    def test_discretionary_hyphen_and_space_are_rejoined(self):
        text = "We treat every cus- tomer with honesty and offer high- quality work."
        found = self.check({111: text})

        self.assertNotIn("cus", found)
        self.assertNotIn("tomer", found)
        self.assertEqual(found, {})

    def test_contractions_possessives_and_urls_are_not_flagged(self):
        text = "It doesn’t matter; the companies’ trucks. Visit www.bestpickreports.com or info@acmeco.com."
        self.assertEqual(self.check({1: text}), {})


if __name__ == "__main__":
    unittest.main()
