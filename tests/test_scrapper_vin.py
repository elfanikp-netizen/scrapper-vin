import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import Mock, patch

from openpyxl import Workbook

from scrapper_vin import (
    BrowserClosedError,
    build_vehicle_url,
    extract_vins,
    is_completed_result,
    page_needs_verification,
    read_vehicles,
    slugify,
    wait_for_manual_verification,
)


class VinCrawlerTests(unittest.TestCase):
    def test_builds_expected_vehicle_url(self):
        self.assertEqual(
            build_vehicle_url(2006, "ACURA", "MDX"),
            "https://vininspect.com/vin/acura/mdx/2006",
        )

    def test_slugifies_multiword_model(self):
        self.assertEqual(slugify("TSX SEDAN"), "tsx-sedan")

    def test_extracts_unique_valid_17_character_vins(self):
        text = "VIN: 1HGCM82633A004352; invalid: 1HGCM82633I004352; duplicate 1HGCM82633A004352"
        self.assertEqual(extract_vins(text), ["1HGCM82633A004352"])

    def test_resume_skips_completed_matches_and_no_match_rows(self):
        self.assertTrue(is_completed_result("1HGCM82633A004352", "success"))
        self.assertTrue(is_completed_result("", "vin_not_found"))
        self.assertFalse(is_completed_result("", "success"))
        self.assertFalse(is_completed_result("", "pending"))

    def test_reads_vehicle_rows_by_header(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "vehicles.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(("Year", "Make", "Model", "VIN"))
            sheet.append((2006, "ACURA", "MDX", None))
            sheet.append((None, "ACURA", "TL", None))
            workbook.save(path)
            workbook.close()
            self.assertEqual(read_vehicles(path, None, 1), [(2006, "ACURA", "MDX")])

    def test_stops_manual_verification_wait_when_browser_closes(self):
        page = Mock()
        page.is_closed.side_effect = [False, True]
        with (
            patch("scrapper_vin.page_needs_verification", return_value=True),
            self.assertRaises(BrowserClosedError),
        ):
            wait_for_manual_verification(page)

    def test_manual_verification_wait_continues_only_after_challenge_clears(self):
        page = Mock()
        page.is_closed.return_value = False
        with (
            patch("scrapper_vin.page_needs_verification", side_effect=[True, False]),
            patch("scrapper_vin.time.monotonic", side_effect=[0, 0]),
        ):
            wait_for_manual_verification(page)
        page.wait_for_timeout.assert_not_called()

    def test_detects_verifying_you_are_human_challenge_text(self):
        page = Mock()
        page.title.return_value = "VinInspect"
        page.locator.return_value.inner_text.return_value = (
            "Verifying you are human. This may take a few seconds."
        )
        self.assertTrue(page_needs_verification(page))


if __name__ == "__main__":
    unittest.main()