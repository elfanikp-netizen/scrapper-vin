import unittest
from unittest.mock import Mock, patch

from playwright.sync_api import Error as PWError

from vin_crawler import (
    Blocked,
    close_context,
    google_search,
    is_blocked,
    is_completed_status,
    wait_for_human,
)


class VinCrawlerResumeTests(unittest.TestCase):
    def test_reloads_exact_search_url_after_manual_captcha_clearance(self):
        page = Mock()
        page.url = "https://www.google.com/sorry/index?continue=https://www.google.com/search?q=car"
        page.inner_text.return_value = "Search results"
        page.eval_on_selector_all.return_value = []

        with (
            patch("vin_crawler.accept_consent"),
            patch("vin_crawler.is_blocked", side_effect=[True, False]),
            patch("vin_crawler.wait_for_human") as wait_for_human,
        ):
            google_search(page, "2006 ACURA MDX VIN", headless=False)

        expected_url = page.goto.call_args_list[0].args[0]
        self.assertEqual(page.goto.call_count, 2)
        self.assertEqual(page.goto.call_args_list[1].args[0], expected_url)
        wait_for_human.assert_called_once_with(page, False)

    def test_only_terminal_search_results_are_considered_complete(self):
        self.assertTrue(is_completed_status("found"))
        self.assertTrue(is_completed_status("not_found"))
        self.assertFalse(is_completed_status("error: TimeoutError"))
        self.assertFalse(is_completed_status(""))

    def test_captcha_wait_stops_if_browser_closes(self):
        page = Mock()
        page.is_closed.return_value = True
        with self.assertRaisesRegex(Blocked, "Browser was closed"):
            wait_for_human(page, headless=False)

    def test_cleared_sorry_url_with_search_results_is_not_blocked(self):
        page = Mock()
        page.is_closed.return_value = False
        page.url = "https://www.google.com/sorry/index"
        page.locator.return_value.inner_text.return_value = "Search results for the query"
        page.locator.return_value.count.return_value = 0
        page.locator.side_effect = lambda selector: self._locator_for(page, selector)
        self.assertFalse(is_blocked(page))

    def test_cleanup_ignores_context_already_closed_by_user(self):
        context = Mock()
        context.close.side_effect = PWError("Target page, context or browser has been closed")
        close_context(context)
        context.close.assert_called_once_with()

    @staticmethod
    def _locator_for(page, selector):
        locator = Mock()
        if selector == "body":
            locator.inner_text.return_value = "Search results for the query"
        elif selector == "#search:visible, #rso:visible":
            locator.count.return_value = 1
        else:
            locator.count.return_value = 0
        return locator


if __name__ == "__main__":
    unittest.main()