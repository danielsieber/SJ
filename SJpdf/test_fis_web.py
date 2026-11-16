import unittest
from datetime import date

from fis_web import calendar_events, current_season


class FisWebTests(unittest.TestCase):
    def test_season_switches_in_july(self):
        self.assertEqual(current_season(date(2026, 6, 30)), 2026)
        self.assertEqual(current_season(date(2026, 7, 1)), 2027)
        self.assertEqual(current_season(date(2027, 2, 14)), 2027)

    def test_calendar_place_falls_back_to_other_bold_clip(self):
        html = """
        <div id="calendardata">
          <div class="table-row" id="63280">
            <div class="clip bold clip-xs">GER</div>
            <div class="clip bold">Klingenthal</div>
          </div>
        </div>
        """

        self.assertEqual(calendar_events(html), [("63280", "Klingenthal")])

    def test_calendar_without_event_list_reports_changed_page(self):
        with self.assertRaisesRegex(ValueError, "calendardata"):
            calendar_events("<html></html>")


if __name__ == "__main__":
    unittest.main()
