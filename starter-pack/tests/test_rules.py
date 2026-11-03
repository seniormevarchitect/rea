"""Edge cases called out on 2 November 2026: 2/9, 3/9 and 2/12/2026.

Those dates are day/month, matching the Italian review note and the
worked example (a shift on 2 December 2026 guides to 1 March 2027).
"""

from __future__ import annotations

import sys
import unittest
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import build_tracker
import rules

TODAY = date(2026, 11, 3)


def _as_date(value):
    if isinstance(value, datetime):
        return value.date()
    return value


class ClaimByTests(unittest.TestCase):
    def test_2_september_lands_on_the_cap(self):
        self.assertEqual(rules.claim_by_date(date(2026, 9, 2)), date(2026, 11, 30))

    def test_3_september_is_capped_back_to_30_november(self):
        raw = date(2026, 9, 3) + timedelta(days=89)
        self.assertEqual(raw, date(2026, 12, 1))
        self.assertEqual(rules.claim_by_date(date(2026, 9, 3)), date(2026, 11, 30))

    def test_2_december_guides_to_1_march_2027(self):
        self.assertEqual(rules.claim_by_date(date(2026, 12, 2)), date(2027, 3, 1))

    def test_rule_start_day_is_not_capped(self):
        self.assertEqual(rules.claim_by_date(date(2026, 12, 1)), date(2027, 2, 28))

    def test_weekend_guide_date_is_not_moved(self):
        # 7 December 2026 + 89 days is Saturday 6 March 2027.
        self.assertEqual(rules.claim_by_date(date(2026, 12, 7)), date(2027, 3, 6))
        self.assertEqual(date(2027, 3, 6).weekday(), 5)

    def test_who_claims(self):
        self.assertEqual(rules.who_usually_claims("NDIA-managed"), "You")
        self.assertEqual(rules.who_usually_claims("Plan-managed"), "Plan manager")
        self.assertEqual(rules.who_usually_claims("Self-managed"), "Participant")
        self.assertEqual(rules.who_usually_claims("Agency"), "")

    def test_buffer_and_alert_on_review_day(self):
        support = date(2026, 9, 2)
        self.assertEqual(rules.target_date(support, 14), date(2026, 11, 16))
        self.assertEqual(rules.days_left(support, TODAY, buffer_days=14), 13)
        self.assertEqual(rules.alert(support, TODAY, buffer_days=14), "Due soon")
        self.assertEqual(
            rules.alert(support, TODAY, buffer_days=14, submission=date(2026, 11, 1)),
            "Lodged",
        )
        self.assertIsNone(
            rules.days_left(support, TODAY, buffer_days=14, submission=date(2026, 11, 1))
        )

    def test_december_shift_is_on_track_in_early_november(self):
        self.assertEqual(
            rules.alert(date(2026, 12, 2), TODAY, buffer_days=14),
            "On track",
        )


class WorkbookTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.path = Path(__file__).resolve().parents[1] / "claims-tracker-90-day.xlsx"
        build_tracker.build(cls.path)
        cls.wb = load_workbook(cls.path)

    def test_headers_match_the_0153_names_plus_participant_code(self):
        ws = self.wb["Tracker"]
        headers = [ws.cell(4, col).value for col in range(1, 14)]
        self.assertEqual(
            headers,
            [
                "Participant code",
                "Plan management",
                "Who usually claims (auto)",
                "Support date",
                "Claim-by date (guide only)",
                "Your target date (auto)",
                "Days left (auto)",
                "Status",
                "Alert (auto)",
                "Submission date",
                "Reference / confirmation",
                "Follow-up needed",
                "Where the record is kept",
            ],
        )

    def test_example_rows_are_the_three_edge_cases(self):
        ws = self.wb["Tracker"]
        self.assertEqual(_as_date(ws["D5"].value), date(2026, 9, 2))
        self.assertEqual(_as_date(ws["D6"].value), date(2026, 9, 3))
        self.assertEqual(_as_date(ws["D7"].value), date(2026, 12, 2))
        self.assertEqual(ws["B5"].value, "NDIA-managed")
        self.assertEqual(ws["B6"].value, "Plan-managed")
        self.assertEqual(ws["B7"].value, "Self-managed")

    def test_formulas_encode_the_cap_and_the_89_day_offset(self):
        ws = self.wb["Tracker"]
        self.assertEqual(
            ws["E5"].value,
            '=IF(D5="","",IF(D5<Settings!$B$2,MIN(D5+89,Settings!$B$3),D5+89))',
        )
        self.assertEqual(ws["F5"].value, '=IF(E5="","",E5-Settings!$B$1)')
        self.assertIn("TODAY()", ws["G5"].value)
        self.assertIn("Past guide date", ws["I5"].value)
        self.assertEqual(self.wb["Settings"]["B1"].value, 14)
        self.assertEqual(_as_date(self.wb["Settings"]["B2"].value), date(2026, 12, 1))
        self.assertEqual(_as_date(self.wb["Settings"]["B3"].value), date(2026, 11, 30))

    def test_colour_is_only_on_days_left_and_alert(self):
        ws = self.wb["Tracker"]
        ranges = []
        for ref in ws.conditional_formatting._cf_rules:
            ranges.append(str(getattr(ref, "sqref", ref)))
        self.assertTrue(ranges)
        for ref in ranges:
            self.assertRegex(ref, r"^[GI]\d")


class DraftCopyTests(unittest.TestCase):
    root = Path(__file__).resolve().parents[2]

    def test_public_drafts_do_not_link_official_ndis_sites(self):
        drafts = self.root / "drafts"
        banned = ("ndis.gov.au", "ndiscommission.gov.au", "legislation.gov.au")
        for path in drafts.glob("*.html"):
            text = path.read_text(encoding="utf-8")
            for host in banned:
                self.assertNotIn(host, text, path.name)
            self.assertIn("Not published", text, path.name)
            self.assertIn("Independent decision-support resource", text, path.name)

    def test_slugs_do_not_contain_ndis(self):
        names = [p.name for p in (self.root / "drafts").glob("*.html")]
        self.assertIn("free-90-day-claims-tracker.html", names)
        self.assertIn("registration-workbook-small-providers.html", names)
        for name in names:
            self.assertNotIn("ndis", name.lower())


if __name__ == "__main__":
    unittest.main()
