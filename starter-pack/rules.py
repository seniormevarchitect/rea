"""Guide-date rules for the 90-day claims tracker.

Rebuilt from the 2 November 2026 funnel notes. The original
claims-tracker-90-day.xlsx (01:53) was not in this workspace.

The guide date is the support date plus 89 days, so the support day
counts as day 1. Supports delivered before 1 December 2026 are capped
at 30 November 2026. Weekend and public-holiday movement under the
Acts Interpretation Act is not applied.
"""

from __future__ import annotations

from datetime import date, timedelta

RULE_START = date(2026, 12, 1)
OLDER_SUPPORT_CAP = date(2026, 11, 30)
GUIDE_OFFSET_DAYS = 89

PLAN_MANAGEMENT = ("NDIA-managed", "Plan-managed", "Self-managed")

WHO_CLAIMS = {
    "NDIA-managed": "You",
    "Plan-managed": "Plan manager",
    "Self-managed": "Participant",
}

# Rebuild default. The 01:53 workbook's buffer was not recovered.
DEFAULT_BUFFER_DAYS = 14


def who_usually_claims(plan_management: str) -> str:
    return WHO_CLAIMS.get(plan_management.strip(), "")


def claim_by_date(support: date) -> date:
    """Guide-only claim-by date. Not a personalised legal deadline."""
    raw = support + timedelta(days=GUIDE_OFFSET_DAYS)
    if support < RULE_START:
        return min(raw, OLDER_SUPPORT_CAP)
    return raw


def target_date(support: date, buffer_days: int = DEFAULT_BUFFER_DAYS) -> date:
    return claim_by_date(support) - timedelta(days=buffer_days)


def days_left(
    support: date,
    today: date,
    *,
    buffer_days: int = DEFAULT_BUFFER_DAYS,
    submission: date | None = None,
) -> int | None:
    """Days until the target date. Blank once a submission date is entered."""
    if submission is not None:
        return None
    return (target_date(support, buffer_days) - today).days


def alert(
    support: date | None,
    today: date,
    *,
    buffer_days: int = DEFAULT_BUFFER_DAYS,
    submission: date | None = None,
) -> str:
    if support is None:
        return ""
    if submission is not None:
        return "Lodged"
    remaining = days_left(support, today, buffer_days=buffer_days, submission=None)
    if remaining is None:
        return ""
    if remaining < 0:
        return "Past guide date"
    if remaining <= buffer_days:
        return "Due soon"
    return "On track"
