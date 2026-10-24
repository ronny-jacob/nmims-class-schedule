"""Routing tests for the timetable mail bot (check_mail.py).

Run from repo root:
    python tests/test_mail_routing.py

These cover the pure helpers that decide *which week* an incoming
attachment belongs to. That decision is what makes a rescheduled
timetable get applied to the right week instead of being ignored (or,
worse, overwriting the wrong week).
"""
from __future__ import annotations

import os
import sys
from datetime import date

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO_ROOT)

import check_mail as cm  # noqa: E402


def d(y, m, dd):
    return date(y, m, dd)


CASES = [
    # (name, attachment filename, subject, body, expected week)
    (
        "plain filename",
        "19.10.2026 to 25.10.2026.xlsx",
        "Time table",
        "",
        d(2026, 10, 19),
    ),
    (
        "filename without spaces around 'to'",
        "19.10.2026 to25.10.2026.xlsx",
        "Time table",
        "",
        d(2026, 10, 19),
    ),
    (
        "sender typo in the month (12.010.2026)",
        "12.010.2026 to 18.10.2026.xlsx",
        "Time table",
        "",
        d(2026, 10, 12),
    ),
    (
        "rescheduled file with a suffix keeps the week",
        "19.10.2026 to 25.10.2026 (Revised).xlsx",
        "Revised time table",
        "",
        d(2026, 10, 19),
    ),
    (
        "vague filename, dates only in the subject",
        "Revised Timetable.xlsx",
        "Rescheduled time table for 19.10.2026 to 25.10.2026",
        "",
        d(2026, 10, 19),
    ),
    (
        "vague filename, date only in the subject",
        "Revised.xlsx",
        "Rescheduled time table w.e.f 20.10.2026",
        "",
        d(2026, 10, 19),
    ),
    (
        "vague filename, dates only in the body",
        "Timetable Revised.xlsx",
        "Rescheduled timetable",
        "The revised time table for 19.10.2026 to 25.10.2026 is attached.",
        d(2026, 10, 19),
    ),
    (
        "hyphen separated dates",
        "19-10-2026 to 25-10-2026.xlsx",
        "Time table",
        "",
        d(2026, 10, 19),
    ),
    (
        "no date anywhere resolves to None",
        "Revised Timetable.xlsx",
        "Rescheduled timetable",
        "please find attached",
        None,
    ),
]


def main() -> int:
    failures = []

    for name, att, subject, body, expected in CASES:
        got = cm.resolve_week(att, subject, body)
        if got != expected:
            failures.append(f"{name}: expected {expected}, got {got}")

    # A reschedule must be recognised as such.
    if not cm.looks_like_revision("Rescheduled timetable for the week"):
        failures.append("revision detection: 'Rescheduled timetable' not detected")
    if not cm.looks_like_revision("Revised time table"):
        failures.append("revision detection: 'Revised time table' not detected")
    if cm.looks_like_revision("Next week time table"):
        failures.append("revision detection: plain timetable wrongly flagged as revision")

    # The filename repair must produce a canonical, parseable name.
    fixed = cm.repair_timetable_filename("12.010.2026 to 18.10.2026.xlsx")
    if fixed != "12.10.2026 to 18.10.2026.xlsx":
        failures.append(f"filename repair: got {fixed!r}")
    if cm.parse_file_dates(fixed) != d(2026, 10, 12):
        failures.append("filename repair: repaired name is not parseable")

    # A well-formed name must be left alone.
    keep = "19.10.2026 to 25.10.2026.xlsx"
    if cm.repair_timetable_filename(keep) != keep:
        failures.append("filename repair: well-formed name was modified")

    # Week arithmetic: the week of a Sunday belongs to the week that
    # started on the preceding Monday.
    if cm.monday_of(d(2026, 10, 25)) != d(2026, 10, 19):
        failures.append("monday_of: Sunday should map to the same week")
    if cm.week_of_path("downloads/19.10.2026 to25.10.2026.xlsx") != d(2026, 10, 19):
        failures.append("week_of_path: wrong week for the wired file")

    if failures:
        print("FAIL: mail routing violations:")
        for f in failures:
            print(f"  - {f}")
        return 1

    print(f"OK: {len(CASES)} routing cases pass.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
