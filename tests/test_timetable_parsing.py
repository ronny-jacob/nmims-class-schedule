"""Parsing tests for the Trim V timetable sheet (extract.py).

Run from repo root:
    python tests/test_timetable_parsing.py

Covers two real defects in the sender's sheets:

1. A repeated day label ("Sat." twice in a row) — the second one is
   really the next day of the week, so the grid must show Sat then Sun.
2. A blank column header in the evening block — the classes in it state
   their own time, so they must get their own slot instead of piling up
   under the last labelled column.
"""
from __future__ import annotations

import os
import sys
import tempfile

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO_ROOT)

import extract  # noqa: E402


def build_sheet(path):
    """A sheet shaped like the sender's: duplicated 'Sat.' label and a
    blank-headed evening column."""
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "TT"
    ws.append(["Day/Time"])
    ws.append(["Day/Time",
               "10:10 am- 11:40 am",
               "11:50 am - 1:20 pm",
               "2:00 pm- 3:30pm",
               "03:40 pm - 5:10pm",
               None,
               None])
    ws.append(["Mon.", "IF\nProf. Vittal\nLR 07", None, "MACR\nL07", None, None, None])
    ws.append(["Tue.", None, "PM\nL12", None, "GOS\nLR06", None, None])
    ws.append(["Wed.", None, None, "VA\nL06", None,
               "MACR\n(6:00pm to 7:30pm)\nL07", None])
    ws.append(["Sat.", "SBM\nL06", None, "SBM\nL06", None,
               "SM\n(6:00 pm to 7:30 pm)\nL06", None])
    ws.append(["Sat.", "SBM\nL06", None, "SBM\nL06", None,
               "SM\n(7:35:00 to 9:05:00pm)\nL06", None])
    wb.save(path)
    return path


def main() -> int:
    failures = []
    tmp = tempfile.mkdtemp()
    path = os.path.join(tmp, "sheet.xlsx")
    build_sheet(path)

    entries = extract.parse_timetable(path)

    days = []
    for e in entries:
        if e["day"] not in days:
            days.append(e["day"])
    if days != ["Mon", "Tue", "Wed", "Sat", "Sun"]:
        failures.append(f"days: expected ['Mon','Tue','Wed','Sat','Sun'], got {days}")

    slots = extract.build_time_slots(
        _tt_sheet(path), used_labels={e["time"] for e in entries})
    if "6:00 to 7:30" not in slots:
        failures.append(f"evening slot missing from grid columns: {slots}")
    if "7:35 to 9:05" not in slots:
        failures.append(f"late evening slot missing from grid columns: {slots}")

    for e in entries:
        text = e.get("raw_text") or ""
        if "6:00pm" in text and e["time"] != "6:00 to 7:30":
            failures.append(f"evening MACR landed in {e['time']!r}, expected '6:00 to 7:30'")
        if e["subject"] == "MACR" and "6:00pm" not in text and e["time"] != "2:00-3:30":
            failures.append(f"daytime MACR landed in {e['time']!r}, expected '2:00-3:30'")
        if e["subject"] == "SM" and e["time"] not in ("6:00 to 7:30", "7:35 to 9:05"):
            failures.append(f"SM landed in {e['time']!r}")

    # A labelled column stays authoritative.
    if extract.cell_time_label("3:40-5:10", True, "X\n(6:00pm to 7:30pm)") != "3:40-5:10":
        failures.append("a labelled column must keep its own slot")
    if extract.cell_time_label("3:40-5:10", False, "X\n(6:00pm to 7:30pm)") != "6:00 to 7:30":
        failures.append("a blank-headed column must take the time from the cell")

    if failures:
        print("FAIL: timetable parsing violations:")
        for f in failures:
            print(f"  - {f}")
        return 1

    print(f"OK: {len(entries)} entries parsed, grid columns {slots}")
    return 0


def _tt_sheet(path):
    import openpyxl
    wb = openpyxl.load_workbook(path)
    return wb["TT"]


if __name__ == "__main__":
    sys.exit(main())
