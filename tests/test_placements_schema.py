"""Schema and sanity checks for sources/placements.json.

Run from repo root:
    python tests/test_placements_schema.py

The check_mail.py workflow runs this before regenerating data.json so a
mistyped date, a missing roster match, or a wrong type fails CI loudly
instead of silently disappearing from the site.
"""
from __future__ import annotations

import json
import os
import re
import sys
from collections import defaultdict
from datetime import date, datetime

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PLACEMENTS_PATH = os.path.join(REPO_ROOT, "sources", "placements.json")
DATA_PATH = os.path.join(REPO_ROOT, "data.json")

DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
NAME_RE = re.compile(r"^[A-Z][A-Za-z .'\-]*$")

REQUIRED_KEYS = {"company", "announced_on", "names"}
VALID_TYPES = {"ppo", "final"}


def load_json(path: str):
    try:
        with open(path) as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"FAIL: missing file: {path}")
        sys.exit(1)
    except json.JSONDecodeError as e:
        print(f"FAIL: invalid JSON in {path}: {e}")
        sys.exit(1)


def roster_names() -> set[str]:
    d = load_json(DATA_PATH)
    return {s["name"].strip() for s in d.get("students", [])}


def main() -> int:
    placements = load_json(PLACEMENTS_PATH)
    announcements = placements.get("announcements", [])
    if not announcements:
        print("FAIL: no announcements in sources/placements.json")
        return 1

    errors: list[str] = []
    roster = roster_names()
    seen_names: dict[str, str] = {}
    by_date: dict[str, int] = defaultdict(int)

    for i, a in enumerate(announcements):
        prefix = f"announcements[{i}]"
        missing = REQUIRED_KEYS - set(a.keys())
        if missing:
            errors.append(f"{prefix}: missing keys {sorted(missing)}")
            continue

        company = (a.get("company") or "").strip()
        if not company:
            errors.append(f"{prefix}: empty company")

        date_str = (a.get("announced_on") or "").strip()
        if not DATE_RE.match(date_str):
            errors.append(
                f"{prefix}: announced_on='{date_str}' is not YYYY-MM-DD"
            )
        else:
            try:
                announced = datetime.strptime(date_str, "%Y-%m-%d").date()
            except ValueError as e:
                errors.append(f"{prefix}: {e}")
                announced = None
            if announced and announced > date.today():
                errors.append(
                    f"{prefix}: announced_on={date_str} is in the future"
                )
            if announced:
                by_date[date_str] += len(a.get("names", []) or [])

        type_ = (a.get("type") or "").strip().lower()
        if type_ and type_ not in VALID_TYPES:
            errors.append(
                f"{prefix}: type='{type_}' not in {sorted(VALID_TYPES)}"
            )

        names = a.get("names") or []
        if not names:
            errors.append(f"{prefix}: empty names list")
        for j, n in enumerate(names):
            np = f"{prefix}.names[{j}]"
            if isinstance(n, str):
                display = n.strip()
                roster_match = None
            elif isinstance(n, dict):
                display = (n.get("name") or "").strip()
                roster_match = (n.get("roster") or "").strip() or None
            else:
                errors.append(f"{np}: must be string or object, got {type(n).__name__}")
                continue

            if not display:
                errors.append(f"{np}: empty name")
                continue
            if not NAME_RE.match(display):
                errors.append(
                    f"{np}: '{display}' has unexpected characters"
                )

            candidate_names = {display}
            if roster_match:
                candidate_names.add(roster_match)
            if not (candidate_names & roster):
                errors.append(
                    f"{np}: '{display}'"
                    + (f" (roster='{roster_match}')" if roster_match else "")
                    + " does not match any student in data.json"
                )

            key = display.lower()
            if key in seen_names:
                errors.append(
                    f"{np}: '{display}' was already announced on "
                    f"{seen_names[key]}"
                )
            else:
                seen_names[key] = date_str

    # Sanity: warn on bursts (more than ~10 placements in one day)
    for d, n in sorted(by_date.items()):
        if n > 12:
            errors.append(
                f"warning: {n} placements on {d} (>12) — verify this is intentional"
            )

    if errors:
        print("FAIL: placements schema violations:")
        for e in errors:
            print(f"  - {e}")
        return 1

    print(f"OK: {len(announcements)} announcements, "
          f"{sum(len(a.get('names', [])) for a in announcements)} names, "
          f"all match the roster of {len(roster)} students.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
