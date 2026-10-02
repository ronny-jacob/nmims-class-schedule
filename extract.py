import json, re, os, sys
from collections import OrderedDict
from datetime import datetime, timedelta
import openpyxl

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
def src(name):
    return os.path.join(BASE_DIR, "sources", name)
def resolve(p):
    if not p:
        return p
    return p if os.path.isabs(p) else os.path.join(BASE_DIR, p)

STUDENT_LIST = src("Trim-V student list.xlsx")
# TIMETABLE is normally written here by check_mail.py when a new weekly
# timetable email arrives. When empty (e.g. fresh clone, or the bot
# hasn't run yet), fall back to the most recently dated file in
# downloads/, then to sources/Trim V time table.xlsx (manual starter).
TIMETABLE    = "downloads/05.10.2026-11.10.2026.xlsx"
TIMETABLE_NEXT = ""
FOOD_MENU    = src("August-Sept Menu Updated.xlsx")
FOOD_MENU_ANCHOR = "2026-08-03"
PLACEMENTS   = "sources/placements.json"
OUTPUT       = "data.json"
PLACEMENT_WINDOW_DAYS = 1

HOLIDAYS = {
    "2026-08-15": "Independence Day",
}

TIMETABLE = resolve(TIMETABLE)
FOOD_MENU = resolve(FOOD_MENU)


def _file_has_trim_v_layout(path):
    """Quick sniff: does this xlsx have a 'TT' sheet (Trim V layout)?"""
    try:
        wb = openpyxl.load_workbook(path, read_only=True)
    except Exception:
        return False
    layout = 'TT' in wb.sheetnames
    wb.close()
    return layout


def pick_timetable_fallback():
    """Pick the most useful timetable when TIMETABLE is empty.

    The mail bot (check_mail.py) is the production source. When it has
    not yet written a TIMETABLE constant, we fall back to the most
    recently dated file in downloads/ that has the Trim V layout
    (i.e., a 'TT' sheet). That is exactly what the bot would do.

    For local development (CI env var unset), we additionally allow
    sources/Trim V time table.xlsx as a starter so a fresh clone yields a
    runnable build without needing to invoke the bot. CI / staging /
    production builds do NOT fall back to the manual starter.
    """
    in_ci = bool(os.environ.get("CI"))

    downloads_dir = os.path.join(BASE_DIR, "downloads")
    pattern = re.compile(r'\d+\.\d+\.\d{4}\s*(?:to|-|–)\s*\d+\.\d+\.\d{4}')
    candidates = []
    if os.path.isdir(downloads_dir):
        for name in os.listdir(downloads_dir):
            if not name.endswith('.xlsx') or name.startswith('.'):
                continue
            full = os.path.join(downloads_dir, name)
            if not pattern.search(name):
                continue
            try:
                mtime = os.path.getmtime(full)
            except OSError:
                continue
            if _file_has_trim_v_layout(full):
                candidates.append((mtime, os.path.relpath(full, BASE_DIR)))
    if candidates:
        candidates.sort(reverse=True)
        return candidates[0][1]

    if not in_ci:
        starter = src("Trim V time table.xlsx")
        if os.path.isfile(starter):
            return os.path.relpath(starter, BASE_DIR)

    return ""

FOOD_MEALS = [
    {"key": "breakfast", "label": "Breakfast", "time": "8:00 To 9:30"},
    {"key": "lunch",     "label": "Lunch",     "time": "12:30 To 2:30"},
    {"key": "snacks",    "label": "Snacks",    "time": "5:30 To 6:00"},
    {"key": "dinner",    "label": "Dinner",    "time": "8:00 To 9:30"},
]

# ─── Trim V subjects ─────────────────────────────────────────────────
SUBJECT_NAMES = {
    "IB_A":   "Investment Banking A",
    "IB_B":   "Investment Banking B",
    "IF":     "International Finance",
    "WM":     "Wealth Management",
    "DM":     "Digital Marketing",
    "SBM_A":  "Strategic Brand Management A",
    "SBM_B":  "Strategic Brand Management B",
    "SM":     "Services Marketing",
    "CTA_A":  "Corporate Turnaround A",
    "CTA_B":  "Corporate Turnaround B",
    "MACR":   "Mergers, Acquisitions and Corporate Restructuring",
    "GS":     "Games of Strategy",
    "PM":     "Performance Management",
    "VA":     "Visual Analytics",
}

# Timetable cells begin with one of these prefixes.
TIMETABLE_SUBJECT_MAP = {
    "IB Div A":  ("IB_A",  "Div A"),
    "IB Div B":  ("IB_B",  "Div B"),
    "IBDiv A":   ("IB_A",  "Div A"),  # typo in source xlsx (no space)
    "IBDiv B":   ("IB_B",  "Div B"),
    "CT-A":      ("CTA_A", "Div A"),
    "CT-B":      ("CTA_B", "Div B"),
    "IF":        ("IF",    None),
    "WM":        ("WM",    None),
    "DM":        ("DM",    None),
    "MACR":      ("MACR",  None),
    "GOS":       ("GS",    None),  # subject detail sheet uses GOS, student list uses GS
    "VA":        ("VA",    None),
    "PM":        ("PM",    None),
    "SM":        ("SM",    None),
    "SBM Div A": ("SBM_A", "Div A"),
    "SBM Div B": ("SBM_B", "Div B"),
}

FOOD_DAY_KEYS = {
    'MON': 'Mon', 'TUE': 'Tue', 'WED': 'Wed', 'THU': 'Thu',
    'FRI': 'Fri', 'SAT': 'Sat', 'SUN': 'Sun',
}

DAY_LABELS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']


def normalize(s, underscores=False):
    s = re.sub(r'\s+', ' ', s).strip()
    if underscores:
        s = re.sub(r'\s*_\s*', '_', s)
    return s


def parse_time_labels_from_sheet(ws, header_row=2):
    """Read column time labels from `header_row`. Empty cells inherit previous label."""
    labels = OrderedDict()
    prev = ''
    for col_idx in range(2, ws.max_column + 1):
        cell = ws.cell(row=header_row, column=col_idx)
        if cell and cell.value:
            raw = normalize(str(cell.value))
            raw = re.sub(r'\s*[ap]m\s*', '', raw, flags=re.I)
            raw = re.sub(r'^0(\d)', r'\1', raw)
            raw = re.sub(r'(\s)0(\d)', r'\1\2', raw)
            raw = re.sub(r'\s*-\s*', '-', raw)
            labels[col_idx] = raw
            prev = raw
        elif prev:
            labels[col_idx] = prev
    return labels


def dedupe_consecutive(lst):
    if not lst:
        return lst
    result = [lst[0]]
    for item in lst[1:]:
        if item != result[-1]:
            result.append(item)
    return result


def start_and_end_from_filepath(filepath):
    """Return (start_date, end_date) parsed from a timetable filename like
    '07.09.2026 to 13.09.2026.xlsx' or '05.10.2026-11.10.2026.xlsx'.
    If the parsed end is not after the start, fall back to start + 6 days."""
    mt = re.search(r'(\d+)\.(\d+)\.(\d+)\s*(?:to|-|–)\s*(\d+)\.(\d+)\.(\d+)', filepath)
    if not mt:
        return (None, None)
    d1, m1, y1, d2, m2, y2 = mt.groups()
    try:
        dt_from = datetime.strptime(f"{d1}.{m1}.{y1}", "%d.%m.%Y")
        dt_to   = datetime.strptime(f"{d2}.{m2}.{y2}", "%d.%m.%Y")
    except ValueError:
        return (None, None)
    if dt_to <= dt_from:
        dt_to = dt_from + timedelta(days=6)
    return (dt_from, dt_to)


def parse_date_range(filepath):
    dt_from, dt_to = start_and_end_from_filepath(filepath)
    if dt_from is None:
        return ""
    return (dt_from.strftime("%a %-d %b") + " – "
            + dt_to.strftime("%a %-d %b %Y"))


def get_week_iso(filepath):
    dt_from, dt_to = start_and_end_from_filepath(filepath)
    if dt_from is None:
        return ("", "")
    return (dt_from.strftime("%Y-%m-%d"), dt_to.strftime("%Y-%m-%d"))


def match_subject(text, key):
    """Loose prefix match ignoring case, breaking on non-alnum boundary."""
    if text.startswith(key):
        return True
    if key.startswith(text):
        if len(text) == len(key):
            return True
        nxt = key[len(text)]
        return not nxt.isalnum()
    return False


def parse_timetable(filepath):
    """Trim V single-sheet timetable (sheet 'TT', day labels in column A)."""
    wb = openpyxl.load_workbook(filepath)
    if 'TT' in wb.sheetnames:
        ws = wb['TT']
    else:
        ws = wb[wb.sheetnames[0]]

    time_labels = parse_time_labels_from_sheet(ws, header_row=2)

    # Find row ranges for each day. Day labels look like "Mon. ", "Tue. ", etc.
    # Multi-line cells may have e.g. "Mon.\n" or just "Mon". Continuation rows
    # have None in column A.
    day_row_ranges = []
    last_day = None
    last_start = None
    last_end = None

    for row_idx in range(3, ws.max_row + 1):
        v = ws.cell(row=row_idx, column=1).value
        matched_day = None
        if v and isinstance(v, str):
            for d in DAY_LABELS:
                if v.strip().startswith(d):
                    matched_day = d
                    break
        if matched_day:
            if last_day is not None:
                day_row_ranges.append((last_day, last_start, last_end or last_start))
            last_day = matched_day
            last_start = row_idx
            last_end = row_idx
        elif last_day is not None:
            last_end = row_idx
    if last_day is not None:
        day_row_ranges.append((last_day, last_start, last_end or last_start))

    if not day_row_ranges:
        # Fallback: assume rows 3..12, single day block.
        day_row_ranges = [('Mon', 3, 12)]

    timetable = []
    for day_name, r1, r2 in day_row_ranges:
        for row_idx in range(r1, r2 + 1):
            for col_idx in range(2, ws.max_column + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                raw = cell.value
                if not raw:
                    continue
                raw_lines = [l.strip() for l in str(raw).split('\n') if l.strip()]
                # Skip decorative cells (e.g. vertical LUNCH BREAK text)
                if all(len(l) <= 1 for l in raw_lines):
                    continue
                lines = [normalize(l) for l in raw_lines]
                if not lines:
                    continue
                time_label = time_labels.get(col_idx, "?")

                subject_text = normalize(lines[0], underscores=True)
                alt_text = None
                used_lines = 1
                if (len(lines) > 1
                        and not lines[1].startswith('L')
                        and lines[1] != 'Hybrid'
                        and not re.match(r'^(Dr|Prof)\b', lines[1])):
                    alt_text = normalize(subject_text + ' ' + lines[1], underscores=True)

                match = None
                for key, (code, div) in TIMETABLE_SUBJECT_MAP.items():
                    key_norm = normalize(key, underscores=True)
                    if match_subject(subject_text, key_norm):
                        match = (code, div)
                        break
                    if alt_text and match_subject(alt_text, key_norm):
                        match = (code, div)
                        used_lines = 2
                        break

                if not match:
                    continue

                code, div = match
                remaining = lines[used_lines:]
                room = ''
                professor = ''
                for line in remaining:
                    if line == 'Hybrid':
                        room = 'Hybrid'
                    elif re.match(r'^L\s?\d', line) or re.match(r'^LR\s?\d', line):
                        room = line
                    elif 'Prof' in line or re.match(r'^Dr\b', line) or 'Dr ' in line:
                        if not professor:
                            professor = line
                entry = {
                    "day": day_name,
                    "time": time_label,
                    "subject": code,
                    "raw_text": '\n'.join(lines),
                    "professor": professor or "",
                    "room": room or "",
                    "div": div or "",
                }
                existing = None
                for e in timetable:
                    if (e["day"] == entry["day"]
                            and e["time"] == entry["time"]
                            and e["subject"] == entry["subject"]
                            and e["professor"] == entry["professor"]
                            and e["room"] == entry["room"]):
                        existing = e
                        break
                if not existing:
                    timetable.append(entry)

    return timetable


def parse_food_menu(filepath):
    """Read the food menu xlsx → {week1: {day: {meal: text}}, week2: {...}}."""
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb["Table 1"] if "Table 1" in wb.sheetnames else wb[wb.sheetnames[0]]

    menu = {"week1": OrderedDict(), "week2": OrderedDict()}
    current = menu["week1"]

    for row in ws.iter_rows(min_row=1, values_only=True):
        day_key = row[0]
        if not day_key:
            continue
        first = normalize(str(day_key))
        if first.lower().startswith('2nd'):
            current = menu["week2"]
            continue
        day_short = FOOD_DAY_KEYS.get(first[:3].upper())
        if day_short is None:
            continue
        entry = OrderedDict()
        for i, meal in enumerate(FOOD_MEALS):
            raw = row[1 + i] if 1 + i < len(row) else None
            text = normalize(str(raw)) if raw is not None else ""
            text = text.replace('\xa0', ' ').strip()
            entry[meal["key"]] = text
        current[day_short] = entry

    return menu


def parse_students():
    """Read Trim V roster from sources/Trim-V student list.xlsx.

    Layout (verified):
      row 1: junk counts row (skip)
      row 2: headers — Email, Name, Last modified, SAP, Name2, Roll, Email,
                       Minor, [subject codes 9..22]
      row 3+: student rows. Name2 is the canonical display name.
    """
    wb = openpyxl.load_workbook(STUDENT_LIST)
    ws = wb[wb.sheetnames[0]]

    headers = [str(c.value).strip() if c.value else '' for c in ws[2]]
    subject_cols = [(i, h) for i, h in enumerate(headers) if h in SUBJECT_NAMES]

    students = []
    for row in ws.iter_rows(min_row=3, values_only=True):
        if not any(row):
            continue
        # Trim V column map (0-indexed):
        #   0:Email 1:Name 2:LastMod 3:SAP 4:Name2 5:Roll 7:Major 8:Minor
        canonical = str(row[4]).strip() if row[4] else ''
        name_caps = str(row[1]).strip() if row[1] else ''
        if not canonical and not name_caps:
            continue
        # Skip purely junk rows (no canonical, no SAP, no roll)
        sap = str(row[3]).strip() if row[3] else ''
        roll_raw = str(row[5]).strip() if row[5] else ''
        if not canonical and not sap and not roll_raw:
            continue

        name = normalize(canonical) if canonical else normalize(name_caps)
        roll = normalize(roll_raw)
        major = normalize(str(row[7])) if row[7] else ''
        minor = str(row[8]).replace('\xa0', ' ').strip() if row[8] else ''
        email = normalize(str(row[6])) if row[6] else ''
        # Trim V has no BS division. Empty placeholder kept for back-compat
        # with index.html's rendering code paths.
        bs_div = ''

        # Normalize roll to Hxxx format if it starts with H
        if roll.startswith('H'):
            digits = re.sub(r'[^0-9]', '', roll)
            if digits:
                roll = 'H' + digits.zfill(3)

        subjects = []
        for col_idx, code in subject_cols:
            val = ''
            if col_idx < len(row) and row[col_idx] is not None:
                val = str(row[col_idx]).strip().upper()
            if val == 'YES':
                subjects.append(code)

        students.append({
            "name": name,
            "roll": roll,
            "email": email,
            "major": major,
            "minor": minor,
            "gender": "",  # Trim V roster doesn't include it
            "subjects": subjects,
            "bs_div": bs_div,
        })

    # Students no longer part of the college are excluded from the roster;
    # the source xlsx (regenerated from CI secret on every build) still lists
    # them, so filtering happens here rather than by editing the file.
    DEPARTED = ["Abhijatya Negi"]
    students = [s for s in students if s["name"] not in DEPARTED]

    return students


def parse_placements():
    """Flatten placements.json → [{name, company, date, roster, type}]."""
    placements = []
    try:
        with open(PLACEMENTS) as f:
            raw = json.load(f)
    except FileNotFoundError:
        return placements
    except Exception as e:
        print(f"⚠️ Could not parse placements: {e}")
        return placements

    for a in raw.get("announcements", []):
        company = str(a.get("company", "")).strip()
        date_str = str(a.get("announced_on", "")).strip()
        ptype = str(a.get("type", "ppo")).strip().lower() or "ppo"
        for name in a.get("names", []):
            roster_name = None
            if isinstance(name, dict):
                nm = str(name.get("name", "")).strip()
                roster_name = str(name.get("roster", "")).strip() or None
            else:
                nm = str(name).strip()
            if not nm:
                continue
            placements.append({
                "name": nm,
                "company": company,
                "date": date_str,
                "roster": roster_name,
                "type": ptype,
            })
    return placements


def build_classes_and_subjects(students):
    """Generate the JS CLASSES / SUBJECTS arrays from DATA.students.

    Returns (classes_by_trim, subjects_array, mx_count).
    """
    by_trim = {"v": [], "vi": []}
    subjects_arr = []
    for code, full_name in SUBJECT_NAMES.items():
        enrolled = sorted(
            s["name"] for s in students if code in s["subjects"]
        )
        by_trim["v"].append({
            "cls": full_name + " (Faculty)",
            "students": enrolled,
        })
        subjects_arr.append({
            "trim": "v",
            "trimLabel": "Trim V",
            "cls": full_name + " (Faculty)",
            "short": full_name,
            "faculty": "Faculty",
            "count": len(enrolled),
            "students": enrolled,
        })
    mx_count = max((len(c["students"]) for c in by_trim["v"]), default=0)
    return by_trim, subjects_arr, mx_count


def main():
    students = parse_students()

    # Timetable source. The committed TIMETABLE constant is updated by the
    # mail bot (check_mail.py) whenever a new weekly xlsx arrives in
    # downloads/. When that constant is empty we fall back to a fresh
    # dated file in downloads/ (CI / staging / prod) or to the manual
    # starter in sources/ (local dev only). If we end up with nothing,
    # the build fails loudly rather than shipping an empty timetable.
    global TIMETABLE, TIMETABLE_NEXT
    timeline_today = datetime.now().date()

    def this_week_sunday():
        monday = timeline_today - timedelta(days=timeline_today.weekday())
        return monday + timedelta(days=6)

    def file_belongs_to_future_week(path):
        ws, _we = get_week_iso(path)
        return bool(ws) and ws > this_week_sunday().strftime("%Y-%m-%d")

    timetable_path = TIMETABLE
    if not timetable_path or not os.path.isfile(resolve(timetable_path) if timetable_path else ""):
        fallback = pick_timetable_fallback()
        if fallback:
            if file_belongs_to_future_week(fallback) and not TIMETABLE_NEXT:
                # The newest dated file is for an upcoming week, not the
                # current one — keep it in the NEXT slot (mirrors the mail
                # bot, which routes future-week emails to TIMETABLE_NEXT).
                print(f"ℹ️ TIMETABLE unset; {fallback} is a future week → "
                      f"wired to TIMETABLE_NEXT")
                TIMETABLE_NEXT = fallback  # already a BASE_DIR-relative path
            else:
                timetable_path = fallback
                print(f"ℹ️ TIMETABLE unset, using fallback: {fallback}")
        elif not TIMETABLE_NEXT:
            sys.exit("❌ No timetable source. Run check_mail.py first to "
                     "download the latest timetable email into downloads/, "
                     "or wire extract.py's TIMETABLE constant.")
    TIMETABLE = resolve(timetable_path) if timetable_path else ""

    # Timetable date range. The file may or may not contain a date pattern
    # in its name. When it does, parse it; otherwise anchor to the current
    # week's Monday → Sunday.
    today = datetime.now().date()
    monday = today - timedelta(days=today.weekday())
    try:
        timetable = parse_timetable(TIMETABLE)
        week_start_iso, week_end_iso = get_week_iso(TIMETABLE)
        if not week_start_iso:
            week_start_iso = monday.strftime("%Y-%m-%d")
            week_end_iso = (monday + timedelta(days=6)).strftime("%Y-%m-%d")
        date_range = parse_date_range(TIMETABLE) or (
            monday.strftime("%a %-d %b") + " – "
            + (monday + timedelta(days=6)).strftime("%a %-d %b %Y")
        )
    except Exception as e:
        print(f"❌ Could not parse timetable: {e}")
        timetable = []
        date_range = ""
        week_start_iso = ""
        week_end_iso = ""

    timetable_next = []
    date_range_next = ""
    week_start_next = ""
    week_end_next = ""
    if TIMETABLE_NEXT and os.path.isfile(resolve(TIMETABLE_NEXT)):
        try:
            timetable_next = parse_timetable(resolve(TIMETABLE_NEXT))
            week_start_next, week_end_next = get_week_iso(resolve(TIMETABLE_NEXT))
            date_range_next = parse_date_range(resolve(TIMETABLE_NEXT))
        except Exception as e:
            print(f"⚠️ Could not parse next week's timetable: {e}")
            timetable_next = []
            date_range_next = ""
            week_start_next = ""
            week_end_next = ""

    subjects = {}
    for code, full in SUBJECT_NAMES.items():
        students_in_subject = [s for s in students if code in s["subjects"]]
        subjects[code] = {
            "code": code,
            "name": full,
            "student_count": len(students_in_subject),
        }

    # Time slots from timetable sheet
    try:
        wb = openpyxl.load_workbook(TIMETABLE)
        ws = wb['TT'] if 'TT' in wb.sheetnames else wb[wb.sheetnames[0]]
        time_slots = dedupe_consecutive(list(parse_time_labels_from_sheet(ws).values()))
    except Exception:
        time_slots = []
    time_slots_next = []
    if TIMETABLE_NEXT and os.path.isfile(resolve(TIMETABLE_NEXT)):
        try:
            wb_next = openpyxl.load_workbook(resolve(TIMETABLE_NEXT))
            ws_next = wb_next['TT'] if 'TT' in wb_next.sheetnames else wb_next[wb_next.sheetnames[0]]
            time_slots_next = dedupe_consecutive(list(parse_time_labels_from_sheet(ws_next).values()))
        except Exception:
            pass

    food_menu = {}
    if os.path.exists(FOOD_MENU):
        try:
            food_menu = parse_food_menu(FOOD_MENU)
        except Exception as e:
            print(f"⚠️ Could not parse food menu: {e}")

    placements = parse_placements()

    classes_by_trim, subjects_arr, mx_count = build_classes_and_subjects(students)
    students_sorted = sorted(s["name"] for s in students)

    data = {
        "students": students,
        "subjects": subjects,
        "timetable": timetable,
        "timetable_next": timetable_next,
        "time_slots": time_slots,
        "time_slots_next": time_slots_next,
        "food_menu": food_menu,
        "food_meals": FOOD_MEALS,
        "food_menu_anchor": FOOD_MENU_ANCHOR,
        "placements": placements,
        "placements_window_days": PLACEMENT_WINDOW_DAYS,
        "holidays": HOLIDAYS,
        "days": ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"],
        "date_range": date_range,
        "date_range_next": date_range_next,
        "week_start": week_start_iso,
        "week_end": week_end_iso,
        "week_start_next": week_start_next,
        "week_end_next": week_end_next,
        "classes": classes_by_trim,
        "subjects_arr": subjects_arr,
        "mx": mx_count,
    }

    # Read last_updated metadata if available (written by check_mail.py)
    last_updated_path = os.path.join(os.path.dirname(__file__), "last_updated.json")
    if os.path.exists(last_updated_path):
        try:
            with open(last_updated_path) as f:
                data["last_updated"] = json.load(f)
        except Exception:
            pass

    with open(OUTPUT, 'w') as f:
        json.dump(data, f, indent=2)
    print(f"✅ Wrote {OUTPUT}")
    print(f"   Students: {len(students)}")
    print(f"   Subjects: {len(subjects)}")
    print(f"   Timetable entries: {len(timetable)}")
    if timetable_next:
        print(f"   Next week entries: {len(timetable_next)}")
    print(f"   Food menu weeks: {len(food_menu)}")
    print(f"   Placements: {len(placements)}")

    # ─── Regenerate index.html with embedded data + JS arrays ───
    INDEX_HTML = "index.html"
    data_json = json.dumps(data)
    students_json = json.dumps(students_sorted)
    classes_json = json.dumps(classes_by_trim)
    subjects_arr_json = json.dumps(subjects_arr)

    with open(INDEX_HTML) as f:
        html = f.read()

    markers = {
        "DATA":   "/* DATA_INSERT_HERE */",
        "STUDENTS": "/* STUDENTS_HERE */",
        "CLASSES":  "/* CLASSES_HERE */",
        "SUBJECTS": "/* SUBJECTS_HERE */",
        "MX":      "/* MX_HERE */",
    }

    # If a previous successful run already replaced the JS-array markers,
    # re-inject them so this run can replace them again. The generated
    # `var STUDENTS = [..];var CLASSES = ...;` lines may sit on a single
    # line, so we don't anchor on `^`.
    import re as _re

    def _find_top_level_close(text, start, open_ch, close_ch):
        depth = 0
        i = start
        in_string = False
        escape = False
        while i < len(text):
            c = text[i]
            if in_string:
                if escape:
                    escape = False
                elif c == '\\':
                    escape = True
                elif c == '"':
                    in_string = False
            else:
                if c == '"':
                    in_string = True
                elif c == open_ch:
                    depth += 1
                elif c == close_ch:
                    depth -= 1
                    if depth == 0:
                        return i
            i += 1
        return -1

    for arr_name, open_ch, close_ch in (
        ('STUDENTS', '[', ']'),
        ('CLASSES',  '{', '}'),
        ('SUBJECTS', '[', ']'),
    ):
        marker = markers[arr_name]
        if marker in html:
            continue
        # Find "var STUDENTS = " or "var CLASSES = " etc. anywhere in the file.
        # (No MULTILINE because of the all-on-one-line issue.)
        token = 'var ' + arr_name + ' = '
        idx = html.find(token)
        if idx < 0:
            continue
        end = _find_top_level_close(html, idx + len(token), open_ch, close_ch)
        if end < 0:
            continue
        end_idx = end + 1
        if end_idx < len(html) and html[end_idx] == ';':
            end_idx += 1
        html = html[:idx] + marker + html[end_idx:]

    # MX: var MX = NN;
    if markers['MX'] not in html:
        m = _re.search(r'var MX\s*=\s*\d+\s*;', html)
        if m:
            html = html[:m.start()] + markers['MX'] + html[m.end():]

    for name, marker in markers.items():
        if marker not in html:
            print(f"❌ Marker '{marker}' not found in {INDEX_HTML}")
            return

    before, sep, after = html.partition(markers["DATA"])
    if sep != markers["DATA"]:
        print(f"❌ DATA marker missing")
        return

    # After DATA marker, skip until var STUDENTS marker line.
    lines_after = after.split('\n')
    real_start = len(lines_after)
    for i, line in enumerate(lines_after):
        if markers["STUDENTS"] in line:
            real_start = i
            break

    after = '\n'.join(lines_after[real_start:])

    # Replace STUDENTS / CLASSES / SUBJECTS markers with generated arrays.
    js_tail = (
        f"var STUDENTS = {students_json};"
        f"var CLASSES = {classes_json};"
        f"var SUBJECTS = {subjects_arr_json};"
        f"var MX = {mx_count};"
    )
    # The four markers sit on their own lines in this order in index.html.
    for marker in (markers["STUDENTS"], markers["CLASSES"],
                   markers["SUBJECTS"], markers["MX"]):
        after = after.replace(marker, "", 1)
    # Insert emitted JS arrays just after the last replacement removed
    # (right before any remaining script code).
    new_html = (
        before + markers["DATA"] + f"\nDATA = {data_json};"
        f"\n{js_tail}\n"
        + after
    )

    with open(INDEX_HTML, 'w') as f:
        f.write(new_html)
    print(f"✅ Updated {INDEX_HTML}")


if __name__ == '__main__':
    main()