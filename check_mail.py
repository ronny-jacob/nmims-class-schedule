import imaplib, email, email.utils, json, os, re, subprocess, sys, hashlib, tempfile
from datetime import datetime, timezone, timedelta
from email.header import decode_header

try:
    import openpyxl
except Exception:      # installed in CI; degrade gracefully if missing
    openpyxl = None

# ─── Config ───────────────────────────────────────────────────────────
# Outlook forwards timetable emails to Gmail, so we poll Gmail via IMAP.
IMAP_SERVER   = os.getenv("IMAP_SERVER", "imap.gmail.com")
IMAP_USER     = os.getenv("IMAP_USER", "")
IMAP_PASS     = os.getenv("IMAP_PASS", "")
LOOKBACK_DAYS = int(os.getenv("IMAP_LOOKBACK_DAYS", "7"))
TIMETABLE_DIR = os.path.dirname(os.path.abspath(__file__))
# Month group allows 3 digits so sender typos like '12.010.2026' still match.
TIMETABLE_PATTERN = re.compile(r'\d{1,2}\.\d{1,3}\.\d{4}\s*(?:to|-|–)\s*\d{1,2}\.\d{1,3}\.\d{4}')
DATE_TOKEN_RE   = re.compile(r'(\d{1,2})[./-](\d{1,3})[./-](\d{4})')
RANGE_RE        = re.compile(r'(\d{1,2})[./-](\d{1,3})[./-](\d{4})\s*(?:to|-|–)\s*(\d{1,2})[./-](\d{1,3})[./-](\d{4})')
TIMETABLE_WORDS = re.compile(r'timetable|time\s*table|schedule|class\s*timing|slot\s*chart|teaching\s*plan', re.I)
REVISION_WORDS  = re.compile(r'revis|reschedul|correct|modif|updat|amend|replace', re.I)

def decode_str(raw):
    """Decode email header to plain string."""
    if raw is None:
        return ""
    parts = decode_header(raw)
    result = []
    for part, charset in parts:
        if isinstance(part, bytes):
            try:
                result.append(part.decode(charset or "utf-8", errors="replace"))
            except Exception:
                result.append(part.decode("utf-8", errors="replace"))
        else:
            result.append(str(part))
    return " ".join(result).strip()

def search_timetable_mail():
    """Connect via IMAP and find the most recent timetable email.

    Returns the newest message that carries a timetable .xlsx. Newest
    wins, so a *rescheduled* timetable (which arrives later) always
    beats the original email it replaces. If the newest message has no
    recoverable date we keep looking at older ones, and only fall back
    to a date-less revision when nothing else is available.
    """
    mail = imaplib.IMAP4_SSL(IMAP_SERVER, 993)
    try:
        mail.login(IMAP_USER, IMAP_PASS)
    except imaplib.IMAP4.error as e:
        print(f"❌ IMAP login failed: {e}")
        sys.exit(1)

    mail.select("INBOX")

    # Search for messages from the last N days
    since_date = (datetime.now(timezone.utc) - timedelta(days=LOOKBACK_DAYS)).strftime("%d-%b-%Y")
    status, ids = mail.search(None, f'(SINCE {since_date})')
    if status != "OK" or not ids[0]:
        mail.logout()
        return None

    msg_ids = ids[0].split()
    msg_ids.reverse()  # newest first

    candidates = []
    for uid in msg_ids:
        status, data = mail.fetch(uid, "(RFC822)")
        if status != "OK":
            continue
        raw_email = data[0][1]
        msg = email.message_from_bytes(raw_email)

        subject = decode_str(msg.get("Subject"))
        from_addr = decode_str(msg.get("From"))
        date_str = decode_str(msg.get("Date"))
        body = extract_body_text(msg)

        atts = []
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_maintype() == "multipart":
                    continue
                fn = part.get_filename()
                if not fn:
                    continue
                fn_decoded = decode_str(fn)
                if not fn_decoded.lower().endswith(".xlsx"):
                    continue
                atts.append((fn_decoded, part.get_payload(decode=True)))
        if not atts:
            continue

        # Prefer an attachment whose filename carries a date range
        atts.sort(key=lambda a: 0 if TIMETABLE_PATTERN.search(a[0] or "") else 1)
        att_name, att_content = atts[0]
        if att_content is None:
            continue

        # Parse date for display
        try:
            parsed_dt = email.utils.parsedate_to_datetime(date_str)
            ist = parsed_dt.astimezone(timezone(timedelta(hours=5, minutes=30)))
            received_label = ist.strftime("%I:%M %p IST on %d %b %Y").lstrip("0").replace(" 0", " ")
        except Exception:
            received_label = date_str
            ist = None

        is_timetable = bool(TIMETABLE_PATTERN.search(att_name or ""))
        is_timetable_subj = bool(TIMETABLE_WORDS.search(f"{subject} {body}"))
        if not is_timetable and not is_timetable_subj:
            continue

        candidates.append({
            "is_timetable": is_timetable,
            "sender": from_addr,
            "received_label": received_label,
            "received_dt": ist,
            "subject": subject,
            "body": body,
            "att_name": att_name,
            "att_content": att_content,
        })

    mail.logout()

    if not candidates:
        return None

    # candidates are newest-first: take the newest one we can date.
    for c in candidates:
        if resolve_week(c["att_name"], c["subject"], c["body"]):
            return c
    # Nothing dateable — fall back to the newest revision-style email
    # (a reschedule whose dates are missing everywhere).
    for c in candidates:
        if looks_like_revision(c["subject"], c["body"], c["att_name"]):
            return c
    return candidates[0]


def file_hash(path):
    if not os.path.exists(path):
        return None
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def rebuild_and_commit(new_path, sender, received_at, subject):
    """Run extract.py, commit and push updates."""
    print(f"📥 New timetable: {os.path.basename(new_path)}")
    print(f"   From: {sender}, Received: {received_at}")

    last_updated = {
        "sender": sender,
        "received_at": received_at,
        "subject": subject,
        "updated_at": datetime.now(timezone.utc).isoformat(),
    }
    meta_path = os.path.join(TIMETABLE_DIR, "last_updated.json")
    with open(meta_path, "w") as f:
        json.dump(last_updated, f, indent=2)

    result = subprocess.run(
        [sys.executable, "extract.py"],
        capture_output=True, text=True,
        cwd=TIMETABLE_DIR,
    )
    print(result.stdout)
    if result.returncode != 0:
        print(f"❌ extract.py failed: {result.stderr}")
        return False

    subprocess.run(["git", "add", "-A"], cwd=TIMETABLE_DIR, capture_output=True)
    commit_msg = f"auto: timetable update from {sender}"
    result = subprocess.run(
        ["git", "commit", "-m", commit_msg],
        cwd=TIMETABLE_DIR, capture_output=True, text=True,
    )
    if result.returncode == 0:
        print(f"✅ Committed: {commit_msg}")
        subprocess.run(["git", "push"], cwd=TIMETABLE_DIR, capture_output=True)
        print("✅ Pushed to GitHub")
    else:
        print(f"ℹ️ No changes to commit: {result.stdout}")

    print(f"\n✅ Done — site updated")
    return True


def ist_now():
    utc = datetime.now(timezone.utc)
    ist = utc + timedelta(hours=5, minutes=30)
    return ist.strftime("%I:%M %p IST on %d %b %Y").lstrip("0").replace(" 0", " ")


def _norm_month(mo):
    """Normalise a month token that may carry a typo ('010' -> '10')."""
    for cand in (str(mo).lstrip("0"), str(mo)[-2:]):
        if cand.isdigit() and 1 <= int(cand) <= 12:
            return cand
    return None


def _mk_date(d, mo, y):
    """Build a date from dd.mm.yyyy-ish parts, tolerating month typos."""
    m = _norm_month(mo)
    if not m:
        return None
    try:
        return datetime.strptime(f"{int(d)}.{m}.{int(y)}", "%d.%m.%Y").date()
    except (ValueError, TypeError):
        return None


def parse_file_dates(name):
    """Extract the start date (dd.mm.yyyy) from a timetable filename."""
    m = TIMETABLE_PATTERN.search(name or "")
    if not m:
        return None
    start_str = re.split(r'\s*(?:to|-|–)\s*', m.group(0), 1)[0].strip()
    parts = start_str.split(".")
    if len(parts) != 3:
        return None
    return _mk_date(parts[0], parts[1], parts[2])


def parse_any_date(text):
    """Best-effort dd.mm.yyyy parse from any text (repairs typos like
    '12.010.2026' -> 12 Oct 2026, or picks the first date out of the
    email subject/body)."""
    if not text:
        return None
    m = DATE_TOKEN_RE.search(text)
    if not m:
        return None
    return _mk_date(m.group(1), m.group(2), m.group(3))


def resolve_week(att_name, subject, body):
    """Work out which week (its Monday) an incoming attachment belongs to.

    Order of evidence: the filename's date range, then a date range
    anywhere in the filename/subject/body, then any single date in the
    subject, then any single date in the body. This is what lets a
    *rescheduled* timetable that arrives with a vague name ("Revised
    Timetable.xlsx") or with the dates only in the subject still be
    routed to the right week.
    """
    d = parse_file_dates(att_name)
    if d:
        return monday_of(d)
    for text in (att_name or "", subject or "", body or ""):
        m = RANGE_RE.search(text)
        if m:
            dt = _mk_date(m.group(1), m.group(2), m.group(3))
            if dt:
                return monday_of(dt)
    for text in (subject or "", body or ""):
        dt = parse_any_date(text)
        if dt:
            return monday_of(dt)
    return None


def looks_like_revision(*texts):
    """True when the email/file looks like a correction of a timetable."""
    for t in texts:
        if t and REVISION_WORDS.search(t):
            return True
    return False


def file_has_timetable_layout(path):
    """Sniff whether an xlsx is a Trim V timetable (has a 'TT' sheet)."""
    if not openpyxl:
        return None          # unknown — don't block on it
    try:
        wb = openpyxl.load_workbook(path, read_only=True)
        ok = 'TT' in wb.sheetnames
        wb.close()
        return ok
    except Exception:
        return False


def extract_body_text(msg, limit=4000):
    """Plain-text body of an email, used as a last-resort date source."""
    chunks = []
    try:
        parts = msg.walk() if msg.is_multipart() else [msg]
        for part in parts:
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get_content_type() != 'text/plain':
                continue
            payload = part.get_payload(decode=True)
            if not payload:
                continue
            chunks.append(payload.decode(part.get_content_charset() or 'utf-8', errors='replace'))
    except Exception:
        pass
    return ' '.join(chunks)[:limit]


def repair_timetable_filename(name):
    """Return a canonical, parseable filename, repairing typos like the
    sender's '12.010.2026 to 18.10.2026.xlsx' (extra digit in the month)
    so the date parser can route the file to the right TIMETABLE slot."""
    if not name:
        return None
    # A 3-digit month ("12.010.2026") is not valid dd.mm.yyyy, so rewrite
    # it even though the date itself is recoverable.
    tokens = re.findall(r'(\d+)\.(\d+)\.(\d+)', name)
    if tokens and all(len(t[1]) <= 2 for t in tokens):
        if parse_file_dates(name):
            return name
    # Tolerant capture: allow oversized month groups (sender typos) so we
    # can isolate the FIRST token as the start date.
    tokens = re.findall(r'(\d+)\.(\d+)\.(\d+)', name or "")
    if len(tokens) < 2:
        return None

    def fix(token):
        d, mo, y = token
        return _mk_date(d, mo, y)

    start = fix(tokens[0])
    if not start:
        return None
    end = fix(tokens[1]) or start + timedelta(days=6)
    if end < start:
        end = start + timedelta(days=6)
    return f"{start.strftime('%d.%m.%Y')} to {end.strftime('%d.%m.%Y')}.xlsx"


def monday_of(d):
    return d - timedelta(days=d.weekday())


def read_extract():
    path = os.path.join(TIMETABLE_DIR, "extract.py")
    with open(path) as f:
        src = f.read()
    cur = re.search(r'^TIMETABLE\s*=\s*"([^"]*)"', src, re.M)
    cur_next = re.search(r'^TIMETABLE_NEXT\s*=\s*"([^"]*)"', src, re.M)
    return path, src, (cur.group(1) if cur else ""), (cur_next.group(1) if cur_next else "")


def week_of_path(p):
    if not p:
        return None
    start = parse_file_dates(os.path.basename(p))
    return monday_of(start) if start else None


def main():
    missing = [v for v in ["IMAP_USER", "IMAP_PASS"] if not os.getenv(v)]
    if missing:
        print(f"❌ Missing env vars: {', '.join(missing)}")
        print("   See SETUP.md for instructions.")
        sys.exit(1)

    result = search_timetable_mail()
    if not result:
        print(f"ℹ️ No timetable email found as of {ist_now()}")
        return

    sender = result["sender"]
    received_label = result["received_label"]
    subject = result["subject"]
    body = result["body"]
    att_name = result["att_name"]
    att_content = result["att_content"]

    # Repair sender filename typos (e.g. '12.010.2026 to ...xlsx') so the
    # date parser can route the file to the correct TIMETABLE slot.
    repaired_name = repair_timetable_filename(att_name)
    if repaired_name and repaired_name != att_name:
        print(f"🔧 Repaired attachment filename: {att_name!r} -> {repaired_name!r}")
        att_name = repaired_name

    extract_path, src, cur, cur_next = read_extract()
    cw = week_of_path(cur)
    nw = week_of_path(cur_next)

    # "Today" must be the IST date, otherwise the week is computed one
    # day early for a few hours around midnight.
    today = datetime.now(timezone(timedelta(hours=5, minutes=30))).date()
    this_week = monday_of(today)

    # Promote TIMETABLE_NEXT to TIMETABLE once its week has started.
    if nw and nw <= this_week and (cw is None or nw > cw):
        src = re.sub(r'^TIMETABLE\s*=\s*"[^"]*"', f'TIMETABLE    = "{cur_next}"', src, count=1, flags=re.M)
        src = re.sub(r'^TIMETABLE_NEXT\s*=\s*"[^"]*"', 'TIMETABLE_NEXT = ""', src, count=1, flags=re.M)
        with open(extract_path, "w") as f:
            f.write(src)
        print(f"♻️ Promoted {cur_next} to current timetable")
        cur, cur_next = cur_next, ""
        cw, nw = nw, None

    # Which week does this attachment belong to?
    new_week = resolve_week(att_name, subject, body)
    revision = looks_like_revision(subject, body, att_name)

    if new_week is None:
        # No date anywhere in filename / subject / body. If it is clearly a
        # reschedule, apply it to the week that is currently published —
        # a correction for "this week" is the common case. Otherwise do
        # NOT guess: leave the wiring untouched.
        if revision and cw:
            new_week = cw
            print(f"♻️ No date found in {att_name!r} — applying it to the current week")
        else:
            print(f"⚠️ {att_name!r} has no recoverable week date — not wired in")
            return

    if new_week < this_week and new_week != cw:
        print(f"ℹ️ {att_name} is for a past week ({new_week}) — not wired in")
        return

    # Where should the file live?
    #   • a reschedule of the week that is already wired → overwrite it, so
    #     the site shows the corrected timetable and the filename (which
    #     drives the displayed date range) stays stable
    #   • a reschedule of next week → overwrite that file
    #   • a brand-new week → keep the (repaired) sender filename
    if cw and new_week == cw and cur:
        target_rel = cur
        print(f"♻️ Reschedule of the current week — overwriting {cur}")
    elif nw and new_week == nw and cur_next:
        target_rel = cur_next
        print(f"♻️ Reschedule of next week — overwriting {cur_next}")
    else:
        target_rel = os.path.join("downloads", att_name).replace("\\", "/")

    new_path = os.path.join(TIMETABLE_DIR, target_rel)
    os.makedirs(os.path.dirname(new_path), exist_ok=True)
    existing_hash = file_hash(new_path)

    with open(new_path, "wb") as f:
        f.write(att_content)
    print(f"💾 Saved: {target_rel}")

    # Safety net: never wire a file that is not a Trim V timetable.
    layout_ok = file_has_timetable_layout(new_path)
    if layout_ok is False and not TIMETABLE_PATTERN.search(os.path.basename(att_name)):
        print(f"⚠️ {att_name!r} does not look like a timetable — not wired in")
        try:
            os.remove(new_path)
        except OSError:
            pass
        return

    new_src = src
    if target_rel == cur:
        new_slot = "TIMETABLE"
    elif target_rel == cur_next:
        new_slot = "TIMETABLE_NEXT"
    else:
        new_slot = "TIMETABLE" if new_week <= this_week else "TIMETABLE_NEXT"

    if new_slot == "TIMETABLE":
        new_src = re.sub(r'^TIMETABLE\s*=\s*"[^"]*"', f'TIMETABLE    = "{target_rel}"', new_src, count=1, flags=re.M)
    else:
        new_src = re.sub(r'^TIMETABLE_NEXT\s*=\s*"[^"]*"', f'TIMETABLE_NEXT = "{target_rel}"', new_src, count=1, flags=re.M)

    src_changed = new_src != src
    if src_changed:
        with open(extract_path, "w") as f:
            f.write(new_src)
        print(f"🔧 Updated extract.py ({new_slot})")

    new_hash = hashlib.sha256(att_content).hexdigest()
    if new_hash == existing_hash and not src_changed:
        print(f"ℹ️ {target_rel} unchanged — no update needed")
        return

    rebuild_and_commit(new_path, sender, received_label, subject)


if __name__ == "__main__":
    main()
