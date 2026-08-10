import imaplib, email, json, os, re, subprocess, sys, hashlib, tempfile
from datetime import datetime, timezone, timedelta
from email.header import decode_header

# ─── Config ───────────────────────────────────────────────────────────
# Outlook forwards timetable emails to Gmail, so we poll Gmail via IMAP.
IMAP_SERVER   = os.getenv("IMAP_SERVER", "imap.gmail.com")
IMAP_USER     = os.getenv("IMAP_USER", "")
IMAP_PASS     = os.getenv("IMAP_PASS", "")
LOOKBACK_DAYS = int(os.getenv("IMAP_LOOKBACK_DAYS", "7"))
TIMETABLE_DIR = os.path.dirname(os.path.abspath(__file__))
TIMETABLE_PATTERN = re.compile(r'\d{1,2}\.\d{1,2}\.\d{4}\s*to\s*\d{1,2}\.\d{1,2}\.\d{4}')

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
    """Connect via IMAP and find the most recent email with a timetable .xlsx."""
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

    best = None
    for uid in msg_ids:
        status, data = mail.fetch(uid, "(RFC822)")
        if status != "OK":
            continue
        raw_email = data[0][1]
        msg = email.message_from_bytes(raw_email)

        subject = decode_str(msg.get("Subject"))
        from_addr = decode_str(msg.get("From"))
        date_str = decode_str(msg.get("Date"))

        if not subject:
            continue

        has_xlsx = False
        att_name = None
        att_content = None

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
                has_xlsx = True
                att_name = fn_decoded
                att_content = part.get_payload(decode=True)
                break

        if not has_xlsx or att_content is None:
            continue

        # Prefer messages where filename matches timetable pattern
        is_timetable = bool(TIMETABLE_PATTERN.search(att_name or ""))
        is_timetable_subj = bool(re.search(r'timetable|schedule|time\s*table', subject, re.I))

        if not is_timetable and not is_timetable_subj:
            continue

        # Parse date for display
        try:
            parsed_dt = email.utils.parsedate_to_datetime(date_str)
            ist = parsed_dt.astimezone(timezone(timedelta(hours=5, minutes=30)))
            received_label = ist.strftime("%I:%M %p IST on %d %b %Y").lstrip("0").replace(" 0", " ")
        except Exception:
            received_label = date_str

        # Save best match (prefer timetable-patterned filenames)
        if best is None or (is_timetable and not best[0]):
            best = (is_timetable, from_addr, received_label, subject, att_name, att_content)
            print(f"DEBUG picked: {received_label} | {subject[:60]} | {att_name} | sha={hashlib.sha256(att_content).hexdigest()[:12]}")

    mail.logout()
    return best


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

    result = subprocess.run(
        [sys.executable, "extract.py"],
        capture_output=True, text=True,
        cwd=TIMETABLE_DIR,
    )
    print(result.stdout)
    if result.returncode != 0:
        print(f"❌ extract.py failed: {result.stderr}")
        return False

    last_updated = {
        "sender": sender,
        "received_at": received_at,
        "subject": subject,
        "updated_at": datetime.now(timezone.utc).isoformat(),
    }
    meta_path = os.path.join(TIMETABLE_DIR, "last_updated.json")
    with open(meta_path, "w") as f:
        json.dump(last_updated, f, indent=2)

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


def parse_file_dates(name):
    """Extract the start date (dd.mm.yyyy) from a timetable filename."""
    m = TIMETABLE_PATTERN.search(name or "")
    if not m:
        return None
    start_str = m.group(0).split("to", 1)[0].strip()
    try:
        return datetime.strptime(start_str, "%d.%m.%Y").date()
    except ValueError:
        return None


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

    is_timetable, sender, received_label, subject, att_name, att_content = result

    # Build path — save in repo's downloads/ directory
    os.makedirs(os.path.join(TIMETABLE_DIR, "downloads"), exist_ok=True)
    new_path = os.path.join(TIMETABLE_DIR, "downloads", att_name)
    existing_hash = file_hash(new_path)

    with open(new_path, "wb") as f:
        f.write(att_content)
    print(f"💾 Saved: {new_path}")
    new_rel = os.path.relpath(new_path, TIMETABLE_DIR)

    # Decide slot based on week: current week -> TIMETABLE, future -> TIMETABLE_NEXT
    today = datetime.now(timezone.utc).date()
    this_week = monday_of(today)
    new_week = week_of_path(new_path)

    extract_path, src, cur, cur_next = read_extract()
    new_src = src
    new_slot = None
    promoted = False

    if new_week:
        if new_week < this_week:
            print(f"ℹ️ {att_name} is for a past week — saved but not wired in")
            return
        # Promote TIMETABLE_NEXT to TIMETABLE when its week becomes current/latest
        nw = week_of_path(cur_next)
        cw = week_of_path(cur)
        if nw and nw <= this_week and (cw is None or nw > cw):
            new_src = re.sub(r'^TIMETABLE\s*=\s*"[^"]*"', f'TIMETABLE    = "{cur_next}"', new_src, count=1, flags=re.M)
            new_src = re.sub(r'^TIMETABLE_NEXT\s*=\s*"[^"]*"', 'TIMETABLE_NEXT = ""', new_src, count=1, flags=re.M)
            promoted = True
            print(f"♻️ Promoted {cur_next} to current timetable")
        new_slot = "TIMETABLE_NEXT" if new_week > this_week else "TIMETABLE"
    else:
        new_slot = "TIMETABLE"

    if new_slot == "TIMETABLE":
        new_src = re.sub(r'^TIMETABLE\s*=\s*"[^"]*"', f'TIMETABLE    = "{new_rel}"', new_src, count=1, flags=re.M)
    elif new_slot == "TIMETABLE_NEXT":
        new_src = re.sub(r'^TIMETABLE_NEXT\s*=\s*"[^"]*"', f'TIMETABLE_NEXT = "{new_rel}"', new_src, count=1, flags=re.M)

    src_changed = new_src != src
    if src_changed:
        with open(extract_path, "w") as f:
            f.write(new_src)
        print(f"🔧 Updated extract.py ({new_slot or 'promotion'})")

    existing_hash = file_hash(new_path)
    new_hash = hashlib.sha256(att_content).hexdigest()
    if new_hash == existing_hash and not src_changed:
        print(f"ℹ️ {att_name} unchanged — no update needed")
        return

    rebuild_and_commit(new_path, sender, received_label, subject)


if __name__ == "__main__":
    main()
