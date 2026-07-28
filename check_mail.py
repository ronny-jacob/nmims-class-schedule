import json, os, re, subprocess, sys, tempfile, hashlib, base64
from datetime import datetime, timezone, timedelta

import requests

# ─── Config ───────────────────────────────────────────────────────────
AUTHORITY    = os.getenv("MS_AUTHORITY", "https://login.microsoftonline.com/common")
CLIENT_ID    = os.getenv("MS_CLIENT_ID", "")
REFRESH_TOKEN = os.getenv("MS_REFRESH_TOKEN", "")
USER_EMAIL   = os.getenv("MS_USER_EMAIL", "")
TIMETABLE_DIR = os.path.dirname(os.path.abspath(__file__))
TIMETABLE_PATTERN = re.compile(r'\d{1,2}\.\d{1,2}\.\d{4}\s*to\s*\d{1,2}\.\d{1,2}\.\d{4}')

# Graph API scopes for delegated Mail.Read
SCOPE = ["https://graph.microsoft.com/Mail.Read", "https://graph.microsoft.com/User.Read"]
GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def get_access_token():
    """Exchange refresh token for access token using MSAL-compatible OAuth2."""
    tenant = AUTHORITY.rstrip("/").split("/")[-1]
    token_url = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
    data = {
        "client_id": CLIENT_ID,
        "refresh_token": REFRESH_TOKEN,
        "grant_type": "refresh_token",
        "scope": " ".join(SCOPE),
    }
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    resp = requests.post(token_url, data=data, headers=headers, timeout=30)
    if resp.status_code != 200:
        print(f"❌ Token refresh failed: {resp.status_code} {resp.text}")
        sys.exit(1)
    tokens = resp.json()
    return tokens["access_token"], tokens.get("refresh_token", REFRESH_TOKEN)


def search_timetable_messages(token):
    """Fetch the most recent email with an xlsx attachment matching timetable pattern."""
    headers = {
        "Authorization": f"Bearer {token}",
        "Prefer": "outlook.body-content-type=text",
    }
    # Search inbox for messages with .xlsx attachments
    url = f"{GRAPH_BASE}/me/mailFolders/inbox/messages"
    params = {
        "$top": 10,
        "$orderby": "receivedDateTime desc",
        "$filter": "hasAttachments eq true",
    }
    resp = requests.get(url, headers=headers, params=params, timeout=30)
    if resp.status_code != 200:
        print(f"❌ Graph API error: {resp.status_code} {resp.text}")
        return None

    messages = resp.json().get("value", [])
    for msg in messages:
        subject = msg.get("subject", "")
        if not subject:
            continue
        # Check if any attachment matches timetable pattern
        att_url = f"{GRAPH_BASE}/me/messages/{msg['id']}/attachments"
        att_resp = requests.get(att_url, headers=headers, timeout=30)
        if att_resp.status_code != 200:
            continue
        for att in att_resp.json().get("value", []):
            name = att.get("name", "")
            if not name.endswith(".xlsx"):
                continue
            # If filename already contains a date range, accept it
            if TIMETABLE_PATTERN.search(name):
                return msg, att
            # If the subject has "timetable" or "schedule" keywords
            if re.search(r'timetable|schedule|time\s*table', subject, re.I):
                return msg, att

    print("ℹ️ No new timetable email found")
    return None


def download_attachment(token, attachment):
    """Download attachment content bytes."""
    if "@odata.mediaContentType" in attachment:
        # Small attachment (inline)
        raw = attachment.get("contentBytes", "")
        if raw:
            return base64.b64decode(raw)

    # Large attachment — use download URL
    content_url = attachment.get("@microsoft.graph.downloadUrl")
    if content_url:
        resp = requests.get(content_url, timeout=60)
        if resp.status_code == 200:
            return resp.content

    print("⚠️ Could not download attachment")
    return None


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

    # Run extract.py
    result = subprocess.run(
        [sys.executable, "extract.py"],
        capture_output=True, text=True,
        cwd=TIMETABLE_DIR,
    )
    print(result.stdout)
    if result.returncode != 0:
        print(f"❌ extract.py failed: {result.stderr}")
        return False

    # Write last_updated metadata
    last_updated = {
        "sender": sender,
        "received_at": received_at,
        "subject": subject,
        "updated_at": datetime.now(timezone.utc).isoformat(),
    }
    meta_path = os.path.join(TIMETABLE_DIR, "last_updated.json")
    with open(meta_path, "w") as f:
        json.dump(last_updated, f, indent=2)

    # Git commit and push
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
    """Current time in IST as formatted string."""
    utc = datetime.now(timezone.utc)
    ist = utc + timedelta(hours=5, minutes=30)
    return ist.strftime("%I:%M %p IST on %d %b %Y").lstrip("0").replace(" 0", " ")


def main():
    # Validate env
    missing = [v for v in ["MS_CLIENT_ID", "MS_REFRESH_TOKEN"] if not os.getenv(v)]
    if missing:
        print(f"❌ Missing env vars: {', '.join(missing)}")
        print("   See SETUP.md for instructions.")
        sys.exit(1)

    token, new_refresh = get_access_token()

    result = search_timetable_messages(token)
    if not result:
        print(f"ℹ️ No timetable update needed as of {ist_now()}")
        return

    msg, att = result
    sender = msg.get("from", {}).get("emailAddress", {}).get("address", "unknown")
    sender_name = msg.get("from", {}).get("emailAddress", {}).get("name", sender)
    received_raw = msg.get("receivedDateTime", "")
    subject = msg.get("subject", "")

    # Parse received time to IST
    try:
        dt = datetime.fromisoformat(received_raw.replace("Z", "+00:00"))
        received_ist = dt.astimezone(timezone(timedelta(hours=5, minutes=30)))
        received_label = received_ist.strftime("%I:%M %p IST on %d %b %Y").lstrip("0").replace(" 0", " ")
    except Exception:
        received_label = received_raw

    # Download the attachment
    content = download_attachment(token, att)
    if not content:
        print("❌ Failed to download attachment")
        sys.exit(1)

    # Determine filename — use date from attachment name or generate
    att_name = att.get("name", "timetable.xlsx")
    new_path = os.path.join(os.path.dirname(TIMETABLE_DIR), att_name)

    # Check if already have this file (by hash)
    existing_hash = file_hash(new_path)
    new_hash = hashlib.sha256(content).hexdigest()
    if new_hash == existing_hash:
        print(f"ℹ️ {att_name} unchanged — no update needed")
        return

    # Save the new file
    with open(new_path, "wb") as f:
        f.write(content)
    print(f"💾 Saved: {new_path}")

    # Update extract.py paths to point to new file
    extract_path = os.path.join(TIMETABLE_DIR, "extract.py")
    with open(extract_path) as f:
        src = f.read()

    # Replace TIMETABLE path with absolute path to downloaded file
    new_abs = new_path
    src = re.sub(
        r'^TIMETABLE\s*=.*',
        f'TIMETABLE    = "{new_abs}"',
        src,
        count=1,
    )

    # Also handle TIMETABLE_NEXT if the new file is a next-week file
    # For simplicity, we set TIMETABLE to the downloaded file and clear TIMETABLE_NEXT
    src = re.sub(
        r'^TIMETABLE_NEXT\s*=.*',
        'TIMETABLE_NEXT = ""',
        src,
        count=1,
    )

    with open(extract_path, "w") as f:
        f.write(src)

    # Rebuild and deploy
    rebuild_and_commit(new_path, sender_name, received_label, subject)

    # If we got a new refresh token, print it (in workflow, it's re-fetched from secret)
    if new_refresh != REFRESH_TOKEN:
        print("⚠️ Refresh token changed — update MS_REFRESH_TOKEN secret")


if __name__ == "__main__":
    main()
