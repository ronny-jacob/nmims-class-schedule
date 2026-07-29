# Automated Timetable Updates

This setup checks your Outlook inbox via IMAP every hour and
auto-updates the site whenever a new timetable is emailed to you.

## Step 1 — Enable IMAP on your Outlook account

1. Go to https://outlook.office365.com → ⚙️ Settings → **Mail**
2. Click **Forwarding and IMAP**
3. Under **IMAP access**, select **Yes**
4. Click **Save**

## Step 2 — Create an app password

1. Go to https://account.microsoft.com/security → **Advanced security options**
2. Scroll down to **App passwords** and click **Create a new app password**
3. A 16-character password will appear — copy it now (you won't see it again)
4. Use this app password below, NOT your regular Microsoft password

## Step 3 — Add GitHub secrets

Go to your repo → Settings → Secrets and variables → Actions → **New repository secret**:

| Secret | Value |
|--------|-------|
| `IMAP_USER` | Your full email address (e.g. `you@nmims.in`) |
| `IMAP_PASS` | The 16-character app password from Step 2 |
| `IMAP_LOOKBACK_DAYS` | `7` (how many days back to search for timetable emails) |

## Step 4 — Enable the workflow

The workflow file `.github/workflows/check-timetable.yml` is already in the repo.
Once the secrets are added, GitHub Actions will run it automatically every hour
(6 AM to 11 PM IST, weekdays). You can also trigger it manually from the Actions tab.

## How it works

1. GitHub Action runs hourly via cron
2. `check_mail.py` connects to Outlook via IMAP with your app password
3. Searches inbox for the most recent email with a `.xlsx` attachment
4. Looks for timetable filenames (containing a date range like `27.07.2026 to 02.08.2026`)
5. Compares file hash — if unchanged, does nothing
6. If new: saves file, runs `extract.py`, commits & pushes to GitHub
7. Site shows "Updated from [sender] on [date] at [time] IST"

## Testing

Run the script manually to verify everything works:

```bash
IMAP_USER="you@nmims.in" IMAP_PASS="16-char-app-password" python3 check_mail.py
```
