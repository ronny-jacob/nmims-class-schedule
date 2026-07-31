# Automated Timetable Updates

Your Outlook account forwards timetable emails to Gmail. A GitHub Action
checks that Gmail inbox via IMAP every hour and auto-updates the site
whenever a new timetable arrives.

## Step 1 — Enable IMAP on Gmail

1. Open https://mail.google.com → ⚙️ **Settings** → **See all settings**
2. Go to the **Forwarding and POP/IMAP** tab
3. Under **IMAP access**, select **Enable IMAP** → **Save Changes**

## Step 2 — Enable 2-Step Verification

Google only allows "app passwords" when 2-Step Verification is on:

1. Go to https://myaccount.google.com/security
2. Click **2-Step Verification** and turn it on (follow the prompts)

## Step 3 — Create an app password

1. Go to https://myaccount.google.com/apppasswords
2. Create an app called e.g. `timetable-bot`
3. A 16-character password appears (e.g. `abcd efgh ijkl mnop`) — copy it now
4. Use this as `IMAP_PASS`, NOT your regular Gmail password

## Step 4 — Add GitHub secrets

Go to your repo → Settings → Secrets and variables → Actions → **New repository secret**:

| Secret | Value |
|--------|-------|
| `IMAP_USER` | Your Gmail address (e.g. `you@gmail.com`) |
| `IMAP_PASS` | The 16-character app password from Step 3 (spaces optional) |
| `IMAP_LOOKBACK_DAYS` | `7` (how many days back to search for timetable emails) |

## Step 5 — Enable the workflow

The workflow file `.github/workflows/check-timetable.yml` is already in the repo.
Once the secrets are added, GitHub Actions will run it automatically every hour
(6 AM to 11 PM IST, weekdays). You can also trigger it manually from the Actions tab.

## How it works

1. GitHub Action runs hourly via cron
2. `check_mail.py` connects to Gmail (`imap.gmail.com`) via IMAP with the app password
3. Searches inbox for the most recent email with a `.xlsx` attachment
4. Looks for timetable filenames (containing a date range like `27.07.2026 to 02.08.2026`)
5. Compares file hash — if unchanged, does nothing
6. If new: saves file, runs `extract.py`, commits & pushes to GitHub
7. Site shows "Updated from [sender] on [date] at [time] IST"

## Testing

Run the script manually to verify everything works:

```bash
IMAP_USER="you@gmail.com" IMAP_PASS="16-char-app-password" python3 check_mail.py
```
