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
| `STUDENT_LIST_B64` | Base64 of `sources/Division wise List- Trimester IV.xlsx` |
| `LAST_YEAR_LIST_B64` | Base64 of `sources/First Year Division list.xlsx` |

> The roster spreadsheets (which contain students' SAP IDs) are **not committed**
> to this public repo. Instead they're stored as base64 secrets and decoded on the
> runner before each build. Non-sensitive inputs (`sources/*.xlsx` timetables and
> food menu) are committed directly.

Set the two base64 secrets from the terminal:

```bash
base64 -i "/path/to/Division wise List- Trimester IV.xlsx" | tr -d '\n' | gh secret set STUDENT_LIST_B64
base64 -i "/path/to/First Year Division list.xlsx" | tr -d '\n' | gh secret set LAST_YEAR_LIST_B64
```

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
6. Files dated for the current week go into `TIMETABLE`; future weeks into
   `TIMETABLE_NEXT` (the site shows both, and the next week is promoted to
   current automatically once it arrives)
7. If new: saves file, runs `extract.py`, commits & pushes to GitHub
8. Site shows "Updated from [sender] on [date] at [time] IST"

## Step 6 — Auto-deploy the Analytics Apps Script (optional)

The analytics web app (logging button clicks / page views to Google Sheets) is
kept in `google-apps-script/Code.gs`. Instead of pasting & redeploying manually,
a GitHub Action deploys it automatically on every push to that folder.

1. Authorize `clasp` **as the account that owns the analytics project** (found in
   the Apps Script editor's Project Settings):
   ```bash
   npm i -g @google/clasp && clasp login
   ```
2. Enable the Apps Script API for that account at
   https://script.google.com/home/usersettings
3. Add the secrets (from your terminal):
   ```bash
   base64 -i ~/.clasprc.json | tr -d '\n' | gh secret set CLASP_CREDS
   gh secret set CLASP_SCRIPT_ID "<Script ID from Project Settings>"
   gh secret set CLASP_DEPLOYMENT_ID "<deployment ID from the /exec URL>"
   ```
4. Push any change to `google-apps-script/`, or run the `deploy-apps-script`
   workflow manually from the Actions tab. A new version is deployed to the same URL.

## Testing

Run the script manually to verify everything works:

```bash
IMAP_USER="you@gmail.com" IMAP_PASS="16-char-app-password" python3 check_mail.py
```
