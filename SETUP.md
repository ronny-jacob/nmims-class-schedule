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
Once the secrets are added, GitHub Actions runs it automatically on this schedule:

- **Monday–Friday:** every hour from **8 AM to 11 PM IST** (no checks 12 AM–8 AM)
- **Saturday:** every hour from 8 AM to 11 PM IST as a safety net (the every-minute watcher below handles real-time)
- **Sunday:** every hour from 8 AM to 11 PM IST as a weekend fail-safe

You can also trigger it manually from the Actions tab.

> The dedicated, free Apps Script watcher (`mail-watch/`) adds an **every-minute**
> check on Saturdays only — see Step 7 below.

## How it works

1. GitHub Action runs on cron (`check-timetable.yml`) — see Step 5 for the schedule
2. `check_mail.py` connects to Gmail (`imap.gmail.com`) via IMAP with the app password
3. Searches inbox for the most recent email with a `.xlsx` attachment
4. Looks for timetable filenames (containing a date range like `27.07.2026 to 02.08.2026`)
5. Compares file hash — if unchanged, does nothing
6. Files dated for the current week go into `TIMETABLE`; future weeks into
   `TIMETABLE_NEXT` (the site shows both, and the next week is promoted to
   current automatically once it arrives)
7. If new: saves file, runs `extract.py`, commits & pushes to GitHub
8. Site shows "Updated from [sender] on [date] at [time] IST"
9. After the push, `deploy-pages.yml` publishes the updated site to GitHub Pages.

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

## Step 7 — Real-time Saturday watcher (dedicated Apps Script, optional)

A small dedicated Apps Script in `mail-watch/` adds an **every-minute** check on
**Saturdays between 8 AM and 11 PM IST**. It does NOT need billing or a GCP
project — it uses only the built-in `GmailApp`. It scans for a new timetable email
and, when it finds one, calls the GitHub API to run `check-timetable.yml` right
away instead of waiting for the next hourly cron. All the heavy lifting
(extract, commit, push, deploy) still happens in GitHub Actions.

It does nothing Mon–Fri (those days rely on the hourly cron) and nothing
between midnight and 8 AM.

### Set it up once

1. Create a standalone Apps Script project (script.google.com → New project).
2. Paste `mail-watch/Code.gs` in as `Code.gs` and set `appsscript.json`.
3. Open **Project Settings**, copy the **Script ID**, and click **Enable
   Google Apps Script API** at https://script.google.com/home/usersettings.
4. Create a **Fine-grained personal access token (PAT)** with **Actions: read
   and write** on the target repo:
   https://github.com/settings/tokens?type=beta
   (For a public repo you can use a classic `repo`+`workflow` PAT.)
5. Open **Project Settings → Script properties** and add:
   | Property | Value |
   |----------|-------|
   | `GITHUB_PAT` | the token from step 4 |
   | `GITHUB_REPO` | e.g. `ronny-jacob/nmims-class-schedule` |
   | `GITHUB_WORKFLOW` | `check-timetable.yml` (optional) |
   | `GITHUB_REF` | `main` (optional) |
6. In the editor, run **`installSaturdayTrigger()`** once and authorize. This
   installs an every-minute time trigger; you can remove it later with
   `uninstallTrigger()`.

### Auto-deploy (optional)

To deploy `mail-watch/` automatically on push (like the analytics script):
1. Get the Script ID and a deployment ID (Deploy → Manage deployments → create
   a new deployment; copy its ID).
2. Add repo secrets:
   ```bash
   gh secret set MAIL_WATCH_CLASP_SCRIPT_ID "<Script ID>"
   gh secret set MAIL_WATCH_CLASP_DEPLOYMENT_ID "<Deployment ID>"
   ```
   (`CLASP_CREDS` is reused — must be the project owner's account.)
3. Push any change to `mail-watch/`, or run the `deploy-mail-watch` workflow from
   the Actions tab.

> Note: the trigger still needs to be installed once (step 6). The workflow only
> deploys the CODE; it does not install the time trigger.

## Testing

Run the script manually to verify everything works:

```bash
IMAP_USER="you@gmail.com" IMAP_PASS="16-char-app-password" python3 check_mail.py
```
