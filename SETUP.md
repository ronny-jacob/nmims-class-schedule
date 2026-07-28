# Automated Timetable Updates

This setup connects your Outlook mailbox to GitHub Actions so the site
auto-updates whenever a new timetable is emailed to you.

## What you need

- A Microsoft 365 school/work account with Outlook
- Ability to register an Azure AD app (or a personal Microsoft account to create one)

## Step 1 — Register an Azure AD app

1. Go to https://entra.microsoft.com → Applications → App registrations
2. Click **New registration**
   - Name: `Timetable Auto-Updater`
   - Supported account types: **Accounts in any organizational directory** (or personal)
   - Redirect URI: (leave blank)
3. After creation, note the **Application (client) ID** and **Directory (tenant) ID**
4. Under **Certificates & secrets** → **Client secrets** → New client secret
   - Copy the secret value immediately

5. Under **API permissions** → Add a permission → Microsoft Graph → **Delegated permissions**
   - Add `Mail.Read`
   - Add `User.Read`
   - Click **Grant admin consent** (if you're an admin) or skip if you'll use device code flow

## Step 2 — Get a refresh token

Install the MSAL library and run the device code flow:

```bash
pip install msal
python3 -c "
import msal, json, os
app = msal.PublicClientApplication(
    'YOUR_CLIENT_ID',
    authority='https://login.microsoftonline.com/YOUR_TENANT_ID'
)
flow = app.initiate_device_flow(scopes=['Mail.Read', 'User.Read', 'offline_access'])
print(flow['message'])
result = app.acquire_token_by_device_flow(flow)
if 'refresh_token' in result:
    print('REFRESH_TOKEN:', result['refresh_token'])
else:
    print('Error:', result.get('error_description'))
"
```

A browser will open — log in with your Outlook account and approve the permissions.
Copy the refresh token that prints.

## Step 3 — Add GitHub secrets

Go to your repo → Settings → Secrets and variables → Actions → New repository secret:

| Secret | Value |
|--------|-------|
| `MS_AUTHORITY` | `https://login.microsoftonline.com/YOUR_TENANT_ID` |
| `MS_CLIENT_ID` | Your Azure AD app client ID |
| `MS_REFRESH_TOKEN` | The refresh token from Step 2 |
| `MS_USER_EMAIL` | Your email address (e.g. `you@nmims.in`) |

## Step 4 — Enable the workflow

Push the code. The GitHub Action runs automatically every hour
(6 AM to 11 PM IST, weekdays only). You can also trigger it manually
from the Actions tab.

## How it works

1. GitHub Action runs on cron (hourly)
2. `check_mail.py` authenticates via Microsoft Graph API
3. Searches inbox for the most recent .xlsx attachment
4. Compares file hash — if unchanged, does nothing
5. If new: saves the file, runs extract.py, commits & pushes
6. Site shows "Updated from [sender] on [date] at [time] IST"

## Troubleshooting

- **Refresh token expired**: Refresh tokens expire after 90 days if unused.
  Run Step 2 again to get a new one.
- **No emails found**: Check that timetable emails have .xlsx attachments
  with dates in the filename (e.g. `27.07.2026 to 02.08.2026.xlsx`)
- **Action failing**: Check GitHub Actions logs for error details
