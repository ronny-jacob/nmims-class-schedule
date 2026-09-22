# GitHub flow

Two repos, three triggers, one rule: **production (`nmims-class-schedule`) is the live site; staging (`nmims-class-schedule-staging`) is where manual edits live.**

## Repos

| Repo | Branch | URL | Purpose |
|---|---|---|---|
| `nmims-class-schedule` (prod) | `main` | https://ronny-jacob.github.io/nmims-class-schedule/ | Live site (Pages). Auto-updated by the timetable bot. |
| `nmims-class-schedule-staging` (staging) | `main` | https://ronny-jacob.github.io/nmims-class-schedule-staging/ | Preview. Use for hand-edits (CSS, copy, new features, placement announcements) before they ship. |

## Workflows

| File | Repo it lives on | Trigger | What it does |
|---|---|---|---|
| `check-timetable.yml` | prod | cron (hourly, 8 AM – 11 PM IST) + every-minute Apps Script on Saturday | IMAP-fetch the latest timetable, regenerate `data.json` + `index.html`, commit, push, then mirror prod → staging. |
| `deploy-pages.yml` | prod | push to main | Bump SW cache version, deploy to GitHub Pages. |
| `promote-to-production.yml` | staging | `workflow_dispatch` | Merge prod main into staging main, push staging main → prod main. Refuses to run if prod is ahead. |
| `deploy-apps-script.yml` | prod | push to `google-apps-script/**` | Push Apps Script code to the analytics project. |
| `deploy-mail-watch.yml` | prod | push to `mail-watch/**` | Push the every-minute trigger script. |

## Normal flows

### Automatic: timetable arrives

```
sadashiv@nmims.edu → Outlook → Gmail (IMAP) → check-timetable.yml
                                       ↓
                  rebuild data.json + index.html (extract.py)
                                       ↓
            commit + push to prod main → deploy-pages.yml → Pages (live)
                                       ↓
            mirror prod main → staging main
```

### Manual: edit placements.json or any other file

1. Edit on **staging**. Commit, push to staging main.
2. Open Actions on **staging** → run "Promote to Production".
3. The workflow merges prod into staging first, then pushes staging main → prod main, which triggers `deploy-pages.yml` on prod.

### Manual: edit directly on prod (e.g. urgent placement)

Editing prod directly works because the bot does `git pull`-equivalent (rebase). But this is **discouraged**: every bot run will try to mirror to staging, and a divergence from staging will block the mirror until you resolve it. **Always edit on staging and promote.**

## When staging diverges

A diverged staging branch means staging has commits that prod doesn't, AND prod has commits that staging doesn't (typically the bot's latest timetable push). The mirror step in `check-timetable.yml` will fail with a non-fast-forward error and surface `::warning::` in the Actions log.

Two ways to resolve:

1. **Promote staging → prod.** This is the standard resolution. Run the "Promote to Production" workflow on staging. After promotion, staging and prod share a history and the next bot mirror succeeds.
2. **Reset staging to prod.** On a clone of staging, `git fetch origin && git reset --hard origin/main`. Use this only if you genuinely don't want the unpromoted staging edits.

The bot will keep pushing to prod and logging mirror failures until you resolve.

## Secrets

| Secret | Repo | Used for |
|---|---|---|
| `STUDENT_LIST_B64`, `LAST_YEAR_LIST_B64` | prod | Roster xlsx (decoded at CI run-time, never committed). |
| `IMAP_USER`, `IMAP_PASS`, `IMAP_LOOKBACK_DAYS` | prod | Gmail IMAP for timetable bot. |
| `STAGING_PUSH_TOKEN` | prod | Bot's push token for the prod → staging mirror. |
| `PROD_PUSH_TOKEN` | staging | Promote workflow's push token for staging → prod. |
| `CLASP_CREDS`, `CLASP_SCRIPT_ID`, `CLASP_DEPLOYMENT_ID` | prod | Apps Script (analytics) deployment. |
| `MAIL_WATCH_CLASP_*` | prod | Apps Script (mail watch) deployment. |

## Branch protection (recommended)

For both repos, set the following on `main`:

- ✅ Require a pull request before merging (only for staging — prod is auto-pushed by the bot)
- ✅ Require status checks to pass before merging
- ✅ Require linear history (no merge commits to prod; merge commits to staging are expected from the Promote workflow)
- ❌ Do **not** require PR reviews on prod (the bot pushes directly). The bot uses `contents: write` and bypasses PRs by design.
