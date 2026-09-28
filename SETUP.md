# Weekly Fortinet keyword positions: setup

Every Tuesday 09:00 IST, GitHub Actions:
1. Looks up last week's (Sun-Sat) position for every row in `Keyword_input.csv`
   (GSC first, then Ahrefs, then Semrush).
2. Adds the week as a new tab (e.g. `Sep 20-26`) to `Fortinet_Keyword_Positions.xlsx`,
   and rebuilds the `Summary` and `Trend` tabs.
3. Commits the workbook back to this repo, so next week builds on it.
4. Emails the workbook to the addresses in `MAIL_TO` (in `.github/workflows/weekly-positions.yml`).

## One-time setup

1. Create a **private** repo on GitHub and push this folder (see the commands in the chat / below).
2. Repo > Settings > Secrets and variables > Actions > New repository secret:

| Secret | Value |
|---|---|
| `GSC_SERVICE_ACCOUNT_JSON` | Full contents of `fortinet-gsc-api-c8ebcebdce98.json` |
| `AHREFS_API_TOKEN` | Ahrefs API key |
| `SEMRUSH_API_KEY` | Semrush API key |
| `SMTP_HOST` | `smtp.gmail.com` |
| `SMTP_PORT` | `587` |
| `SMTP_USER` | the sending mailbox, e.g. kritika@leadwalnut.com |
| `SMTP_PASSWORD` | a Google **App Password** for that mailbox (not the normal password) |
| `MAIL_FROM` | optional, defaults to SMTP_USER |

3. Repo > Settings > Actions > General > Workflow permissions: **Read and write permissions**
   (so the job can commit the workbook).
4. Actions tab > Weekly Fortinet Keyword Positions > Run workflow, tick **dry_run** for a test.

## Everyday use

- Change keywords: edit `Keyword_input.csv` and push. The Date Range column is ignored by the weekly job.
- Backfill a missed week: Actions > Run workflow > `week_start` = the Sunday, e.g. `2026-09-13`.
- Change recipients: edit `MAIL_TO` in the workflow file.
- Run locally: `python send_weekly_report.py --dry-run` (uses `keys.env`).
