# garmin-agent

Daily export of Garmin Connect health data to Google Drive, with a Telegram summary.

Each run collects one day of metrics (steps, heart rate, sleep, stress, body battery,
HRV, training readiness, activities, and more) and writes three artifacts to a Drive folder:

- `<prefix>-<date>.json` — full raw export for that day
- `<prefix>-<year>.xlsx` — one row per day, appended/updated, including sport-activity
  columns (count, sports, totals, and detail for the day's longest activity)
- `<prefix>-<year>` — native Google Doc, human-readable, one block per day

On a normal run it also sends a Telegram message summarizing the most recent day.

## How it picks the day

With no arguments, the agent lists the existing `<prefix>-<date>.json` files in the Drive
folder, finds the latest date, and processes every day from there through yesterday (UTC).
If nothing missed, it does only yesterday. If no prior data exists, it does yesterday only.
This makes daily runs self-healing: if the machine was off for a week, the next run backfills
the gap automatically.

Telegram notification is sent only for yesterday, not for backfilled days.

## Requirements

- Python 3.12+ and [uv](https://docs.astral.sh/uv/)
- A Garmin Connect account
- A Google Cloud OAuth client (Desktop) with the Drive API enabled
- A Telegram bot token and chat id

## Setup

Install dependencies into `.venv`:

```
uv sync
```

Create a `.env` file:

```
TELEGRAM_BOT_TOKEN="..."
TELEGRAM_CHAT_ID="..."

GDRIVE_CLIENT_SECRET_FILE="/path/to/oauth_client.json"
GDRIVE_FOLDER_ID="..."

GARMIN_EMAIL="you@example.com"
GARMIN_PASSWORD="..."
GARMIN_TOKEN_DIR="./secrets"
GARMIN_PREFIX="garmin"
```

## First run authorization

Two credentials are established on first use and then cached:

- Garmin: logs in with email and password, saves tokens to `GARMIN_TOKEN_DIR`. Later runs
  reuse the tokens. If Garmin requires MFA, run once interactively to complete it.
- Google Drive: opens a browser for OAuth, saves the token to `secrets/gdrive_token.json`.
  Because the scope is `drive.file`, the app only sees files it created itself.

Run once by hand to complete both:

```
uv run main.py
```

## Usage

Normal run (auto-resume from last processed day, notify for yesterday):

```
uv run main.py
```

Backfill from an explicit start date through yesterday (no notifications):

```
uv run main.py --since 2026-01-01
```

## Cron

Run daily at 00:30. The committed `crontab` file targets the VPS path `/opt/garmin-agent`.
For a different install, point the line at your own path:

See `crontab` file.

Install it for the current user:

```
crontab crontab
```

The target day is computed in UTC, so the schedule is timezone independent for which day
gets exported.

## Files

- `main.py` — entry point, day selection, orchestration
- `garmin.py` — Garmin Connect auth and data collection
- `gdrive.py` — Google Drive upload, download, listing
- `sheets.py` — row extraction, Excel and text-doc writing, Telegram summary formatting
- `telegram.py` — Telegram Bot API client
- `backfill.py` — standalone local-only backfill helper (writes to a local output dir)
