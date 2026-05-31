import argparse
import json
import logging
import os
import re
import tempfile
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

from dotenv import load_dotenv
load_dotenv()

from garmin import init_garmin, collect_daily_data
from gdrive import upload_to_drive, download_from_drive, upload_google_doc, download_google_doc, list_drive_files
from sheets import append_to_excel, append_to_text_doc, format_summary
from telegram import send_message

GDRIVE_FOLDER_ID = os.getenv("GDRIVE_FOLDER_ID", "")
GARMIN_PREFIX = os.getenv("GARMIN_PREFIX", "garmin")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logging.getLogger("garminconnect").setLevel(logging.WARNING)
logging.getLogger("garth").setLevel(logging.WARNING)
log = logging.getLogger(__name__)


def _download_drive_file(name: str, folder_id: str, local_path: Path) -> bool:
    """Download a file from Drive. Returns True if safe to upload back."""
    try:
        download_from_drive(name, folder_id, local_path)
        return True
    except Exception as e:
        log.error(f"Failed to download {name} from Drive: {e}")
        return False


def existing_json_dates() -> set[date]:
    """Dates already uploaded to Drive, parsed from daily json filenames."""
    if not GDRIVE_FOLDER_ID:
        return set()
    try:
        names = list_drive_files(GDRIVE_FOLDER_ID, name_contains=f"{GARMIN_PREFIX}-")
    except Exception as e:
        log.error(f"Failed to list Drive files: {e}")
        return set()
    pattern = re.compile(rf"^{re.escape(GARMIN_PREFIX)}-(\d{{4}}-\d{{2}}-\d{{2}})\.json$")
    dates: set[date] = set()
    for name in names:
        m = pattern.match(name)
        if m:
            try:
                dates.add(date.fromisoformat(m.group(1)))
            except ValueError:
                pass
    return dates


def _daterange(start: date, end: date) -> list[date]:
    days = []
    d = start
    while d <= end:
        days.append(d)
        d += timedelta(days=1)
    return days


def process_day(client, target_date: date, tmpdir: Path, notify: bool = True):
    """Collect data for a single day, update Drive files."""
    data = collect_daily_data(client, target_date)
    d = target_date.isoformat()

    json_path = tmpdir / f"{GARMIN_PREFIX}-{d}.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, default=str, ensure_ascii=False)

    xlsx_name = f"{GARMIN_PREFIX}-{target_date.year}.xlsx"
    xlsx_path = tmpdir / xlsx_name
    upload_xlsx = True
    if GDRIVE_FOLDER_ID:
        upload_xlsx = _download_drive_file(xlsx_name, GDRIVE_FOLDER_ID, xlsx_path)
    xlsx_path = append_to_excel(data, target_date, tmpdir, xlsx_name)

    doc_name = f"{GARMIN_PREFIX}-{target_date.year}"
    txt_path = tmpdir / f"{doc_name}.txt"
    upload_doc = True
    if GDRIVE_FOLDER_ID:
        try:
            download_google_doc(doc_name, GDRIVE_FOLDER_ID, txt_path)
        except Exception as e:
            log.error(f"Failed to download Google Doc '{doc_name}': {e}")
            upload_doc = False
    txt_path = append_to_text_doc(data, target_date, txt_path)

    if GDRIVE_FOLDER_ID:
        upload_to_drive(json_path, GDRIVE_FOLDER_ID)
        log.info(f"[{d}] Uploaded json to Google Drive")
        if upload_xlsx:
            upload_to_drive(xlsx_path, GDRIVE_FOLDER_ID)
            log.info(f"[{d}] Uploaded xlsx to Google Drive")
        else:
            log.error(f"[{d}] Skipping xlsx upload to prevent data loss")
        if upload_doc:
            upload_google_doc(txt_path, doc_name, GDRIVE_FOLDER_ID)
            log.info(f"[{d}] Uploaded Google Doc to Drive")
        else:
            log.error(f"[{d}] Skipping Google Doc upload to prevent data loss")

    if notify:
        send_message(format_summary(data, target_date))
        log.info(f"[{d}] Telegram notification sent")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--since", type=str, help="Backfill from this date (YYYY-MM-DD) to yesterday")
    args = parser.parse_args()

    yesterday = (datetime.now(timezone.utc) - timedelta(days=1)).date()

    if args.since:
        start = date.fromisoformat(args.since)
        days = _daterange(start, yesterday)
        log.info(f"Backfilling {len(days)} day(s) from {start} to {yesterday}")
    else:
        existing = existing_json_dates()
        if existing:
            # Fill every gap from the earliest known day through yesterday, so a
            # stray newer file can't mask missing days in between.
            floor = min(existing)
            days = [d for d in _daterange(floor, yesterday) if d not in existing]
            log.info(
                f"Drive has {len(existing)} day(s) from {floor} to {max(existing)}; "
                f"{len(days)} day(s) missing through {yesterday}"
            )
        else:
            days = [yesterday]
            log.info("No prior data on Drive, processing yesterday only")

    if not days:
        log.info("Already up to date, nothing to do")
        return

    client = init_garmin()
    log.info(f"Processing {len(days)} day(s)")
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        for day in days:
            process_day(client, day, tmpdir, notify=(day == yesterday))


if __name__ == "__main__":
    main()
