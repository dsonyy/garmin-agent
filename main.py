import argparse
import json
import logging
import os
import tempfile
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

from dotenv import load_dotenv
load_dotenv()

from garmin import init_garmin, collect_daily_data
from gdrive import upload_to_drive, download_from_drive, upload_google_doc, download_google_doc
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
    client = init_garmin()

    if args.since:
        start = date.fromisoformat(args.since)
        days = []
        d = start
        while d <= yesterday:
            days.append(d)
            d += timedelta(days=1)
        log.info(f"Backfilling {len(days)} days from {start} to {yesterday}")
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            for day in days:
                process_day(client, day, tmpdir, notify=False)
    else:
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)
            process_day(client, yesterday, tmpdir, notify=True)


if __name__ == "__main__":
    main()
