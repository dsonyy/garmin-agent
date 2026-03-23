import json
import logging
import os
import tempfile
from datetime import datetime, timedelta, timezone
from pathlib import Path

from dotenv import load_dotenv
load_dotenv()

from garmin import init_garmin, collect_daily_data
from gdrive import upload_to_drive, download_from_drive
from sheets import append_to_excel, append_to_text_doc, format_summary
from telegram import send_message

GDRIVE_FOLDER_ID = os.getenv("GDRIVE_FOLDER_ID", "")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logging.getLogger("garminconnect").setLevel(logging.WARNING)
logging.getLogger("garth").setLevel(logging.WARNING)
log = logging.getLogger(__name__)


def main():
    target_date = (datetime.now(timezone.utc) - timedelta(days=1)).date()

    client = init_garmin()
    data = collect_daily_data(client, target_date)

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        d = target_date.isoformat()

        json_path = tmpdir / f"{d}-garmin-raw.json"
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, default=str, ensure_ascii=False)

        xlsx_name = f"{target_date.year}-garmin.xlsx"
        xlsx_path = tmpdir / xlsx_name
        upload_xlsx = True
        if GDRIVE_FOLDER_ID:
            try:
                download_from_drive(xlsx_name, GDRIVE_FOLDER_ID, xlsx_path)
            except Exception as e:
                log.error(f"Failed to download {xlsx_name} from Drive: {e}")
                upload_xlsx = False

        xlsx_path = append_to_excel(data, target_date, tmpdir)

        txt_name = f"{target_date.year}-garmin.txt"
        txt_path = tmpdir / txt_name
        upload_txt = True
        if GDRIVE_FOLDER_ID:
            try:
                download_from_drive(txt_name, GDRIVE_FOLDER_ID, txt_path)
            except Exception as e:
                log.error(f"Failed to download {txt_name} from Drive: {e}")
                upload_txt = False

        txt_path = append_to_text_doc(data, target_date, txt_path)

        if GDRIVE_FOLDER_ID:
            upload_to_drive(json_path, GDRIVE_FOLDER_ID)
            log.info("Uploaded json to Google Drive")
            if upload_xlsx:
                upload_to_drive(xlsx_path, GDRIVE_FOLDER_ID)
                log.info("Uploaded xlsx to Google Drive")
            else:
                log.error("Skipping xlsx upload to prevent data loss")
            if upload_txt:
                upload_to_drive(txt_path, GDRIVE_FOLDER_ID)
                log.info("Uploaded txt to Google Drive")
            else:
                log.error("Skipping txt upload to prevent data loss")

    send_message(format_summary(data, target_date))
    log.info("Telegram notification sent")


if __name__ == "__main__":
    main()
