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
from sheets import append_to_excel, format_summary
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
        if GDRIVE_FOLDER_ID:
            try:
                download_from_drive(xlsx_name, GDRIVE_FOLDER_ID, xlsx_path)
            except Exception as e:
                log.warning(f"Could not download {xlsx_name} from Drive: {e}")

        xlsx_path = append_to_excel(data, target_date, tmpdir)

        if GDRIVE_FOLDER_ID:
            upload_to_drive(json_path, GDRIVE_FOLDER_ID)
            upload_to_drive(xlsx_path, GDRIVE_FOLDER_ID)
            log.info("Uploaded to Google Drive")

    send_message(format_summary(data, target_date))
    log.info("Telegram notification sent")


if __name__ == "__main__":
    main()
