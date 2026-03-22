import json
import logging
import os
from datetime import date, timedelta
from pathlib import Path

from dotenv import load_dotenv
load_dotenv()

from garmin import init_garmin, collect_daily_data
from gdrive import upload_to_drive
from sheets import append_to_excel
from telegram import send_message

OUTPUT_DIR = Path(os.getenv("GARMIN_OUTPUT_DIR", str(Path(__file__).parent / "output")))
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
    target_date = date.today() - timedelta(days=1)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    client = init_garmin()
    data = collect_daily_data(client, target_date)

    d = target_date.isoformat()

    json_path = OUTPUT_DIR / f"{d}-garmin-raw.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, default=str, ensure_ascii=False)
    log.info(f"Saved JSON: {json_path}")

    xlsx_path = append_to_excel(data, target_date, OUTPUT_DIR)
    log.info(f"Saved Excel: {xlsx_path}")

    if GDRIVE_FOLDER_ID:
        try:
            upload_to_drive(json_path, GDRIVE_FOLDER_ID)
            upload_to_drive(xlsx_path, GDRIVE_FOLDER_ID)
            log.info("Uploaded to Google Drive")
        except Exception as e:
            log.error(f"Google Drive upload failed: {e}")

    send_message(f"Garmin report ready for {d}")
    log.info("Telegram notification sent")


if __name__ == "__main__":
    main()
