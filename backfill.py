import json
import logging
import os
from datetime import date, timedelta
from pathlib import Path

from dotenv import load_dotenv
load_dotenv()

from garmin import init_garmin, collect_daily_data
from sheets import append_to_excel

OUTPUT_DIR = Path(os.getenv("GARMIN_OUTPUT_DIR", str(Path(__file__).parent / "output")))

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logging.getLogger("garminconnect").setLevel(logging.WARNING)
logging.getLogger("garth").setLevel(logging.WARNING)
log = logging.getLogger(__name__)


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    start = date(2026, 1, 1)
    end = date.today() - timedelta(days=1)

    client = init_garmin()

    current = start
    while current <= end:
        d = current.isoformat()
        json_path = OUTPUT_DIR / f"{d}-garmin-raw.json"

        if json_path.exists():
            log.info(f"Skipping {d}, already exists")
            data = json.loads(json_path.read_text())
            append_to_excel(data, current, OUTPUT_DIR)
            current += timedelta(days=1)
            continue

        try:
            data = collect_daily_data(client, current)

            with open(json_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, default=str, ensure_ascii=False)
            log.info(f"Saved JSON: {json_path}")

            append_to_excel(data, current, OUTPUT_DIR)
            log.info(f"Appended to Excel for {d}")

        except Exception as e:
            log.error(f"Failed {d}: {e}")

        current += timedelta(days=1)

    log.info("Backfill complete")


if __name__ == "__main__":
    main()
