import logging
import os
import time
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

from garminconnect import Garmin
from garth.exc import GarthHTTPError

GARMIN_EMAIL = os.getenv("GARMIN_EMAIL", "")
GARMIN_PASSWORD = os.getenv("GARMIN_PASSWORD", "")
TOKEN_DIR = os.getenv("GARMIN_TOKEN_DIR", str(Path.home() / ".garminconnect"))

log = logging.getLogger(__name__)


def init_garmin() -> Garmin:
    token_path = Path(TOKEN_DIR).expanduser()

    if token_path.exists() and list(token_path.glob("*.json")):
        try:
            client = Garmin()
            client.login(str(token_path))
            log.info("Authenticated via saved tokens")
            return client
        except Exception as e:
            log.warning(f"Token login failed: {e}, falling back to credentials")

    if not GARMIN_EMAIL or not GARMIN_PASSWORD:
        raise RuntimeError(
            "No valid tokens and GARMIN_EMAIL/GARMIN_PASSWORD not set."
        )

    client = Garmin(email=GARMIN_EMAIL, password=GARMIN_PASSWORD, is_cn=False)
    client.login()
    client.garth.dump(str(token_path))
    log.info("Authenticated via credentials, tokens saved")
    return client


def _safe_call(client: Garmin, method_name: str, *args, **kwargs) -> Any | None:
    method = getattr(client, method_name, None)
    if method is None:
        log.warning(f"Method {method_name} not found on client")
        return None
    try:
        return method(*args, **kwargs)
    except GarthHTTPError as e:
        status = getattr(getattr(e, "response", None), "status_code", None)
        if status in (400, 404):
            log.debug(f"{method_name}: not available ({status})")
        elif status == 429:
            log.warning(f"{method_name}: rate limited, sleeping 60s")
            time.sleep(60)
            try:
                return method(*args, **kwargs)
            except Exception:
                pass
        else:
            log.warning(f"{method_name} failed: {e}")
        return None
    except Exception as e:
        log.warning(f"{method_name} failed: {e}")
        return None


def collect_daily_data(client: Garmin, target_date: date) -> dict:
    d = target_date.isoformat()
    week_ago = (target_date - timedelta(days=7)).isoformat()

    log.info(f"Collecting data for {d}")

    data: dict[str, Any] = {
        "_meta": {
            "export_date": d,
            "export_timestamp": datetime.utcnow().isoformat() + "Z",
            "garmin_device": "Forerunner 965",
        }
    }

    all_calls = {
        "stats": ("get_stats", [d]),
        "user_summary": ("get_user_summary", [d]),
        "steps": ("get_steps_data", [d]),
        "heart_rates": ("get_heart_rates", [d]),
        "sleep": ("get_sleep_data", [d]),
        "stress_all_day": ("get_all_day_stress", [d]),
        "lifestyle_logging": ("get_lifestyle_logging_data", [d]),
        "training_readiness": ("get_training_readiness", [d]),
        "training_status": ("get_training_status", [d]),
        "respiration": ("get_respiration_data", [d]),
        "spo2": ("get_spo2_data", [d]),
        "max_metrics": ("get_max_metrics", [d]),
        "hrv": ("get_hrv_data", [d]),
        "fitness_age": ("get_fitnessage_data", [d]),
        "stress_detailed": ("get_stress_data", [d]),
        "intensity_minutes": ("get_intensity_minutes_data", [d]),
        "body_composition": ("get_body_composition", [d]),
        "stats_and_body": ("get_stats_and_body", [d]),
        "body_battery": ("get_body_battery", [week_ago, d]),
        "body_battery_events": ("get_body_battery_events", [d]),
        "floors": ("get_floors", [d]),
        "blood_pressure": ("get_blood_pressure", [week_ago, d]),
        "hydration": ("get_hydration_data", [d]),
        "activities_for_date": ("get_activities_fordate", [d]),
    }

    for key, (method, args) in all_calls.items():
        result = _safe_call(client, method, *args)
        if result is not None:
            data[key] = result
        time.sleep(0.3)

    raw_activities = data.get("activities_for_date")
    if isinstance(raw_activities, dict):
        data["activities_for_date"] = raw_activities.get("ActivitiesForDay", {}).get("payload", [])
    activities = data.get("activities_for_date")
    if activities and isinstance(activities, list):
        detailed_activities = []
        for act in activities[:5]:
            act_id = act.get("activityId")
            if not act_id:
                continue
            detail = {
                "summary": act,
                "splits": _safe_call(client, "get_activity_splits", act_id),
                "hr_zones": _safe_call(client, "get_activity_hr_in_timezones", act_id),
                "weather": _safe_call(client, "get_activity_weather", act_id),
            }
            detailed_activities.append(detail)
            time.sleep(0.3)
        data["activities_detailed"] = detailed_activities

    data["devices"] = _safe_call(client, "get_devices")

    non_null = sum(1 for k, v in data.items() if v is not None and k != "_meta")
    log.info(f"Collected {non_null} data categories for {d}")

    return data
