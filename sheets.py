from datetime import date
from pathlib import Path

from openpyxl import Workbook, load_workbook


def _safe_get(data: dict | None, *keys, default=None):
    current = data
    for key in keys:
        if not isinstance(current, dict):
            return default
        current = current.get(key)
        if current is None:
            return default
    return current


COLUMNS = [
    "date",
    "steps",
    "step_goal",
    "distance_m",
    "calories_total",
    "calories_active",
    "calories_bmr",
    "floors_up",
    "floors_down",
    "active_seconds",
    "sedentary_seconds",
    "moderate_intensity_min",
    "vigorous_intensity_min",
    "resting_hr",
    "min_hr",
    "max_hr",
    "sleep_seconds",
    "deep_sleep_seconds",
    "light_sleep_seconds",
    "rem_sleep_seconds",
    "awake_seconds",
    "sleep_score",
    "avg_stress",
    "max_stress",
    "rest_stress_seconds",
    "low_stress_seconds",
    "medium_stress_seconds",
    "high_stress_seconds",
    "body_battery_high",
    "body_battery_low",
    "body_battery_charged",
    "body_battery_drained",
    "spo2_avg",
    "spo2_lowest",
    "respiration_waking",
    "respiration_sleep",
    "respiration_highest",
    "respiration_lowest",
    "weight_kg",
    "bmi",
    "body_fat_pct",
    "muscle_mass_kg",
    "hydration_ml",
    "hydration_goal_ml",
    "hrv_last_night",
    "hrv_weekly_avg",
    "hrv_status",
    "training_readiness",
    "training_readiness_level",
    "vo2max",
    "training_load_7d",
    "training_status",
    "fitness_age",
    "chronological_age",
]


def _extract_row(data: dict, target_date: date) -> list:
    d = target_date.isoformat()
    summary = data.get("user_summary") or data.get("stats") or {}
    hr = data.get("heart_rates") or {}
    sleep_dto = _safe_get(data, "sleep", "dailySleepDTO") or {}
    stress = data.get("stress_all_day") or data.get("stress_detailed") or {}
    bb_list = data.get("body_battery")
    body = data.get("body_composition") or {}
    spo2 = data.get("spo2") or {}
    resp = data.get("respiration") or {}
    hydration = data.get("hydration") or {}
    hrv_summary = _safe_get(data, "hrv", "hrvSummary") or data.get("hrv") or {}
    tr = data.get("training_readiness") or {}
    if isinstance(tr, list):
        tr = tr[0] if tr else {}
    ts = data.get("training_status") or {}
    fa = data.get("fitness_age") or {}

    weight = body.get("weight")
    if weight and weight > 500:
        weight = weight / 1000
    muscle = body.get("muscleMass")
    if muscle and muscle > 500:
        muscle = muscle / 1000

    bb = None
    if bb_list and isinstance(bb_list, list):
        for entry in bb_list:
            if entry and (entry.get("calendarDate") == d or entry.get("date") == d):
                bb = entry
                break
        if not bb:
            bb = bb_list[-1] if bb_list else None
    bb = bb or {}

    bb_high = bb.get("bodyBatteryHighValue") or bb.get("highest")
    bb_low = bb.get("bodyBatteryLowValue") or bb.get("lowest")
    if not bb_high and bb.get("bodyBatteryValuesArray"):
        vals = [v[1] for v in bb["bodyBatteryValuesArray"] if v and len(v) > 1 and v[1] is not None]
        if vals:
            bb_high = max(vals)
            bb_low = min(vals)

    return [
        d,
        summary.get("totalSteps"),
        summary.get("dailyStepGoal"),
        summary.get("totalDistanceMeters"),
        summary.get("totalKilocalories"),
        summary.get("activeKilocalories"),
        summary.get("bmrKilocalories"),
        summary.get("floorsAscended"),
        summary.get("floorsDescended"),
        summary.get("activeSeconds"),
        summary.get("sedentarySeconds"),
        summary.get("moderateIntensityMinutes"),
        summary.get("vigorousIntensityMinutes"),
        hr.get("restingHeartRate"),
        hr.get("minHeartRate"),
        hr.get("maxHeartRate"),
        sleep_dto.get("sleepTimeSeconds"),
        sleep_dto.get("deepSleepSeconds"),
        sleep_dto.get("lightSleepSeconds"),
        sleep_dto.get("remSleepSeconds"),
        sleep_dto.get("awakeSleepSeconds"),
        _safe_get(sleep_dto, "sleepScores", "overall", "value") or sleep_dto.get("sleepScore"),
        stress.get("avgStressLevel") or stress.get("overallStressLevel"),
        stress.get("maxStressLevel"),
        stress.get("restStressDuration"),
        stress.get("lowStressDuration"),
        stress.get("mediumStressDuration"),
        stress.get("highStressDuration"),
        bb_high,
        bb_low,
        bb.get("charged"),
        bb.get("drained"),
        _safe_get(spo2, "averageSpO2") or _safe_get(spo2, "dailySpO2Values", "averageSpO2"),
        _safe_get(spo2, "lowestSpO2") or _safe_get(spo2, "dailySpO2Values", "lowestSpO2"),
        resp.get("avgWakingRespirationValue"),
        resp.get("avgSleepRespirationValue"),
        resp.get("highestRespirationValue"),
        resp.get("lowestRespirationValue"),
        weight,
        body.get("bmi"),
        body.get("bodyFat"),
        muscle,
        hydration.get("valueInML") or hydration.get("intakeinML"),
        hydration.get("goalInML"),
        hrv_summary.get("lastNight") or hrv_summary.get("lastNightAvg"),
        hrv_summary.get("weeklyAvg") or hrv_summary.get("weeklyAverage"),
        hrv_summary.get("status") or hrv_summary.get("currentStatus"),
        tr.get("score") or _safe_get(tr, "trainigReadinessDTO", "score"),
        tr.get("level") or _safe_get(tr, "trainigReadinessDTO", "level"),
        ts.get("vo2MaxValue") or _safe_get(ts, "mostRecentVO2Max", "generic", "vo2MaxValue"),
        ts.get("trainingLoad7Day"),
        ts.get("trainingStatusPhrase") or ts.get("currentTrainingStatus"),
        fa.get("fitnessAge"),
        fa.get("chronologicalAge"),
    ]


def _hm(seconds) -> str:
    if not seconds:
        return "N/A"
    s = int(seconds)
    return f"{s // 3600}h {(s % 3600) // 60}m"


def _km(meters) -> str:
    if not meters:
        return "N/A"
    return f"{meters / 1000:.1f} km"


def format_summary(data: dict, target_date: date) -> str:
    row = _extract_row(data, target_date)
    v = dict(zip(COLUMNS, row))
    weekday = target_date.strftime("%A")

    steps = v["steps"] or 0
    goal = v["step_goal"] or 0
    pct = round(steps / goal * 100) if goal else 0

    lines = [
        f"Garmin {v['date']} ({weekday})",
        "",
        f"Steps: {steps:,} / {goal:,} ({pct}%)",
        f"Distance: {_km(v['distance_m'])}",
        f"Calories: {int(v['calories_total'] or 0)} ({int(v['calories_active'] or 0)} active)",
        f"Floors: {int(v['floors_up'] or 0)} up / {int(v['floors_down'] or 0)} down",
        "",
        f"Resting HR: {v['resting_hr'] or 'N/A'} bpm",
        f"HR range: {v['min_hr'] or '?'} - {v['max_hr'] or '?'} bpm",
        f"HRV: {v['hrv_last_night'] or 'N/A'} ms (7d avg: {v['hrv_weekly_avg'] or 'N/A'})",
        "",
        f"Sleep: {_hm(v['sleep_seconds'])} (score: {v['sleep_score'] or 'N/A'})",
        f"  Deep {_hm(v['deep_sleep_seconds'])} / Light {_hm(v['light_sleep_seconds'])} / REM {_hm(v['rem_sleep_seconds'])}",
        "",
        f"Stress: avg {v['avg_stress'] or 'N/A'} / max {v['max_stress'] or 'N/A'}",
        f"Body Battery: {v['body_battery_low'] or '?'} - {v['body_battery_high'] or '?'} (+{v['body_battery_charged'] or '?'} / -{v['body_battery_drained'] or '?'})",
        f"SpO2: {v['spo2_avg'] or 'N/A'}% (low: {v['spo2_lowest'] or 'N/A'}%)",
    ]

    if v["training_readiness"]:
        lines.append(f"Training readiness: {v['training_readiness']} ({v['training_readiness_level'] or ''})")
    if v["vo2max"]:
        lines.append(f"VO2 Max: {v['vo2max']}")
    if v["weight_kg"]:
        lines.append(f"Weight: {v['weight_kg']:.1f} kg")

    return "\n".join(lines)


def append_to_excel(data: dict, target_date: date, output_dir: Path) -> Path:
    xlsx_path = output_dir / f"{target_date.year}-garmin.xlsx"

    if xlsx_path.exists():
        wb = load_workbook(xlsx_path)
        ws = wb.active
        existing_dates = set()
        for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
            if row[0]:
                existing_dates.add(str(row[0]))
        if target_date.isoformat() in existing_dates:
            for idx, row in enumerate(ws.iter_rows(min_row=2, max_col=1, values_only=True), start=2):
                if str(row[0]) == target_date.isoformat():
                    ws.delete_rows(idx)
                    break
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Garmin"
        ws.append(COLUMNS)

    ws.append(_extract_row(data, target_date))
    wb.save(xlsx_path)
    return xlsx_path
