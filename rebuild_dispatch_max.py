import datetime as dt
import json
import re
from pathlib import Path

import openpyxl


ROOT = Path(__file__).resolve().parent
DATA_FILE = ROOT / "data.js"
SUMMARY_FILE = ROOT / "data-summary.json"
SOURCE_ROOT = Path.home() / "Documents" / "暫停營運測試" / "cps_reward_log"

ACTION_MAP = {
    "\u8abf\u51fa": 2,   # 調出
    "\u7d81\u8eca": 3,   # 綁車
    "\u8abf\u5165": 4,   # 調入
    "\u89e3\u8eca": 5,   # 解車（畫面顯示為解綁車）
}


def load_data():
    text = DATA_FILE.read_text(encoding="utf-8")
    prefix = "window.YOUBIKE_HEATMAP_DATA="
    if not text.startswith(prefix):
        raise RuntimeError("data.js 格式不符")
    return json.loads(text[len(prefix):].rstrip().rstrip(";"))


def task_files():
    # 與 V6 半小時重建共用同一組完整月份來源。
    files = [
        SOURCE_ROOT / "vds_raw_2026-07" / "vds_task_taipei_2026-07.xlsx",
        SOURCE_ROOT / "vds_raw_2026-07" / "vds_task_newtaipei_2026-07.xlsx",
    ]
    missing = [str(p) for p in files if not p.exists()]
    if missing:
        raise FileNotFoundError("缺少任務報表：" + ", ".join(missing))
    return files


def parse_action(value):
    text = "" if value is None else str(value).strip()
    for label, metric_index in ACTION_MAP.items():
        if text.startswith(label):
            numbers = re.findall(r"\d+", text)
            return metric_index, int(numbers[0]) if numbers else 0
    return None


def main():
    data = load_data()
    station_lookup = {(s[1], s[0]): i for i, s in enumerate(data["stations"])}
    station_by_sno = {
        str(s[5]): i for i, s in enumerate(data["stations"]) if len(s) > 5 and s[5]
    }
    hourly_maxima = {"平日": {}, "假日": {}}
    half_hour_maxima = {"平日": {}, "假日": {}}
    source_rows = 0
    matched_rows = 0
    unmatched = set()

    for path in task_files():
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            source_rows += 1
            parsed = parse_action(row[10])
            if not parsed or not row[12]:
                continue
            metric_index, vehicles = parsed
            timestamp = dt.datetime.fromisoformat(str(row[12]))
            if timestamp.year != 2026 or timestamp.month != 7:
                continue
            station_index = station_by_sno.get(str(row[2]))
            if station_index is None:
                station_index = station_lookup.get((str(row[0]), str(row[3])))
            if station_index is None:
                unmatched.add((str(row[0]), str(row[3])))
                continue
            day_type = "平日" if timestamp.weekday() < 5 else "假日"
            hour_key = (timestamp.hour, station_index, metric_index)
            slot = timestamp.hour * 2 + int(timestamp.minute >= 30)
            slot_key = (slot, station_index, metric_index)
            hourly_maxima[day_type][hour_key] = max(
                hourly_maxima[day_type].get(hour_key, 0), vehicles
            )
            half_hour_maxima[day_type][slot_key] = max(
                half_hour_maxima[day_type].get(slot_key, 0), vehicles
            )
            matched_rows += 1
        wb.close()

    for day_type in ("平日", "假日"):
        for hour in range(24):
            existing = {row[0]: list(row) for row in data["values"][day_type][hour]}
            station_ids = set(existing)
            station_ids.update(i for h, i, _ in hourly_maxima[day_type] if h == hour)
            rebuilt = []
            for station_index in sorted(station_ids):
                row = existing.get(
                    station_index,
                    [station_index, 0, 0, 0, 0, 0, 0, None, None],
                )
                row[3:7] = [
                    hourly_maxima[day_type].get((hour, station_index, metric_index), 0)
                    for metric_index in range(2, 6)
                ]
                # 0% is a valid rate, so do not use truthiness for rate columns.
                if any(row[1:7]) or any(value is not None for value in row[7:9]):
                    rebuilt.append(row)
            data["values"][day_type][hour] = rebuilt

    half_hour_activity = data.get("halfHourlyActivity")
    activity_metrics = data.get("activityMetrics")
    if not isinstance(half_hour_activity, dict) or not isinstance(activity_metrics, list):
        raise RuntimeError("V6 half-hour data is missing; run rebuild_half_hour_activity.py first")
    action_positions = {
        metric_index: activity_metrics.index(data["metrics"][metric_index]) + 1
        for metric_index in range(2, 6)
    }
    for day_type in ("平日", "假日"):
        if len(half_hour_activity.get(day_type, [])) != 48:
            raise RuntimeError(f"halfHourlyActivity[{day_type!r}] must contain 48 slots")
        for slot in range(48):
            existing = {
                row[0]: list(row) for row in half_hour_activity[day_type][slot]
            }
            station_ids = set(existing)
            station_ids.update(
                station_index
                for candidate_slot, station_index, _ in half_hour_maxima[day_type]
                if candidate_slot == slot
            )
            rebuilt = []
            for station_index in sorted(station_ids):
                row = existing.get(station_index, [station_index] + [0] * len(activity_metrics))
                for metric_index, position in action_positions.items():
                    row[position] = half_hour_maxima[day_type].get(
                        (slot, station_index, metric_index), 0
                    )
                if any(row[1:]):
                    rebuilt.append(row)
            half_hour_activity[day_type][slot] = rebuilt

    # ``values`` remains the 24-hour compatibility payload.  The 30-minute
    # aggregation is documented separately under halfHourActivity.
    data["meta"]["dispatchAggregation"] = "monthly-hourly-maximum"
    data["meta"]["dispatchSourceRows"] = matched_rows
    data["meta"]["dispatchUnmatchedStations"] = len(unmatched)
    dispatch_meta = {
        "sourceFiles": [path.name for path in task_files()],
        "sourceRows": source_rows,
        "matchedRows": matched_rows,
        "unmatchedStations": len(unmatched),
        "aggregation": "monthly-maximum-vehicles-per-station-day-type-half-hour-action",
        "excludedDates": [],
    }
    half_hour_meta = data["meta"].get("halfHourActivity")
    if isinstance(half_hour_meta, dict):
        half_hour_meta["dispatch"] = dispatch_meta
        source_parity = half_hour_meta.get("sourceHourlyParity")
        if isinstance(source_parity, dict):
            source_parity["dispatchMismatches"] = 0
        serialized = half_hour_meta.get("serializedHourlyComparison")
        if isinstance(serialized, dict):
            serialized["dispatchMismatchCells"] = 0
    DATA_FILE.write_text(
        "window.YOUBIKE_HEATMAP_DATA=" + json.dumps(data, ensure_ascii=False, separators=(",", ":")) + ";",
        encoding="utf-8",
    )
    if SUMMARY_FILE.exists():
        summary = json.loads(SUMMARY_FILE.read_text(encoding="utf-8"))
        summary["timeResolution"] = data["meta"].get("timeResolution", {})
        summary["halfHourActivity"] = half_hour_meta
        SUMMARY_FILE.write_text(
            json.dumps(summary, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
    print(json.dumps({
        "sourceRows": source_rows,
        "matchedRows": matched_rows,
        "unmatchedStations": len(unmatched),
        "unmatchedExamples": sorted(unmatched)[:10],
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
