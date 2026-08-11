"""Rebuild exact 30-minute activity data while preserving hourly rate data.

The existing ``values`` payload remains the hourly compatibility layer used by
the rate rebuild pipeline.  This script adds ``halfHourlyActivity`` with the
six activity metrics in ``activityMetrics`` order.  CPS reward events and VDS
tasks both retain timestamps to the second, so their slot is calculated as::

    hour * 2 + (minute >= 30)

見車率 and 見位率 are intentionally not rebuilt here because their report
source only exposes r0-r23.  The front end reads those hourly values with
``floor(slot / 2)``.
"""

from __future__ import annotations

import csv
import datetime as dt
import json
import re
import unicodedata
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

import openpyxl


ROOT = Path(__file__).resolve().parent
DATA_FILE = ROOT / "data.js"
SUMMARY_FILE = ROOT / "data-summary.json"
SOURCE_ROOT = Path.home() / "Documents" / "暫停營運測試" / "cps_reward_log"
EVENT_DIR = SOURCE_ROOT / "raw_2026-07"
DISPATCH_FILES = (
    SOURCE_ROOT / "vds_raw_2026-07" / "vds_task_taipei_2026-07.xlsx",
    SOURCE_ROOT / "vds_raw_2026-07" / "vds_task_newtaipei_2026-07.xlsx",
)

DATA_PREFIX = "window.YOUBIKE_HEATMAP_DATA="
DAY_TYPES = ("平日", "假日")
DAY_DENOMINATORS = {"平日": 22, "假日": 7}
EXCLUDED_EVENT_DATES = {"2026-07-10", "2026-07-11"}
ACTIVITY_METRICS = ("滿借", "空還", "調出", "綁車", "調入", "解綁車")
ACTION_MAP = {"調出": 2, "綁車": 3, "調入": 4, "解車": 5, "解綁車": 5}


def normalize_text(value: Any) -> str:
    text = unicodedata.normalize("NFKC", str(value or "")).strip().replace("臺", "台")
    return re.sub(r"\s+", "", text)


def normalize_sno(value: Any) -> str:
    text = str(value or "").strip()
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return text


def parse_datetime(value: Any) -> dt.datetime | None:
    text = str(value or "").strip()
    for fmt in (
        "%Y/%m/%d %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%Y/%m/%d %H:%M",
        "%Y-%m-%d %H:%M",
    ):
        try:
            return dt.datetime.strptime(text, fmt)
        except ValueError:
            pass
    return None


def slot_for(timestamp: dt.datetime) -> int:
    return timestamp.hour * 2 + int(timestamp.minute >= 30)


def load_data() -> dict[str, Any]:
    text = DATA_FILE.read_text(encoding="utf-8").strip()
    if not text.startswith(DATA_PREFIX):
        raise RuntimeError(f"{DATA_FILE.name} does not start with {DATA_PREFIX!r}")
    payload = text[len(DATA_PREFIX) :]
    if payload.endswith(";"):
        payload = payload[:-1]
    data = json.loads(payload)
    if len(data.get("values", {}).get("平日", [])) != 24:
        raise RuntimeError("Hourly compatibility values must contain 24 buckets")
    return data


def station_lookups(data: dict[str, Any]) -> tuple[dict[tuple[str, str], int], dict[str, int]]:
    by_name: dict[tuple[str, str], int] = {}
    by_sno: dict[str, int] = {}
    for index, station in enumerate(data["stations"]):
        key = (normalize_text(station[1]), normalize_text(station[0]))
        if key in by_name:
            raise RuntimeError(f"Duplicate normalized station key: {key}")
        by_name[key] = index
        if len(station) > 5 and station[5] not in (None, ""):
            sno = normalize_sno(station[5])
            if sno in by_sno:
                raise RuntimeError(f"Duplicate station s_no: {sno}")
            by_sno[sno] = index
    return by_name, by_sno


def load_unique_orders() -> tuple[dict[str, list[str]], Counter[str], list[Path]]:
    files = sorted(EVENT_DIR.glob("*.csv"))
    if len(files) != 44:
        raise RuntimeError(f"Expected 44 CPS event CSV files, found {len(files)} in {EVENT_DIR}")

    orders: dict[str, list[str]] = {}
    stats: Counter[str] = Counter()
    fields = (
        "帳號",
        "分類",
        "借車時間",
        "借車縣市",
        "借車場站",
        "還車時間",
        "還車縣市",
        "還車場站",
    )
    for path in files:
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            missing = {"訂單號", *fields} - set(reader.fieldnames or ())
            if missing:
                raise RuntimeError(f"{path.name} is missing columns: {sorted(missing)}")
            for row in reader:
                stats["sourceRows"] += 1
                order_id = str(row.get("訂單號") or "").strip()
                if not order_id:
                    stats["missingOrderRows"] += 1
                    continue
                values = [str(row.get(field) or "").strip() for field in fields]
                existing = orders.get(order_id)
                if existing is None:
                    orders[order_id] = values
                    continue
                stats["duplicateRows"] += 1
                categories = set(existing[1].split()) | set(values[1].split())
                existing[1] = " ".join(sorted(categories))
    stats["uniqueOrders"] = len(orders)
    return orders, stats, files


def aggregate_events(
    station_by_name: dict[tuple[str, str], int],
) -> tuple[dict[str, dict[tuple[int, int], list[int]]], dict[str, Any]]:
    orders, stats, files = load_unique_orders()
    counts: dict[str, dict[tuple[int, int], list[int]]] = {
        day: defaultdict(lambda: [0, 0]) for day in DAY_TYPES
    }
    unmatched: Counter[tuple[str, str]] = Counter()
    configs = (
        ("滿借", 0, 2, 3, 4),
        ("空還", 1, 5, 6, 7),
    )

    for values in orders.values():
        account, category = values[0], values[1]
        if not account.startswith("09"):
            stats["excludedNonMobileOrders"] += 1
            continue
        stats["mobileOrders"] += 1
        for label, metric_index, time_index, city_index, station_index in configs:
            if label not in category:
                continue
            timestamp = parse_datetime(values[time_index])
            if timestamp is None:
                stats[f"invalid{label}Datetime"] += 1
                continue
            date_text = timestamp.date().isoformat()
            if timestamp.year != 2026 or timestamp.month != 7:
                stats[f"excludedOutsideJuly{label}"] += 1
                continue
            if date_text in EXCLUDED_EVENT_DATES:
                stats[f"excludedTyphoon{label}"] += 1
                continue
            city = normalize_text(values[city_index])
            station = normalize_text(values[station_index])
            if city not in {"台北市", "新北市"} or not station:
                stats[f"excludedBadStation{label}"] += 1
                continue
            map_index = station_by_name.get((city, station))
            if map_index is None:
                unmatched[(city, station)] += 1
                continue
            day_type = "平日" if timestamp.weekday() < 5 else "假日"
            counts[day_type][(slot_for(timestamp), map_index)][metric_index] += 1
            stats[f"included{label}"] += 1

    if unmatched:
        examples = unmatched.most_common(10)
        raise RuntimeError(f"Unmatched CPS event stations ({len(unmatched)}): {examples}")
    return counts, {
        "sourceFiles": len(files),
        "sourceRows": stats["sourceRows"],
        "duplicateRows": stats["duplicateRows"],
        "uniqueOrders": stats["uniqueOrders"],
        "includedEvents": {
            "滿借": stats["included滿借"],
            "空還": stats["included空還"],
        },
        "excludedDates": sorted(EXCLUDED_EVENT_DATES),
        "aggregation": "event-count-per-station-day-type-half-hour-divided-by-calendar-days",
        "denominators": DAY_DENOMINATORS,
    }


def parse_action(value: Any) -> tuple[int, int] | None:
    text = str(value or "").strip()
    for label, metric_index in ACTION_MAP.items():
        if text.startswith(label):
            numbers = re.findall(r"\d+", text)
            return metric_index, int(numbers[0]) if numbers else 0
    return None


def aggregate_dispatch(
    station_by_name: dict[tuple[str, str], int],
    station_by_sno: dict[str, int],
) -> tuple[
    dict[str, dict[tuple[int, int, int], int]],
    dict[str, dict[tuple[int, int, int], int]],
    dict[str, Any],
]:
    missing = [str(path) for path in DISPATCH_FILES if not path.is_file()]
    if missing:
        raise FileNotFoundError("Missing VDS task files: " + ", ".join(missing))

    half_hour_maxima: dict[str, dict[tuple[int, int, int], int]] = {
        day: {} for day in DAY_TYPES
    }
    hourly_maxima: dict[str, dict[tuple[int, int, int], int]] = {
        day: {} for day in DAY_TYPES
    }
    stats: Counter[str] = Counter()
    unmatched: Counter[tuple[str, str, str]] = Counter()

    for path in DISPATCH_FILES:
        workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
        worksheet = workbook.active
        for row in worksheet.iter_rows(min_row=2, values_only=True):
            stats["sourceRows"] += 1
            parsed = parse_action(row[10] if len(row) > 10 else None)
            timestamp = parse_datetime(row[12] if len(row) > 12 else None)
            if parsed is None or timestamp is None:
                continue
            metric_index, vehicles = parsed
            if timestamp.year != 2026 or timestamp.month != 7:
                continue
            sno = normalize_sno(row[2] if len(row) > 2 else None)
            city = normalize_text(row[0] if row else None)
            station = normalize_text(row[3] if len(row) > 3 else None)
            map_index = station_by_sno.get(sno)
            if map_index is None:
                map_index = station_by_name.get((city, station))
            if map_index is None:
                unmatched[(sno, city, station)] += 1
                continue
            day_type = "平日" if timestamp.weekday() < 5 else "假日"
            half_key = (slot_for(timestamp), map_index, metric_index)
            hour_key = (timestamp.hour, map_index, metric_index)
            half_hour_maxima[day_type][half_key] = max(
                half_hour_maxima[day_type].get(half_key, 0), vehicles
            )
            hourly_maxima[day_type][hour_key] = max(
                hourly_maxima[day_type].get(hour_key, 0), vehicles
            )
            stats["matchedRows"] += 1
        workbook.close()

    if unmatched:
        examples = unmatched.most_common(10)
        raise RuntimeError(f"Unmatched VDS stations ({len(unmatched)}): {examples}")
    return half_hour_maxima, hourly_maxima, {
        "sourceFiles": [path.name for path in DISPATCH_FILES],
        "sourceRows": stats["sourceRows"],
        "matchedRows": stats["matchedRows"],
        "unmatchedStations": 0,
        "aggregation": "monthly-maximum-vehicles-per-station-day-type-half-hour-action",
        "excludedDates": [],
    }


def legacy_average(count: int, denominator: int) -> float:
    return round(round(count / denominator, 4), 3)


def verify_hourly_parity(
    data: dict[str, Any],
    event_counts: dict[str, dict[tuple[int, int], list[int]]],
    hourly_dispatch: dict[str, dict[tuple[int, int, int], int]],
) -> dict[str, int]:
    event_mismatches = 0
    dispatch_mismatches = 0
    examples: list[dict[str, Any]] = []
    station_count = len(data["stations"])
    hourly_rows = {
        day_type: [
            {row[0]: row for row in data["values"][day_type][hour]}
            for hour in range(24)
        ]
        for day_type in DAY_TYPES
    }
    for day_type in DAY_TYPES:
        denominator = DAY_DENOMINATORS[day_type]
        for hour in range(24):
            for station_index in range(station_count):
                row = hourly_rows[day_type][hour].get(
                    station_index, [station_index, *([0] * 6), None, None]
                )
                first = event_counts[day_type].get((hour * 2, station_index), [0, 0])
                second = event_counts[day_type].get((hour * 2 + 1, station_index), [0, 0])
                expected_events = [
                    legacy_average(first[index] + second[index], denominator)
                    for index in range(2)
                ]
                actual_events = [float(row[1] or 0), float(row[2] or 0)]
                if actual_events != expected_events:
                    event_mismatches += 1
                    if len(examples) < 8:
                        examples.append({
                            "kind": "event",
                            "dayType": day_type,
                            "hour": hour,
                            "stationIndex": station_index,
                            "expected": expected_events,
                            "actual": actual_events,
                        })
                expected_dispatch = [
                    hourly_dispatch[day_type].get((hour, station_index, metric_index), 0)
                    for metric_index in range(2, 6)
                ]
                actual_dispatch = [float(value or 0) for value in row[3:7]]
                if actual_dispatch != expected_dispatch:
                    dispatch_mismatches += 1
                    if len(examples) < 8:
                        examples.append({
                            "kind": "dispatch",
                            "dayType": day_type,
                            "hour": hour,
                            "stationIndex": station_index,
                            "expected": expected_dispatch,
                            "actual": actual_dispatch,
                        })
    if event_mismatches or dispatch_mismatches:
        raise RuntimeError(
            "Half-hour sources do not reconcile to the current hourly data: "
            + json.dumps(
                {
                    "eventMismatches": event_mismatches,
                    "dispatchMismatches": dispatch_mismatches,
                    "examples": examples,
                },
                ensure_ascii=False,
            )
        )
    return {"eventMismatches": 0, "dispatchMismatches": 0}


def build_half_hour_rows(
    event_counts: dict[str, dict[tuple[int, int], list[int]]],
    dispatch_maxima: dict[str, dict[tuple[int, int, int], int]],
) -> dict[str, list[list[list[float | int]]]]:
    result: dict[str, list[list[list[float | int]]]] = {
        day: [[] for _ in range(48)] for day in DAY_TYPES
    }
    for day_type in DAY_TYPES:
        denominator = DAY_DENOMINATORS[day_type]
        for slot in range(48):
            station_ids = {
                station_index
                for candidate_slot, station_index in event_counts[day_type]
                if candidate_slot == slot
            }
            station_ids.update(
                station_index
                for candidate_slot, station_index, _ in dispatch_maxima[day_type]
                if candidate_slot == slot
            )
            rows: list[list[float | int]] = []
            for station_index in sorted(station_ids):
                event_values = event_counts[day_type].get((slot, station_index), [0, 0])
                values: list[float | int] = [
                    legacy_average(event_values[0], denominator),
                    legacy_average(event_values[1], denominator),
                    *[
                        dispatch_maxima[day_type].get((slot, station_index, metric_index), 0)
                        for metric_index in range(2, 6)
                    ],
                ]
                if any(values):
                    rows.append([station_index, *values])
            result[day_type][slot] = rows
    return result


def compare_serialized_to_hourly(
    data: dict[str, Any],
    half_hour_rows: dict[str, list[list[list[float | int]]]],
) -> dict[str, float | int]:
    """Describe rounding deltas after serializing each half-hour separately."""

    event_mismatch_cells = 0
    event_max_absolute_delta = 0.0
    dispatch_mismatch_cells = 0
    for day_type in DAY_TYPES:
        hourly_maps = [
            {row[0]: row for row in data["values"][day_type][hour]}
            for hour in range(24)
        ]
        half_maps = [
            {row[0]: row for row in half_hour_rows[day_type][slot]}
            for slot in range(48)
        ]
        for hour in range(24):
            ids = set(hourly_maps[hour]) | set(half_maps[hour * 2]) | set(
                half_maps[hour * 2 + 1]
            )
            for station_index in ids:
                hourly = hourly_maps[hour].get(
                    station_index, [station_index, *([0] * 6), None, None]
                )
                first = half_maps[hour * 2].get(
                    station_index, [station_index, *([0] * 6)]
                )
                second = half_maps[hour * 2 + 1].get(
                    station_index, [station_index, *([0] * 6)]
                )
                for position in (1, 2):
                    delta = abs(
                        (float(first[position]) + float(second[position]))
                        - float(hourly[position] or 0)
                    )
                    if delta > 1e-9:
                        event_mismatch_cells += 1
                        event_max_absolute_delta = max(event_max_absolute_delta, delta)
                for position in range(3, 7):
                    if max(first[position], second[position]) != (hourly[position] or 0):
                        dispatch_mismatch_cells += 1
    return {
        "eventRoundedPairMismatchCells": event_mismatch_cells,
        "eventMaximumAbsoluteDelta": round(event_max_absolute_delta, 6),
        "dispatchMismatchCells": dispatch_mismatch_cells,
        "note": "event deltas are caused only by independently rounding each half-hour to 3 decimals",
    }


def main() -> None:
    data = load_data()
    station_by_name, station_by_sno = station_lookups(data)
    event_counts, event_stats = aggregate_events(station_by_name)
    dispatch_maxima, hourly_dispatch, dispatch_stats = aggregate_dispatch(
        station_by_name, station_by_sno
    )
    parity = verify_hourly_parity(data, event_counts, hourly_dispatch)
    half_hour_rows = build_half_hour_rows(event_counts, dispatch_maxima)
    serialized_comparison = compare_serialized_to_hourly(data, half_hour_rows)

    resolution = {
        "displayMinutes": 30,
        "slotCount": 48,
        "eventMinutes": 30,
        "dispatchMinutes": 30,
        "rateMinutes": 60,
        "rateSlotMapping": "floor(slot/2)",
    }
    half_hour_meta = {
        "metrics": list(ACTIVITY_METRICS),
        "event": event_stats,
        "dispatch": dispatch_stats,
        "sourceHourlyParity": parity,
        "serializedHourlyComparison": serialized_comparison,
    }

    data["activityMetrics"] = list(ACTIVITY_METRICS)
    data["halfHourlyActivity"] = half_hour_rows
    data["meta"]["timeResolution"] = resolution
    data["meta"]["halfHourActivity"] = half_hour_meta
    # ``values`` remains hourly for report-rate compatibility; the primary
    # half-hour dispatch definition lives in halfHourActivity.dispatch.
    data["meta"]["dispatchAggregation"] = "monthly-hourly-maximum"
    data["meta"]["dispatchSourceRows"] = dispatch_stats["matchedRows"]
    data["meta"]["dispatchUnmatchedStations"] = 0
    DATA_FILE.write_text(
        DATA_PREFIX + json.dumps(data, ensure_ascii=False, separators=(",", ":")) + ";",
        encoding="utf-8",
    )

    summary = json.loads(SUMMARY_FILE.read_text(encoding="utf-8"))
    summary["timeResolution"] = resolution
    summary["halfHourActivity"] = half_hour_meta
    SUMMARY_FILE.write_text(
        json.dumps(summary, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(
        json.dumps(
            {
                "stations": len(data["stations"]),
                "slots": {day: len(half_hour_rows[day]) for day in DAY_TYPES},
                "nonEmptyRows": {
                    day: sum(len(bucket) for bucket in half_hour_rows[day])
                    for day in DAY_TYPES
                },
                "event": event_stats,
                "dispatch": dispatch_stats,
                "sourceHourlyParity": parity,
                "serializedHourlyComparison": serialized_comparison,
                "dataBytes": DATA_FILE.stat().st_size,
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
