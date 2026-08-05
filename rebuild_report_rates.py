# -*- coding: utf-8 -*-
"""Append hourly bike/dock visibility rates from monthly report JSON exports.

The source reports identify stations by ``s_no``.  Older ``data.js`` files do
not contain that identifier, so the first run bootstraps a mapping only when
the report and map station have one unique, exact ``(city, station name)``
match and the station-basic export confirms that identity.  The identifier is
saved as the sixth field of each station row; later runs use it directly and
reject identity conflicts.

No fuzzy station matching is performed.  Stations with no report match, or no
observations for a requested day type/status, receive JSON ``null`` rates.
"""

from __future__ import annotations

import argparse
import calendar
import datetime as dt
import email.utils
import hashlib
import json
import math
import os
import re
import zipfile
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any, Iterable
import xml.etree.ElementTree as ET


ROOT = Path(__file__).resolve().parent
DEFAULT_REPORT_DIR = Path.home() / "Documents" / "暫停營運測試" / "robot1157" / "data"
DEFAULT_REPORTS = (
    DEFAULT_REPORT_DIR / "report_taipei_2026-07.json",
    DEFAULT_REPORT_DIR / "report_newtaipei_2026-07.json",
)
DEFAULT_STATION_BASIC = DEFAULT_REPORT_DIR / "cps_station_basic_export.xlsx"
DEFAULT_EXCLUDED_DATES = ("2026-07-10", "2026-07-11", "2026-07-12")

DATA_PREFIX = "window.YOUBIKE_HEATMAP_DATA="
DAY_TYPES = ("平日", "假日")
LEGACY_METRICS = ("滿借", "空還", "調出", "綁車", "調入", "解綁車")
RATE_METRICS = ("見車率", "見位率")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--report",
        dest="reports",
        action="append",
        type=Path,
        help="Report JSON path. Repeat for multiple cities; defaults to the two July exports.",
    )
    parser.add_argument("--data-js", type=Path, default=ROOT / "data.js")
    parser.add_argument(
        "--station-basic",
        type=Path,
        default=DEFAULT_STATION_BASIC,
        help="CPS station-basic XLSX used to verify s_no identities and locate report-only stations.",
    )
    parser.add_argument("--summary", type=Path, default=ROOT / "data-summary.json")
    parser.add_argument(
        "--match-report", type=Path, default=ROOT / "station-match-report.json"
    )
    parser.add_argument("--period", default="2026-07", help="Target period in YYYY-MM format.")
    parser.add_argument(
        "--exclude",
        action="append",
        help="Excluded date in YYYY-MM-DD format. Repeat as needed; defaults to 7/10-7/12.",
    )
    parser.add_argument("--precision", type=int, default=2, help="Decimal places for percent values.")
    return parser.parse_args()


def load_data_js(path: Path) -> dict[str, Any]:
    text = path.read_text(encoding="utf-8")
    stripped = text.strip()
    if not stripped.startswith(DATA_PREFIX):
        raise RuntimeError(f"{path} does not start with {DATA_PREFIX!r}")
    payload = stripped[len(DATA_PREFIX) :]
    if payload.endswith(";"):
        payload = payload[:-1]
    data = json.loads(payload)
    if not isinstance(data, dict):
        raise RuntimeError("data.js payload must be a JSON object")
    return data


def parse_period(value: str) -> tuple[int, int]:
    try:
        parsed = dt.datetime.strptime(value, "%Y-%m")
    except ValueError as exc:
        raise ValueError(f"Invalid period {value!r}; expected YYYY-MM") from exc
    return parsed.year, parsed.month


def parse_iso_date(value: str) -> dt.date:
    try:
        return dt.date.fromisoformat(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"Invalid excluded date {value!r}; expected YYYY-MM-DD") from exc


def parse_report_date(value: Any) -> dt.date:
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"Missing or invalid report date: {value!r}")
    text = value.strip()
    try:
        parsed = email.utils.parsedate_to_datetime(text)
    except (TypeError, ValueError):
        parsed = None
    if parsed is not None:
        return parsed.date()
    try:
        return dt.datetime.fromisoformat(text.replace("Z", "+00:00")).date()
    except ValueError as exc:
        raise ValueError(f"Unrecognized report date: {value!r}") from exc


def normalize_text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def normalize_sno(value: Any) -> str:
    if isinstance(value, bool) or value is None:
        raise ValueError(f"Invalid s_no: {value!r}")
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if not value.is_integer():
            raise ValueError(f"Non-integral s_no: {value!r}")
        return str(int(value))
    text = str(value).strip()
    if not text:
        raise ValueError("Blank s_no")
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    if not text.isdigit():
        raise ValueError(f"Non-numeric s_no: {value!r}")
    return text


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


XLSX_MAIN_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
STATION_BASIC_HEADERS = (
    "國家",
    "城市",
    "行政區",
    "期別",
    "場站代號",
    "場站名稱",
    "場站名稱(英)",
    "總車位數",
    "啟用車柱數量",
    "座標",
    "土地來源",
    "場站狀態",
    "開站時間",
    "場站地址",
    "場站地址(英)",
    "備註",
)


def xlsx_column_index(reference: str) -> int:
    match = re.match(r"[A-Z]+", reference)
    if match is None:
        raise RuntimeError(f"Invalid XLSX cell reference: {reference!r}")
    value = 0
    for character in match.group(0):
        value = value * 26 + ord(character) - ord("A") + 1
    return value - 1


def xlsx_cell_text(cell: ET.Element) -> str:
    if cell.attrib.get("t") == "inlineStr":
        return "".join(node.text or "" for node in cell.findall(f".//{XLSX_MAIN_NS}t"))
    value = cell.find(f"{XLSX_MAIN_NS}v")
    return "" if value is None or value.text is None else value.text


def parse_coordinate(value: Any, *, context: str) -> tuple[float, float]:
    parts = normalize_text(value).split(",")
    if len(parts) != 2:
        raise RuntimeError(f"{context}: coordinate must be 'latitude,longitude', got {value!r}")
    try:
        latitude, longitude = (float(part.strip()) for part in parts)
    except ValueError as exc:
        raise RuntimeError(f"{context}: invalid coordinate {value!r}") from exc
    if not (math.isfinite(latitude) and math.isfinite(longitude)):
        raise RuntimeError(f"{context}: non-finite coordinate {value!r}")
    if not (-90 <= latitude <= 90 and -180 <= longitude <= 180):
        raise RuntimeError(f"{context}: coordinate outside valid range {value!r}")
    return latitude, longitude


def read_station_basic(
    path: Path,
) -> tuple[dict[str, dict[str, Any]], dict[tuple[str, str], list[dict[str, Any]]], dict[str, Any]]:
    path = path.resolve()
    if not path.is_file():
        raise FileNotFoundError(f"Missing station-basic XLSX: {path}")
    with zipfile.ZipFile(path) as archive:
        try:
            xml = archive.read("xl/worksheets/sheet1.xml")
        except KeyError as exc:
            raise RuntimeError(f"{path.name}: missing xl/worksheets/sheet1.xml") from exc
    root = ET.fromstring(xml)
    sheet_data = root.find(f"{XLSX_MAIN_NS}sheetData")
    if sheet_data is None:
        raise RuntimeError(f"{path.name}: worksheet has no sheetData")

    rows: list[list[str]] = []
    for row_node in sheet_data.findall(f"{XLSX_MAIN_NS}row"):
        values: dict[int, str] = {}
        for cell in row_node.findall(f"{XLSX_MAIN_NS}c"):
            values[xlsx_column_index(cell.attrib.get("r", ""))] = xlsx_cell_text(cell)
        rows.append([values.get(index, "") for index in range(len(STATION_BASIC_HEADERS))])
    if not rows:
        raise RuntimeError(f"{path.name}: worksheet is empty")
    if tuple(rows[0]) != STATION_BASIC_HEADERS:
        raise RuntimeError(
            f"{path.name}: unexpected headers. Expected={STATION_BASIC_HEADERS}; got={tuple(rows[0])}"
        )

    by_sno: dict[str, dict[str, Any]] = {}
    by_key: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
    for row_number, row in enumerate(rows[1:], start=2):
        if not any(normalize_text(value) for value in row):
            continue
        raw_sno = normalize_text(row[4])
        if not raw_sno.startswith(("5001", "5002")):
            continue
        try:
            sno = normalize_sno(raw_sno)
            latitude, longitude = parse_coordinate(
                row[9], context=f"{path.name} row {row_number}"
            )
        except ValueError as exc:
            raise RuntimeError(f"{path.name} row {row_number}: {exc}") from exc
        record = {
            "s_no": sno,
            "city": normalize_text(row[1]),
            "district": normalize_text(row[2]),
            "name": normalize_text(row[5]),
            "latitude": latitude,
            "longitude": longitude,
            "status": normalize_text(row[11]),
        }
        if not record["city"] or not record["district"] or not record["name"]:
            raise RuntimeError(f"{path.name} row {row_number}: blank city/district/name")
        if sno in by_sno:
            raise RuntimeError(f"{path.name}: duplicate station-basic s_no {sno}")
        by_sno[sno] = record
        by_key[(record["city"], record["name"])].append(record)

    stats = {
        "file": path.name,
        "sha256": sha256_file(path),
        "bytes": path.stat().st_size,
        "rows": len(by_sno),
        "duplicateExactCityNameKeys": sum(len(records) > 1 for records in by_key.values()),
    }
    return by_sno, by_key, stats


def calendar_dates(year: int, month: int) -> set[dt.date]:
    return {
        dt.date(year, month, day)
        for day in range(1, calendar.monthrange(year, month)[1] + 1)
    }


def day_type(value: dt.date) -> str:
    return DAY_TYPES[0] if value.weekday() < 5 else DAY_TYPES[1]


def read_reports(
    paths: Iterable[Path],
    *,
    year: int,
    month: int,
    excluded: set[dt.date],
) -> tuple[
    dict[str, dict[str, Any]],
    dict[tuple[str, str, str], tuple[list[float], int]],
    list[dict[str, Any]],
    set[dt.date],
    dict[str, int],
]:
    identities: dict[str, dict[str, Any]] = {}
    sums: dict[tuple[str, str, str], list[float]] = {}
    counts: Counter[tuple[str, str, str]] = Counter()
    source_stats: list[dict[str, Any]] = []
    included_dates: set[dt.date] = set()
    seen_records: set[tuple[str, dt.date, str]] = set()
    status_rows: Counter[str] = Counter()

    for path in paths:
        path = path.resolve()
        if not path.is_file():
            raise FileNotFoundError(f"Missing report JSON: {path}")
        with path.open("r", encoding="utf-8-sig") as handle:
            rows = json.load(handle)
        if not isinstance(rows, list):
            raise RuntimeError(f"{path.name}: top-level JSON must be an array")

        stats = {
            "file": path.name,
            "sha256": sha256_file(path),
            "bytes": path.stat().st_size,
            "inputRows": len(rows),
            "includedRows": 0,
            "excludedRows": 0,
            "outsidePeriodRows": 0,
        }

        for row_number, row in enumerate(rows, start=1):
            if not isinstance(row, dict):
                raise RuntimeError(f"{path.name} row {row_number}: expected an object")

            try:
                sno = normalize_sno(row.get("s_no"))
                city = normalize_text(row.get("city"))
                name = normalize_text(row.get("s_name"))
                district = normalize_text(row.get("SxPSArea"))
                status = normalize_text(row.get("status"))
                record_date = parse_report_date(row.get("date"))
            except ValueError as exc:
                raise RuntimeError(f"{path.name} row {row_number}: {exc}") from exc

            if not city or not name:
                raise RuntimeError(f"{path.name} row {row_number}: blank city or station name")
            if status not in RATE_METRICS:
                raise RuntimeError(
                    f"{path.name} row {row_number}: unexpected status {status!r}"
                )

            identity = identities.setdefault(
                sno, {"s_no": sno, "city": city, "name": name, "districts": set()}
            )
            if (identity["city"], identity["name"]) != (city, name):
                raise RuntimeError(
                    f"s_no {sno} has conflicting identities: "
                    f"{identity['city']}/{identity['name']} vs {city}/{name}"
                )
            if district:
                identity["districts"].add(district)

            if (record_date.year, record_date.month) != (year, month):
                stats["outsidePeriodRows"] += 1
                continue
            if record_date in excluded:
                stats["excludedRows"] += 1
                continue

            record_key = (sno, record_date, status)
            if record_key in seen_records:
                raise RuntimeError(
                    f"Duplicate report row for s_no/date/status: {sno}, {record_date}, {status}"
                )
            seen_records.add(record_key)

            hourly: list[float] = []
            for hour in range(24):
                key = f"r{hour}"
                value = row.get(key)
                if (
                    isinstance(value, bool)
                    or not isinstance(value, (int, float))
                    or not math.isfinite(float(value))
                    or not 0 <= float(value) <= 1
                ):
                    raise RuntimeError(
                        f"{path.name} row {row_number}: {key} must be numeric in [0, 1], got {value!r}"
                    )
                hourly.append(float(value))

            group_key = (sno, day_type(record_date), status)
            target = sums.setdefault(group_key, [0.0] * 24)
            for hour, value in enumerate(hourly):
                target[hour] += value
            counts[group_key] += 1
            included_dates.add(record_date)
            status_rows[status] += 1
            stats["includedRows"] += 1

        source_stats.append(stats)

    expected_dates = calendar_dates(year, month) - excluded
    if included_dates != expected_dates:
        missing = sorted(expected_dates - included_dates)
        extra = sorted(included_dates - expected_dates)
        raise RuntimeError(
            "Report date coverage mismatch. "
            f"Missing={','.join(map(str, missing)) or 'none'}; "
            f"extra={','.join(map(str, extra)) or 'none'}"
        )

    aggregates = {
        key: (values, counts[key])
        for key, values in sums.items()
    }
    return identities, aggregates, source_stats, included_dates, dict(status_rows)


def establish_station_mapping(
    data: dict[str, Any],
    identities: dict[str, dict[str, Any]],
    basic_by_sno: dict[str, dict[str, Any]],
    basic_by_key: dict[tuple[str, str], list[dict[str, Any]]],
    known_appended_snos: set[str] | None = None,
) -> tuple[list[str | None], dict[str, Any]]:
    stations = data.get("stations")
    if not isinstance(stations, list) or not stations:
        raise RuntimeError("data.js stations must be a non-empty array")

    station_keys: dict[tuple[str, str], int] = {}
    for index, station in enumerate(stations):
        if not isinstance(station, list) or len(station) < 5:
            raise RuntimeError(f"stations[{index}] must have at least five fields")
        key = (normalize_text(station[1]), normalize_text(station[0]))
        if not all(key):
            raise RuntimeError(f"stations[{index}] has a blank city/name")
        if key in station_keys:
            raise RuntimeError(
                f"Duplicate exact map station identity: {key[0]}/{key[1]} "
                f"at indexes {station_keys[key]} and {index}"
            )
        station_keys[key] = index

    report_by_key: dict[tuple[str, str], list[str]] = defaultdict(list)
    for sno, identity in identities.items():
        report_by_key[(identity["city"], identity["name"])].append(sno)
    ambiguous_report_keys = {
        key: sorted(snos) for key, snos in report_by_key.items() if len(snos) != 1
    }
    if ambiguous_report_keys:
        examples = list(ambiguous_report_keys.items())[:10]
        raise RuntimeError(f"Report city/name identities are not unique: {examples}")

    for sno, identity in identities.items():
        basic = basic_by_sno.get(sno)
        if basic is None:
            raise RuntimeError(f"Report s_no {sno} is absent from station-basic XLSX")
        if (basic["city"], basic["name"]) != (identity["city"], identity["name"]):
            raise RuntimeError(
                f"Report/station-basic identity conflict for s_no {sno}: "
                f"report={(identity['city'], identity['name'])}; "
                f"basic={(basic['city'], basic['name'])}"
            )

    station_snos: list[str | None] = []
    mapping_method = (
        "stored-s_no-or-unique-exact-report-city-name-validated-against-station-basic"
    )

    used_snos: dict[str, int] = {}
    for index, station in enumerate(stations):
        key = (normalize_text(station[1]), normalize_text(station[0]))
        stored = station[5] if len(station) >= 6 else None
        sno = None if stored is None or normalize_text(stored) == "" else normalize_sno(stored)
        if sno is not None:
            reference = identities.get(sno) or basic_by_sno.get(sno)
            if reference is None:
                raise RuntimeError(f"Stored s_no {sno} is absent from report and station-basic")
            if (reference["city"], reference["name"]) != key:
                raise RuntimeError(
                    f"Stored s_no {sno} identity conflicts at station {index}: "
                    f"map={key}, reference={(reference['city'], reference['name'])}"
                )
        else:
            report_candidates = report_by_key.get(key, [])
            if len(report_candidates) == 1:
                sno = report_candidates[0]
            else:
                basic_candidates = basic_by_key.get(key, [])
                if len(basic_candidates) == 1:
                    sno = basic_candidates[0]["s_no"]
                elif len(basic_candidates) > 1:
                    # This path is used only for stations absent from the monthly report.
                    # Resolve solely when district and coordinates identify one exact basic row.
                    exact_basic = [
                        candidate
                        for candidate in basic_candidates
                        if candidate["district"] == normalize_text(station[2])
                        and math.hypot(
                            candidate["latitude"] - float(station[3]),
                            candidate["longitude"] - float(station[4]),
                        )
                        <= 1e-7
                    ]
                    if len(exact_basic) == 1:
                        sno = exact_basic[0]["s_no"]

        if sno is not None:
            other = used_snos.get(sno)
            if other is not None:
                raise RuntimeError(f"s_no {sno} maps to station indexes {other} and {index}")
            used_snos[sno] = index
        station_snos.append(sno)
        if len(station) == 5:
            station.append(sno)
        else:
            station[5] = sno

    report_only_before_append = sorted(set(identities) - set(used_snos))
    all_appended_snos = set(known_appended_snos or ())
    all_appended_snos.update(report_only_before_append)
    for sno in report_only_before_append:
        identity = identities[sno]
        basic = basic_by_sno[sno]
        key = (identity["city"], identity["name"])
        if key in station_keys:
            raise RuntimeError(
                f"Report-only s_no {sno} unexpectedly collides with map station key {key}"
            )
        station_index = len(stations)
        station = [
            basic["name"],
            basic["city"],
            basic["district"],
            basic["latitude"],
            basic["longitude"],
            sno,
        ]
        stations.append(station)
        station_snos.append(sno)
        station_keys[key] = station_index
        used_snos[sno] = station_index

    invalid_known_appended = sorted(all_appended_snos - set(used_snos))
    if invalid_known_appended:
        raise RuntimeError(
            f"Previously appended s_no values are no longer mapped: {invalid_known_appended}"
        )
    appended_stations: list[dict[str, Any]] = []
    for sno in sorted(all_appended_snos):
        basic = basic_by_sno[sno]
        station_index = used_snos[sno]
        appended_stations.append(
            {
                "stationIndex": station_index,
                "s_no": sno,
                "city": basic["city"],
                "district": basic["district"],
                "name": basic["name"],
                "latitude": basic["latitude"],
                "longitude": basic["longitude"],
                "stationStatus": basic["status"],
            }
        )

    matched: list[dict[str, Any]] = []
    unmatched_stations: list[dict[str, Any]] = []
    district_mismatches: list[dict[str, Any]] = []
    for index, station in enumerate(stations):
        sno = station_snos[index]
        base = {
            "stationIndex": index,
            "city": normalize_text(station[1]),
            "name": normalize_text(station[0]),
            "mapDistrict": normalize_text(station[2]),
        }
        if sno is None or sno not in identities:
            if sno is not None:
                base["storedSno"] = sno
                base["reason"] = "stored-s_no-not-present-in-current-report"
            else:
                base["reason"] = "no-unique-exact-city-name-report-match"
            unmatched_stations.append(base)
            continue
        identity = identities[sno]
        basic = basic_by_sno[sno]
        coordinate_difference = math.hypot(
            float(station[3]) - basic["latitude"],
            float(station[4]) - basic["longitude"],
        )
        entry = {
            **base,
            "s_no": sno,
            "reportDistricts": sorted(identity["districts"]),
            "basicDistrict": basic["district"],
            "mapCoordinate": [station[3], station[4]],
            "basicCoordinate": [basic["latitude"], basic["longitude"]],
            "coordinateDifferenceDegrees": round(coordinate_difference, 9),
        }
        matched.append(entry)
        if base["mapDistrict"] != basic["district"]:
            district_mismatches.append(entry)

    mapped_snos = {entry["s_no"] for entry in matched}
    report_only = [
        {
            "s_no": sno,
            "city": identity["city"],
            "name": identity["name"],
            "reportDistricts": sorted(identity["districts"]),
        }
        for sno, identity in sorted(identities.items())
        if sno not in mapped_snos
    ]

    coordinate_mismatches = [
        entry for entry in matched if entry["coordinateDifferenceDegrees"] > 0.001
    ]
    report = {
        "mappingMethod": mapping_method,
        "fuzzyMatchingUsed": False,
        "stationBasicIdentityVerified": True,
        "mapStationCount": len(stations),
        "reportStationCount": len(identities),
        "matchedStationCount": len(matched),
        "unmatchedMapStationCount": len(unmatched_stations),
        "reportOnlyStationCount": len(report_only),
        "reportOnlyBeforeAppendCount": len(appended_stations),
        "appendedStationCount": len(appended_stations),
        "appendedSnos": [entry["s_no"] for entry in appended_stations],
        "districtMismatchCount": len(district_mismatches),
        "coordinateMismatchCount": len(coordinate_mismatches),
        "appendedStations": appended_stations,
        "matched": matched,
        "unmatchedMapStations": unmatched_stations,
        "reportOnlyStations": report_only,
        "districtMismatches": district_mismatches,
        "coordinateMismatches": coordinate_mismatches,
    }
    return station_snos, report


def extract_legacy_values(data: dict[str, Any]) -> dict[str, list[dict[int, list[Any]]]]:
    metrics = data.get("metrics")
    values = data.get("values")
    if not isinstance(metrics, list) or not isinstance(values, dict):
        raise RuntimeError("data.js metrics/values schema is invalid")
    missing = [metric for metric in LEGACY_METRICS if metric not in metrics]
    if missing:
        raise RuntimeError(f"data.js is missing legacy metrics: {missing}")
    positions = [metrics.index(metric) + 1 for metric in LEGACY_METRICS]
    station_count = len(data["stations"])

    extracted: dict[str, list[dict[int, list[Any]]]] = {}
    for label in DAY_TYPES:
        buckets = values.get(label)
        if not isinstance(buckets, list) or len(buckets) != 24:
            raise RuntimeError(f"values[{label!r}] must have exactly 24 hourly arrays")
        extracted[label] = []
        for hour, rows in enumerate(buckets):
            if not isinstance(rows, list):
                raise RuntimeError(f"values[{label!r}][{hour}] must be an array")
            by_station: dict[int, list[Any]] = {}
            for row in rows:
                if not isinstance(row, list) or not row:
                    raise RuntimeError(f"Invalid row in values[{label!r}][{hour}]")
                index = row[0]
                if isinstance(index, bool) or not isinstance(index, int) or not 0 <= index < station_count:
                    raise RuntimeError(f"Invalid station index {index!r} in {label} hour {hour}")
                if index in by_station:
                    raise RuntimeError(f"Duplicate station index {index} in {label} hour {hour}")
                try:
                    legacy = [row[position] for position in positions]
                except IndexError as exc:
                    raise RuntimeError(
                        f"Row for station {index} in {label} hour {hour} is too short"
                    ) from exc
                by_station[index] = legacy
            extracted[label].append(by_station)
    return extracted


def rounded_percent(total: float, count: int, precision: int) -> float:
    value = round((total / count) * 100.0, precision)
    return 0.0 if value == 0 else value


def rebuild_values(
    data: dict[str, Any],
    station_snos: list[str | None],
    aggregates: dict[tuple[str, str, str], tuple[list[float], int]],
    *,
    precision: int,
) -> tuple[dict[str, list[list[list[Any]]]], dict[str, Any]]:
    legacy = extract_legacy_values(data)
    result: dict[str, list[list[list[Any]]]] = {label: [] for label in DAY_TYPES}
    observation_counts: dict[str, dict[str, list[int]]] = {
        label: {status: [] for status in RATE_METRICS} for label in DAY_TYPES
    }

    for label in DAY_TYPES:
        for hour in range(24):
            bucket: list[list[Any]] = []
            for station_index, sno in enumerate(station_snos):
                legacy_values = legacy[label][hour].get(station_index, [0] * len(LEGACY_METRICS))
                rates: list[float | None] = []
                for status in RATE_METRICS:
                    aggregate = aggregates.get((sno, label, status)) if sno is not None else None
                    if aggregate is None:
                        rates.append(None)
                    else:
                        hourly_sums, count = aggregate
                        rates.append(rounded_percent(hourly_sums[hour], count, precision))
                        observation_counts[label][status].append(count)
                bucket.append([station_index, *legacy_values, *rates])
            result[label].append(bucket)

    count_summary: dict[str, dict[str, dict[str, int | None]]] = {}
    for label in DAY_TYPES:
        count_summary[label] = {}
        for status in RATE_METRICS:
            counts = observation_counts[label][status]
            count_summary[label][status] = {
                "minimum": min(counts) if counts else None,
                "maximum": max(counts) if counts else None,
            }
    return result, count_summary


def atomic_write_text(path: Path, text: str) -> None:
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.tmp")
    temporary.write_text(text, encoding="utf-8", newline="\n")
    os.replace(temporary, path)


def main() -> None:
    args = parse_args()
    if not 0 <= args.precision <= 6:
        raise ValueError("--precision must be between 0 and 6")
    reports = tuple(args.reports) if args.reports else DEFAULT_REPORTS
    excluded_strings = tuple(args.exclude) if args.exclude else DEFAULT_EXCLUDED_DATES
    excluded = {parse_iso_date(value) for value in excluded_strings}
    year, month = parse_period(args.period)
    if any((value.year, value.month) != (year, month) for value in excluded):
        raise ValueError("Every excluded date must be inside --period")

    data = load_data_js(args.data_js)
    existing_summary = (
        json.loads(args.summary.read_text(encoding="utf-8")) if args.summary.exists() else {}
    )
    identities, aggregates, sources, included_dates, status_rows = read_reports(
        reports, year=year, month=month, excluded=excluded
    )
    basic_by_sno, basic_by_key, basic_stats = read_station_basic(args.station_basic)
    known_appended_snos: set[str] = set()
    previous_usage = data.get("meta", {}).get("usageFrequency", {})
    previous_matching = previous_usage.get("matching", {}) if isinstance(previous_usage, dict) else {}
    if isinstance(previous_matching, dict):
        for value in previous_matching.get("appendedSnos", []):
            known_appended_snos.add(normalize_sno(value))
    if args.match_report.exists():
        previous_match_report = json.loads(args.match_report.read_text(encoding="utf-8"))
        if isinstance(previous_match_report, dict):
            for entry in previous_match_report.get("appendedStations", []):
                if isinstance(entry, dict) and entry.get("s_no") is not None:
                    known_appended_snos.add(normalize_sno(entry["s_no"]))
    station_snos, mapping = establish_station_mapping(
        data,
        identities,
        basic_by_sno,
        basic_by_key,
        known_appended_snos=known_appended_snos,
    )
    new_values, observation_counts = rebuild_values(
        data, station_snos, aggregates, precision=args.precision
    )

    kept_dates = calendar_dates(year, month) - excluded
    weekday_days = sum(value.weekday() < 5 for value in kept_dates)
    holiday_days = len(kept_dates) - weekday_days
    source_public = [
        {
            key: value
            for key, value in source.items()
            if key in {
                "file",
                "sha256",
                "bytes",
                "inputRows",
                "includedRows",
                "excludedRows",
                "outsidePeriodRows",
            }
        }
        for source in sources
    ]

    data["metrics"] = [*LEGACY_METRICS, *RATE_METRICS]
    data["values"] = new_values
    meta = data.setdefault("meta", {})
    data.pop("stationSnos", None)
    usage_frequency = {
        "period": args.period,
        "weekdayDays": weekday_days,
        "holidayDays": holiday_days,
        "excluded": [value.isoformat() for value in sorted(excluded)],
        "metrics": list(RATE_METRICS),
        "aggregation": "arithmetic-mean-by-station-day-type-hour-status",
        "observationDenominator": "available-report-days",
        "scale": "percent-0-100",
        "precision": args.precision,
        "includedDateCount": len(included_dates),
        "includedRowsByStatus": status_rows,
        "sources": source_public,
        "stationBasicSource": basic_stats,
        "matching": {
            "method": mapping["mappingMethod"],
            "fuzzyMatchingUsed": False,
            "stationBasicIdentityVerified": True,
            "mapStationCount": mapping["mapStationCount"],
            "reportStationCount": mapping["reportStationCount"],
            "matchedStationCount": mapping["matchedStationCount"],
            "unmatchedMapStationCount": mapping["unmatchedMapStationCount"],
            "reportOnlyBeforeAppendCount": mapping["reportOnlyBeforeAppendCount"],
            "appendedStationCount": mapping["appendedStationCount"],
            "appendedSnos": mapping["appendedSnos"],
            "reportOnlyStationCount": mapping["reportOnlyStationCount"],
            "districtMismatchCount": mapping["districtMismatchCount"],
            "coordinateMismatchCount": mapping["coordinateMismatchCount"],
        },
        "observationCounts": observation_counts,
    }
    meta["usageFrequency"] = usage_frequency

    if isinstance(existing_summary, dict) and isinstance(existing_summary.get("legacy"), dict):
        legacy_summary = existing_summary["legacy"]
    elif isinstance(existing_summary, dict):
        legacy_summary = {
            key: value
            for key, value in existing_summary.items()
            if key not in {"usageFrequency", "rateData"}
        }
    else:
        legacy_summary = {}
    summary = {
        "legacy": legacy_summary,
        "usageFrequency": usage_frequency,
    }

    mapping_output = {
        "period": args.period,
        "excluded": [value.isoformat() for value in sorted(excluded)],
        "sources": source_public,
        "stationBasicSource": basic_stats,
        **mapping,
    }
    # The complete set of successful matches is represented by the count and
    # reproducible station ids in data.js. Keep the report focused on records
    # that need inspection instead of committing thousands of routine rows.
    mapping_output.pop("matched", None)

    data_text = DATA_PREFIX + json.dumps(
        data, ensure_ascii=False, separators=(",", ":"), allow_nan=False
    ) + ";\n"
    summary_text = json.dumps(summary, ensure_ascii=False, indent=2, allow_nan=False) + "\n"
    mapping_text = json.dumps(mapping_output, ensure_ascii=False, indent=2, allow_nan=False) + "\n"

    atomic_write_text(args.data_js, data_text)
    atomic_write_text(args.summary, summary_text)
    atomic_write_text(args.match_report, mapping_text)

    print(
        json.dumps(
            {
                "period": args.period,
                "excluded": [value.isoformat() for value in sorted(excluded)],
                "weekdayDays": weekday_days,
                "holidayDays": holiday_days,
                "mapStations": len(data["stations"]),
                "reportStations": len(identities),
                "matchedStations": mapping["matchedStationCount"],
                "unmatchedMapStations": mapping["unmatchedMapStationCount"],
                "reportOnlyStations": mapping["reportOnlyStationCount"],
                "appendedStations": mapping["appendedStationCount"],
                "rowsPerHour": len(data["stations"]),
                "metrics": data["metrics"],
                "outputs": {
                    "dataJs": str(args.data_js.resolve()),
                    "summary": str(args.summary.resolve()),
                    "matchReport": str(args.match_report.resolve()),
                },
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
