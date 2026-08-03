import datetime as dt
import json
import os
import re
from pathlib import Path

import openpyxl


ROOT = Path(__file__).resolve().parent
DOWNLOADS = Path.home() / "Downloads"
DATA_FILE = ROOT / "data.js"

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
    # 完整月份：台北市無尾碼檔、新北市 (2)；(1) 是台北市 7/1–7/7 的重複子集。
    files = [DOWNLOADS / "任務報表資料匯出.xlsx", DOWNLOADS / "任務報表資料匯出 (2).xlsx"]
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
    maxima = {"平日": {}, "假日": {}}
    matched_rows = 0
    unmatched = set()

    for path in task_files():
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            parsed = parse_action(row[10])
            if not parsed or not row[12]:
                continue
            metric_index, vehicles = parsed
            timestamp = dt.datetime.fromisoformat(str(row[12]))
            if timestamp.year != 2026 or timestamp.month != 7:
                continue
            station_index = station_lookup.get((str(row[0]), str(row[3])))
            if station_index is None:
                unmatched.add((str(row[0]), str(row[3])))
                continue
            day_type = "平日" if timestamp.weekday() < 5 else "假日"
            key = (timestamp.hour, station_index, metric_index)
            maxima[day_type][key] = max(maxima[day_type].get(key, 0), vehicles)
            matched_rows += 1
        wb.close()

    for day_type in ("平日", "假日"):
        for hour in range(24):
            existing = {row[0]: list(row) for row in data["values"][day_type][hour]}
            station_ids = set(existing)
            station_ids.update(i for h, i, _ in maxima[day_type] if h == hour)
            rebuilt = []
            for station_index in sorted(station_ids):
                row = existing.get(station_index, [station_index, 0, 0, 0, 0, 0, 0])
                row[3:7] = [
                    maxima[day_type].get((hour, station_index, metric_index), 0)
                    for metric_index in range(2, 6)
                ]
                if any(row[1:]):
                    rebuilt.append(row)
            data["values"][day_type][hour] = rebuilt

    data["meta"]["dispatchAggregation"] = "monthly-hourly-maximum"
    data["meta"]["dispatchSourceRows"] = matched_rows
    data["meta"]["dispatchUnmatchedStations"] = len(unmatched)
    DATA_FILE.write_text(
        "window.YOUBIKE_HEATMAP_DATA=" + json.dumps(data, ensure_ascii=False, separators=(",", ":")) + ";",
        encoding="utf-8",
    )
    print(json.dumps({
        "matchedRows": matched_rows,
        "unmatchedStations": len(unmatched),
        "unmatchedExamples": sorted(unmatched)[:10],
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
