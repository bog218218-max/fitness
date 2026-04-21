#!/Users/bogdanromanenko/Desktop/тренировки/.venv/bin/python
from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from datetime import datetime
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parent
STATIC_DIR = ROOT / "static"
WORKBOOK_PATH = ROOT.parent / "программа.xlsx"
HOST = "127.0.0.1"
PORT = 8123


@dataclass
class SetData:
    weight: float | int | None
    reps: int | None
    rir: float | int | str | None


@dataclass
class EntryData:
    date: str | None
    month: int | str | None
    summary: str | None
    notes: str | None
    sets: list[SetData]
    logged: bool
    volume: float | None
    estimated_1rm: float | None


def normalize_number(value: Any) -> float | int | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return int(value) if float(value).is_integer() else float(value)
    if isinstance(value, str):
        stripped = value.strip()
        if not stripped or stripped == "—":
            return None
        try:
            number = float(stripped.replace(",", "."))
        except ValueError:
            return None
        return int(number) if number.is_integer() else number
    return None


def normalize_text(value: Any) -> str | None:
    if value is None:
        return None
    if isinstance(value, str):
        stripped = value.strip()
        return stripped or None
    return str(value)


def is_section_title(value: Any) -> bool:
    return isinstance(value, str) and "|" in value and not value.startswith("Дата")


def parse_entry(row: list[Any]) -> EntryData:
    month_value = row[1]
    month = month_value if isinstance(month_value, str) else normalize_number(month_value)
    date = normalize_text(row[0])
    summary = normalize_text(row[14])
    notes = normalize_text(row[15])

    sets: list[SetData] = []
    volume = 0.0
    best_1rm = None

    for start in (2, 5, 8, 11):
        weight = normalize_number(row[start])
        reps = normalize_number(row[start + 1])
        rir = row[start + 2]
        if isinstance(rir, str):
            rir = normalize_text(rir)
        else:
            rir = normalize_number(rir)
        set_data = SetData(weight=weight, reps=int(reps) if reps is not None else None, rir=rir)
        sets.append(set_data)

        if weight is not None and reps is not None and weight > 0:
            volume += float(weight) * int(reps)
            estimated_1rm = float(weight) * (1 + int(reps) / 30)
            best_1rm = max(best_1rm or estimated_1rm, estimated_1rm)

    if volume == 0:
        volume = None

    logged = bool(summary or notes or any(item.weight is not None or item.reps is not None for item in sets))

    return EntryData(
        date=date,
        month=month,
        summary=summary,
        notes=notes,
        sets=sets,
        logged=logged,
        volume=round(volume, 1) if volume is not None else None,
        estimated_1rm=round(best_1rm, 1) if best_1rm is not None else None,
    )


def parse_structured_sheet(sheet) -> dict[str, Any]:
    section_rows = [
        row_index
        for row_index in range(1, sheet.max_row + 1)
        if is_section_title(sheet.cell(row_index, 1).value)
    ]

    exercises = []
    for index, section_row in enumerate(section_rows):
        next_section = section_rows[index + 1] if index + 1 < len(section_rows) else sheet.max_row + 1
        title = normalize_text(sheet.cell(section_row, 1).value)
        entries: list[EntryData] = []
        for row_index in range(section_row + 2, next_section):
            row = [sheet.cell(row_index, column).value for column in range(1, 17)]
            if all(value in (None, "") for value in row):
                continue
            if row[0] == "Дата":
                continue
            entries.append(parse_entry(row))
        exercises.append({"title": title, "entries": [asdict(entry) for entry in entries]})
    return {"title": sheet.title, "exercises": exercises}


def load_dashboard() -> dict[str, Any]:
    workbook = load_workbook(WORKBOOK_PATH, data_only=True)
    reports = {}
    for sheet_name in workbook.sheetnames:
        if "Отчет" in sheet_name:
            reports[sheet_name] = parse_structured_sheet(workbook[sheet_name])
    workbook.close()

    stat = WORKBOOK_PATH.stat()
    return {
        "workbook": {
            "path": str(WORKBOOK_PATH),
            "updated_at": datetime.fromtimestamp(stat.st_mtime).isoformat(),
            "size_bytes": stat.st_size,
        },
        "reports": reports,
    }


class AppHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(STATIC_DIR), **kwargs)

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/api/dashboard":
            self._send_json(load_dashboard())
            return
        if parsed.path == "/api/health":
            self._send_json({"ok": True})
            return
        if parsed.path == "/":
            self.path = "/index.html"
        return super().do_GET()

    def _send_json(self, payload: dict[str, Any]) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def main() -> None:
    server = ThreadingHTTPServer((HOST, PORT), AppHandler)
    print(f"Serving progress app at http://{HOST}:{PORT}")
    server.serve_forever()


if __name__ == "__main__":
    main()
