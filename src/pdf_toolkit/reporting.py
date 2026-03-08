from __future__ import annotations

from dataclasses import asdict, dataclass, field
from datetime import UTC, datetime
from pathlib import Path
import csv
import json


@dataclass(slots=True)
class CommandReport:
    command: str
    status: str
    input_paths: list[str]
    output_paths: list[str]
    started_at: str
    finished_at: str
    duration_ms: int
    warnings: list[str] = field(default_factory=list)
    details: dict[str, object] = field(default_factory=dict)
    error: str | None = None


def utc_now_iso() -> str:
    return datetime.now(UTC).isoformat()


def write_command_report(report: CommandReport, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(json.dumps(asdict(report), indent=2), encoding="utf-8")


def write_json(data: dict[str, object], destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(json.dumps(data, indent=2), encoding="utf-8")


def write_batch_csv(rows: list[dict[str, object]], destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = [
        "job_name",
        "input_path",
        "output_path",
        "status",
        "duration_ms",
        "pages_processed",
        "warnings",
        "error_message",
    ]
    with destination.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)
