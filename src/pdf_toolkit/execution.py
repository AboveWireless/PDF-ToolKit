from __future__ import annotations

from pathlib import Path
import time
from typing import Callable

from pdf_toolkit.errors import PdfToolkitError, ValidationError
from pdf_toolkit.reporting import CommandReport, utc_now_iso, write_command_report


def check_overwrite(paths: list[Path], overwrite: bool) -> None:
    if overwrite:
        return
    existing = [path for path in paths if path.exists()]
    if existing:
        joined = ", ".join(str(path) for path in existing)
        raise ValidationError(f"Refusing to overwrite existing output(s): {joined}. Use --overwrite to allow it.")


def run_mutation(
    *,
    command: str,
    input_paths: list[Path],
    planned_outputs: list[Path],
    report_path: Path | None,
    overwrite: bool,
    action: Callable[[], dict[str, object]],
) -> dict[str, object]:
    check_overwrite(planned_outputs, overwrite)
    started_at = utc_now_iso()
    started_ms = time.perf_counter()

    try:
        result = action()
        output_paths = [str(path) for path in result.get("outputs", [])]
        report = CommandReport(
            command=command,
            status="success",
            input_paths=[str(path) for path in input_paths],
            output_paths=output_paths,
            started_at=started_at,
            finished_at=utc_now_iso(),
            duration_ms=int((time.perf_counter() - started_ms) * 1000),
            warnings=list(result.get("warnings", [])),
            details=dict(result.get("details", {})),
        )
        if report_path:
            write_command_report(report, report_path)
        return result
    except PdfToolkitError as exc:
        report = CommandReport(
            command=command,
            status="error",
            input_paths=[str(path) for path in input_paths],
            output_paths=[],
            started_at=started_at,
            finished_at=utc_now_iso(),
            duration_ms=int((time.perf_counter() - started_ms) * 1000),
            error=str(exc),
        )
        if report_path:
            write_command_report(report, report_path)
        raise
    except Exception as exc:
        report = CommandReport(
            command=command,
            status="error",
            input_paths=[str(path) for path in input_paths],
            output_paths=[],
            started_at=started_at,
            finished_at=utc_now_iso(),
            duration_ms=int((time.perf_counter() - started_ms) * 1000),
            error=str(exc),
        )
        if report_path:
            write_command_report(report, report_path)
        raise
