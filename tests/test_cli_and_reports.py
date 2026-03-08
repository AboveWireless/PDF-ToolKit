from __future__ import annotations

from pathlib import Path
import json

from typer.testing import CliRunner

from pdf_toolkit.cli import app


runner = CliRunner()


def test_merge_writes_report(sample_pdf: Path, tmp_path: Path) -> None:
    output = tmp_path / "merged.pdf"
    report = tmp_path / "merge-report.json"
    result = runner.invoke(
        app,
        [
            "merge",
            str(sample_pdf),
            str(sample_pdf),
            "--output",
            str(output),
            "--report",
            str(report),
        ],
    )
    assert result.exit_code == 0
    payload = json.loads(report.read_text(encoding="utf-8"))
    assert payload["status"] == "success"
    assert payload["command"] == "merge"


def test_merge_respects_overwrite_exit_code(sample_pdf: Path, tmp_path: Path) -> None:
    output = tmp_path / "merged.pdf"
    output.write_bytes(b"already here")
    result = runner.invoke(
        app,
        [
            "merge",
            str(sample_pdf),
            str(sample_pdf),
            "--output",
            str(output),
        ],
    )
    assert result.exit_code == 2


def test_redact_dry_run_report(sample_pdf: Path, tmp_path: Path) -> None:
    report = tmp_path / "redact-report.json"
    result = runner.invoke(
        app,
        [
            "redact",
            str(sample_pdf),
            "--pattern",
            "Hello",
            "--dry-run",
            "--report",
            str(report),
        ],
    )
    assert result.exit_code == 0
    payload = json.loads(report.read_text(encoding="utf-8"))
    assert payload["details"]["match_count"] >= 1


def test_doctor_returns_dependency_exit_code() -> None:
    result = runner.invoke(app, ["doctor", "--feature", "ocr"])
    assert result.exit_code in {0, 3}
