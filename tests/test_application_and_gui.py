from __future__ import annotations

from pathlib import Path
import json

import pytest

import pdf_toolkit.__main__ as package_main
from pdf_toolkit.application import execute_job, get_operation_definitions, prepare_request
from pdf_toolkit.branding import APP_NAME, APP_VERSION
from pdf_toolkit.cli import app
from pdf_toolkit.environment import collect_doctor_status
from pdf_toolkit.gui import main as gui_main


def test_package_entrypoint_targets_gui() -> None:
    assert package_main.main is gui_main


def test_operation_registry_matches_cli_commands() -> None:
    operation_ids = {definition.id for definition in get_operation_definitions()}
    cli_commands = {command.name for command in app.registered_commands if command.name}
    assert operation_ids == cli_commands


def test_prepare_request_resolves_output_paths(sample_pdf: Path, tmp_path: Path) -> None:
    (tmp_path / "pdf-toolkit.toml").write_text(
        """
[toolkit]
default_output_root = "generated"
""".strip(),
        encoding="utf-8",
    )
    request = prepare_request(
        "merge",
        {"inputs": [sample_pdf], "output": "merged.pdf"},
        report_path="reports/merge.json",
        cwd=tmp_path,
    )
    assert request.values["output"] == Path("generated") / "merged.pdf"
    assert request.report_path == Path("generated") / "reports" / "merge.json"


def test_execute_job_normalizes_success_result(sample_pdf: Path, tmp_path: Path) -> None:
    report_path = tmp_path / "merge-report.json"
    request = prepare_request(
        "merge",
        {"inputs": [sample_pdf, sample_pdf], "output": tmp_path / "merged.pdf"},
        report_path=report_path,
    )
    result = execute_job(request)
    assert result.status == "success"
    assert result.outputs == [tmp_path / "merged.pdf"]
    assert result.duration_ms >= 0
    payload = json.loads(report_path.read_text(encoding="utf-8"))
    assert payload["command"] == "merge"


def test_execute_job_normalizes_failure_result(sample_pdf: Path, tmp_path: Path) -> None:
    output = tmp_path / "merged.pdf"
    output.write_bytes(b"already there")
    request = prepare_request("merge", {"inputs": [sample_pdf], "output": output})
    result = execute_job(request)
    assert result.status == "error"
    assert "overwrite" in (result.error or "").lower()


def test_ocr_dependencies_are_marked_optional() -> None:
    statuses = {status.name: status for status in collect_doctor_status("ocr")}
    assert statuses["ocrmypdf"].required is False
    assert statuses["tesseract"].required is False
    assert statuses["gswin64c"].required is False


pytest.importorskip("PySide6")
pytest.importorskip("pytestqt")

from pdf_toolkit.gui import MainWindow, RedactionBoxEditor, create_app  # noqa: E402


def test_gui_launches_and_lists_operations(qtbot) -> None:
    create_app()
    window = MainWindow()
    qtbot.addWidget(window)
    window.show()
    assert window._tree.topLevelItemCount() >= 4
    assert window._current_definition is not None
    assert window.windowTitle() == f"{APP_NAME} {APP_VERSION}"
    assert not window.windowIcon().isNull()
    assert window._about_action.text() == "About PDF Toolkit"
    assert "desktop" in window._welcome_panel.summary.text().lower()


def test_redaction_box_editor_round_trips_values(qtbot) -> None:
    editor = RedactionBoxEditor()
    qtbot.addWidget(editor)
    editor.set_value(["1,10,20,30,40"])
    assert editor.value() == ["1,10,20,30,40"]
