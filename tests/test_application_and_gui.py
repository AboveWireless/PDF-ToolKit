from __future__ import annotations

from pathlib import Path
import json

import pytest

import pdf_toolkit.__main__ as package_main
from pdf_toolkit.application import JobResult, execute_job, get_operation_definitions, prepare_request
from pdf_toolkit.branding import APP_NAME, APP_VERSION
from pdf_toolkit.cli import app
from pdf_toolkit.environment import collect_doctor_status
from pdf_toolkit.gui import main as gui_main
from pdf_toolkit.workflow_templates import get_workflow_templates


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


def test_llm_dependencies_are_marked_optional() -> None:
    statuses = {status.name: status for status in collect_doctor_status("llm")}
    assert statuses["openai"].required is False
    assert statuses["pydantic"].required is False
    assert statuses["OPENAI_API_KEY"].required is False


def test_prepare_request_requires_question_for_llm_qa(sample_pdf: Path, tmp_path: Path) -> None:
    try:
        prepare_request(
            "analyze-llm",
            {
                "input_path": sample_pdf,
                "output_dir": tmp_path / "analysis",
                "preset": "qa",
                "model": "gpt-5-mini",
            },
        )
    except Exception as exc:
        assert "Question is required" in str(exc)
    else:
        raise AssertionError("Expected validation failure for missing question")


pytest.importorskip("PySide6")
pytest.importorskip("pytestqt")

from PySide6.QtWidgets import QFrame, QLabel, QScrollArea, QWidget  # noqa: E402
from pdf_toolkit.gui import MainWindow, RedactionBoxEditor, create_app  # noqa: E402


def _fresh_window(qtbot) -> MainWindow:
    create_app()
    window = MainWindow()
    qtbot.addWidget(window)
    window._settings.clear()
    window._recent_runs = []
    window._last_request_context = None
    window._refresh_start_here()
    return window


def _is_descendant_of(widget: QWidget, ancestor: QWidget) -> bool:
    current = widget
    while current is not None:
        if current is ancestor:
            return True
        current = current.parentWidget()
    return False


def _workspace_structure(window: MainWindow) -> tuple[QScrollArea, QWidget, QFrame, QFrame]:
    scroll = window.findChild(QScrollArea, "FormScroll")
    assert scroll is not None
    scroll_content = scroll.widget()
    assert scroll_content is not None
    operation_card = window.findChild(QFrame, "OperationHero")
    footer_card = window.findChild(QFrame, "ActionBarCard")
    assert operation_card is not None
    assert footer_card is not None
    return scroll, scroll_content, operation_card, footer_card


def test_gui_launches_and_lists_operations(qtbot) -> None:
    window = _fresh_window(qtbot)
    window.show()
    assert window._tree.topLevelItemCount() >= 4
    assert window._current_definition is not None
    assert window.windowTitle() == f"{APP_NAME} {APP_VERSION}"
    assert not window.windowIcon().isNull()
    assert window._about_action.text() == "About PDF Toolkit"
    assert "desktop" in window._welcome_panel.summary.text().lower()
    assert "analyze-llm" in window._definition_map


def test_redaction_box_editor_round_trips_values(qtbot) -> None:
    editor = RedactionBoxEditor()
    qtbot.addWidget(editor)
    editor.set_value(["1,10,20,30,40"])
    assert editor.value() == ["1,10,20,30,40"]


def test_workflow_templates_cover_start_here_scenarios() -> None:
    templates = {template.id: template for template in get_workflow_templates()}
    assert set(templates) >= {
        "merge-invoice-packet",
        "split-page-ranges",
        "redact-pii-share",
        "export-tables-spreadsheet",
        "ocr-scanned-documents",
        "watch-incoming-folder",
    }
    assert templates["watch-incoming-folder"].target == "batch"
    assert templates["ocr-scanned-documents"].dependency_note is not None


def test_start_here_panel_lists_templates_and_optional_note(qtbot) -> None:
    window = _fresh_window(qtbot)
    labels = [window._start_here_panel._template_list.item(index).text() for index in range(window._start_here_panel._template_list.count())]
    assert "Merge Invoice Packet" in labels
    assert "Watch Incoming Folder" in labels
    window._start_here_panel.select_template("ocr-scanned-documents")
    assert "OCRmyPDF" in window._start_here_panel._template_note.text()
    assert window._start_here_panel._readiness.text()
    assert not window._start_here_panel._repeat_button.isEnabled()


def test_operation_template_prefills_form(qtbot) -> None:
    window = _fresh_window(qtbot)
    window._apply_template("merge-invoice-packet")
    assert window._current_definition is not None
    assert window._current_definition.id == "merge"
    assert window._collect_values()["output"] == "invoice-packet.pdf"
    assert "Template loaded: Merge Invoice Packet" in window._summary.text()


def test_batch_template_prefills_manifest_builder(qtbot) -> None:
    window = _fresh_window(qtbot)
    window._apply_template("watch-incoming-folder")
    assert window._current_definition is not None
    assert window._current_definition.id == "batch-run"
    payload = window._collect_values()["manifest_path"]
    assert isinstance(payload, dict)
    assert payload["output_root"] == "incoming-processed"
    steps = payload["steps"]
    assert isinstance(steps, list)
    assert [step["action"] for step in steps] == ["compress", "extract_text"]


def test_recent_activity_updates_after_success_and_repeat(qtbot, sample_pdf: Path, tmp_path: Path) -> None:
    window = _fresh_window(qtbot)
    output_path = tmp_path / "merged.pdf"
    report_path = tmp_path / "merge-report.json"
    window._select_operation("merge")
    window._apply_values_to_current_form(
        {
            "inputs": [str(sample_pdf), str(sample_pdf)],
            "output": str(output_path),
        },
        report_path=str(report_path),
    )
    window._last_request_context = {
        "operation_id": "merge",
        "label": "Combine PDFs",
        "values": window._collect_values(),
        "report_path": str(report_path),
    }
    window._handle_result(
        JobResult(
            operation_id="merge",
            status="success",
            outputs=[output_path],
            warnings=[],
            details={},
            error=None,
            duration_ms=42,
        ),
        report_path,
    )
    assert window._start_here_panel._repeat_button.isEnabled()
    assert "sample.pdf" in window._start_here_panel._recent_inputs.text()
    assert "Last task: Combine PDFs" in window._start_here_panel._recent_task.text()
    assert window._results._open_primary_button.isEnabled()
    assert window._results._open_folder_button.isEnabled()
    assert window._results._open_report_button.isEnabled()

    window._select_operation("split")
    window._repeat_last_task()
    values = window._collect_values()
    assert values["output"] == str(output_path)
    assert values["inputs"] == [str(sample_pdf), str(sample_pdf)]


def test_batch_recent_activity_tracks_real_input_path(qtbot, tmp_path: Path) -> None:
    window = _fresh_window(qtbot)
    incoming = tmp_path / "incoming"
    processed = tmp_path / "processed"
    incoming.mkdir()
    processed.mkdir()
    window._apply_template("watch-incoming-folder")
    payload = window._collect_values()["manifest_path"]
    assert isinstance(payload, dict)
    payload["input_root"] = str(incoming)
    payload["output_root"] = str(processed)
    payload["report_path"] = str(processed / "batch-report.json")
    window._apply_values_to_current_form({"manifest_path": payload})
    window._last_request_context = {
        "operation_id": "batch-run",
        "label": "Run Folder Workflow",
        "values": window._collect_values()["manifest_path"],
        "report_path": None,
    }
    window._handle_result(
        JobResult(
            operation_id="batch-run",
            status="success",
            outputs=[processed / "job-summary.json"],
            warnings=[],
            details={},
            error=None,
            duration_ms=18,
        ),
        None,
    )
    assert window._recent_runs[0]["input_paths"] == [str(incoming)]
    assert "incoming" in window._start_here_panel._recent_inputs.text().lower()


def test_workspace_scroll_contains_center_stack_and_keeps_footer_fixed(qtbot) -> None:
    window = _fresh_window(qtbot)
    scroll, scroll_content, operation_card, footer_card = _workspace_structure(window)
    assert scroll_content.objectName() == "WorkspaceScrollContent"
    assert scroll.widget() is scroll_content

    parameter_header = next(
        label for label in window.findChildren(QLabel)
        if label.property("workspaceRole") == "parameter-header"
    )
    parameter_caption = next(
        label for label in window.findChildren(QLabel)
        if label.property("workspaceRole") == "parameter-caption"
    )

    for widget in (window._start_here_panel, operation_card, parameter_header, parameter_caption, window._form_host):
        assert _is_descendant_of(widget, scroll_content)

    assert not _is_descendant_of(footer_card, scroll_content)
    assert footer_card.parentWidget() is not None
    assert not _is_descendant_of(window._report_input, scroll_content)


def test_workspace_scroll_structure_survives_operation_switching(qtbot) -> None:
    window = _fresh_window(qtbot)
    _scroll, scroll_content, operation_card, footer_card = _workspace_structure(window)

    window._apply_template("merge-invoice-packet")
    window._select_operation("split")
    window._select_operation("batch-run")

    assert _is_descendant_of(window._start_here_panel, scroll_content)
    assert _is_descendant_of(operation_card, scroll_content)
    assert _is_descendant_of(window._form_host, scroll_content)
    assert not _is_descendant_of(footer_card, scroll_content)


def test_workspace_scroll_structure_survives_all_operation_switches(qtbot) -> None:
    window = _fresh_window(qtbot)
    _scroll, scroll_content, operation_card, footer_card = _workspace_structure(window)

    for operation_id in window._definition_map:
        window._select_operation(operation_id)
        assert _is_descendant_of(window._start_here_panel, scroll_content)
        assert _is_descendant_of(operation_card, scroll_content)
        assert _is_descendant_of(window._form_host, scroll_content)
        assert not _is_descendant_of(footer_card, scroll_content)


def test_workspace_scroll_structure_survives_all_templates(qtbot) -> None:
    window = _fresh_window(qtbot)
    _scroll, scroll_content, operation_card, footer_card = _workspace_structure(window)

    for template in sorted(get_workflow_templates(), key=lambda item: item.id):
        window._apply_template(template.id)
        assert _is_descendant_of(window._start_here_panel, scroll_content)
        assert _is_descendant_of(operation_card, scroll_content)
        assert _is_descendant_of(window._form_host, scroll_content)
        assert not _is_descendant_of(footer_card, scroll_content)
