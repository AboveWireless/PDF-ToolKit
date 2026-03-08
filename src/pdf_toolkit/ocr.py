from __future__ import annotations

from dataclasses import asdict, dataclass
import os
from pathlib import Path
import subprocess

from pdf_toolkit.core import ensure_dir, ensure_parent_dir, extract_text_by_page, inspect_pdf
from pdf_toolkit.environment import ensure_command_available
from pdf_toolkit.reporting import write_json


@dataclass(slots=True)
class ScanPageInfo:
    page_number: int
    mode: str
    text_characters: int
    image_count: int


def scan_detect(input_path: Path) -> dict[str, object]:
    from pypdf import PdfReader

    text_pages = extract_text_by_page(input_path)
    reader = PdfReader(str(input_path))
    pages: list[ScanPageInfo] = []
    modes: set[str] = set()
    for index, page in enumerate(reader.pages, start=1):
        text_chars = len(text_pages[index - 1].strip())
        image_count = len(list(page.images))
        if text_chars > 0 and image_count > 0:
            mode = "mixed"
        elif text_chars > 0:
            mode = "text-based"
        else:
            mode = "image-based"
        pages.append(ScanPageInfo(page_number=index, mode=mode, text_characters=text_chars, image_count=image_count))
        modes.add(mode)

    summary = next(iter(modes)) if len(modes) == 1 else "mixed"
    return {
        "summary": summary,
        "pages": [asdict(page) for page in pages],
        "page_count": len(pages),
    }


def _copy_pdf(input_path: Path, output_path: Path) -> None:
    ensure_parent_dir(output_path)
    shutil.copy2(input_path, output_path)


def run_ocr(
    input_path: Path,
    output_path: Path,
    *,
    language: str,
    skip_existing_text: bool,
    text_output: Path | None,
    json_output: Path | None,
    force: bool,
    temp_dir: Path,
) -> dict[str, object]:
    ocrmypdf_exe = ensure_command_available("ocrmypdf", "Install OCRmyPDF or bundle it in `vendor/bin`.")
    tesseract_exe = ensure_command_available("tesseract", "Install Tesseract OCR or bundle it in `vendor/bin`.")
    ghostscript_exe = ensure_command_available(
        "gswin64c",
        "Install Ghostscript or bundle `gswin64c.exe` in `vendor/bin`.",
        aliases=("gswin32c", "gs"),
    )

    temp_root = temp_dir / "ocr"
    ensure_dir(temp_root)
    detection = scan_detect(input_path)
    summary = str(detection["summary"])

    if summary == "text-based" and not force:
        _copy_pdf(input_path, output_path)
    else:
        ensure_parent_dir(output_path)
        command = [ocrmypdf_exe, "--language", language]
        if skip_existing_text:
            command.append("--skip-text")
        elif force:
            command.append("--force-ocr")
        else:
            command.append("--redo-ocr")
        command.extend([str(input_path), str(output_path)])
        env = dict(os.environ)
        env.setdefault("OCRMYPDF_TESSERACT", tesseract_exe)
        env.setdefault("OCRMYPDF_GS", ghostscript_exe)
        bundled_tessdata = Path(tesseract_exe).resolve().parent.parent / "tessdata"
        if bundled_tessdata.exists():
            env.setdefault("TESSDATA_PREFIX", str(bundled_tessdata))
        subprocess.run(command, check=True, capture_output=True, text=True, env=env)

    page_text = extract_text_by_page(output_path)
    if text_output:
        ensure_parent_dir(text_output)
        text_output.write_text("\n\n".join(page_text), encoding="utf-8")
    if json_output:
        write_json(
            {
                "input_path": str(input_path),
                "output_path": str(output_path),
                "page_count": inspect_pdf(output_path).page_count,
                "pages": [
                    {
                        "page_number": index,
                        "text": text,
                        "confidence": None,
                    }
                    for index, text in enumerate(page_text, start=1)
                ],
            },
            json_output,
        )

    return {
        "outputs": [path for path in [output_path, text_output, json_output] if path is not None],
        "details": {
            "scan_summary": summary,
            "page_count": len(page_text),
            "language": language,
            "skip_existing_text": skip_existing_text,
            "forced": force,
        },
    }
