from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
import re

import pdfplumber
import pypdfium2 as pdfium
from PIL import Image
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, RectangleObject
from reportlab.lib.colors import Color
from reportlab.pdfgen import canvas

from pdf_toolkit.errors import ValidationError


@dataclass(slots=True)
class PdfInfo:
    path: Path
    page_count: int
    is_encrypted: bool
    metadata: dict[str, str]
    page_sizes: list[tuple[float, float]]
    attachment_count: int
    form_field_count: int


@dataclass(slots=True)
class PdfAttachmentInfo:
    name: str
    size: int
    description: str | None


@dataclass(slots=True)
class PdfFormFieldInfo:
    name: str
    field_type: str
    value: str


def ensure_parent_dir(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def sanitize_filename(name: str, fallback: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "-", name).strip("-.")
    return cleaned or fallback


def new_writer_from_reader(reader: PdfReader, *, include_metadata: bool = True) -> PdfWriter:
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    if include_metadata and reader.metadata:
        writer.add_metadata(
            {
                str(key): str(value)
                for key, value in reader.metadata.items()
                if value is not None
            }
        )
    return writer


def read_unencrypted(path: Path) -> PdfReader:
    reader = PdfReader(str(path))
    if reader.is_encrypted:
        raise ValidationError(f"Cannot modify encrypted PDF: {path}")
    return reader


def inspect_pdf(path: Path) -> PdfInfo:
    reader = PdfReader(str(path))
    metadata = {
        str(key).lstrip("/"): str(value)
        for key, value in (reader.metadata or {}).items()
        if value is not None
    }
    page_sizes = [
        (float(page.mediabox.width), float(page.mediabox.height))
        for page in reader.pages
    ]
    return PdfInfo(
        path=path,
        page_count=len(reader.pages),
        is_encrypted=reader.is_encrypted,
        metadata=metadata,
        page_sizes=page_sizes,
        attachment_count=sum(1 for _ in reader.attachment_list),
        form_field_count=len(reader.get_fields() or {}),
    )


def merge_pdfs(inputs: list[Path], output: Path) -> None:
    writer = PdfWriter()
    for input_path in inputs:
        reader = read_unencrypted(input_path)
        for page in reader.pages:
            writer.add_page(page)
    ensure_parent_dir(output)
    with output.open("wb") as handle:
        writer.write(handle)


def parse_page_spec(spec: str, total_pages: int, *, allow_duplicates: bool = True) -> list[int]:
    pages: list[int] = []
    seen: set[int] = set()

    for part in (chunk.strip() for chunk in spec.split(",")):
        if not part:
            continue
        if "-" in part:
            start_text, end_text = part.split("-", maxsplit=1)
            start = int(start_text)
            end = int(end_text)
            if start > end:
                raise ValidationError(f"Invalid page range '{part}'. Start must be <= end.")
            page_numbers = range(start, end + 1)
        else:
            page_numbers = [int(part)]

        for number in page_numbers:
            if number < 1 or number > total_pages:
                raise ValidationError(f"Page {number} is out of bounds for a {total_pages}-page PDF.")
            index = number - 1
            if allow_duplicates:
                pages.append(index)
            elif index not in seen:
                pages.append(index)
                seen.add(index)

    if not pages:
        raise ValidationError("No valid pages were selected.")
    return pages


def merge_ranges_from_spec(spec: str) -> list[tuple[int, int]]:
    ranges: list[tuple[int, int]] = []
    for part in (chunk.strip() for chunk in spec.split(",")):
        if not part:
            continue
        if "-" not in part:
            number = int(part)
            ranges.append((number, number))
            continue
        start_text, end_text = part.split("-", maxsplit=1)
        start = int(start_text)
        end = int(end_text)
        if start > end:
            raise ValidationError(f"Invalid page range '{part}'. Start must be <= end.")
        ranges.append((start, end))
    if not ranges:
        raise ValidationError("No split ranges were provided.")
    return ranges


def split_pdf(input_path: Path, output_dir: Path, *, ranges: str | None = None, every_page: bool = False) -> list[Path]:
    if not ranges and not every_page:
        raise ValidationError("Provide --ranges or --every-page.")
    if ranges and every_page:
        raise ValidationError("Use either --ranges or --every-page, not both.")

    reader = read_unencrypted(input_path)
    ensure_dir(output_dir)
    outputs: list[Path] = []
    source_stem = input_path.stem

    if every_page:
        for index, page in enumerate(reader.pages, start=1):
            writer = PdfWriter()
            writer.add_page(page)
            destination = output_dir / f"{source_stem}-page-{index:03d}.pdf"
            with destination.open("wb") as handle:
                writer.write(handle)
            outputs.append(destination)
        return outputs

    assert ranges is not None
    for part_index, (start, end) in enumerate(merge_ranges_from_spec(ranges), start=1):
        if start < 1 or end > len(reader.pages):
            raise ValidationError(f"Range {start}-{end} exceeds the PDF page count of {len(reader.pages)}.")
        writer = PdfWriter()
        for page_number in range(start - 1, end):
            writer.add_page(reader.pages[page_number])
        destination = output_dir / f"{source_stem}-part-{part_index:02d}-p{start}-{end}.pdf"
        with destination.open("wb") as handle:
            writer.write(handle)
        outputs.append(destination)
    return outputs


def select_pages(input_path: Path, output: Path, page_spec: str) -> None:
    reader = read_unencrypted(input_path)
    selected_pages = parse_page_spec(page_spec, len(reader.pages), allow_duplicates=True)
    writer = PdfWriter()
    for page_index in selected_pages:
        writer.add_page(reader.pages[page_index])
    ensure_parent_dir(output)
    with output.open("wb") as handle:
        writer.write(handle)


def rotate_pdf(input_path: Path, output: Path, degrees: int, page_spec: str | None = None) -> None:
    if degrees % 90 != 0:
        raise ValidationError("Rotation must be a multiple of 90 degrees.")
    reader = read_unencrypted(input_path)
    writer = PdfWriter()
    selected = set(parse_page_spec(page_spec, len(reader.pages), allow_duplicates=False)) if page_spec else None
    for index, page in enumerate(reader.pages):
        if selected is None or index in selected:
            page.rotate(degrees)
        writer.add_page(page)
    ensure_parent_dir(output)
    with output.open("wb") as handle:
        writer.write(handle)


def extract_text_by_page(input_path: Path) -> list[str]:
    pages: list[str] = []
    with pdfplumber.open(str(input_path)) as pdf:
        for page in pdf.pages:
            pages.append((page.extract_text() or "").strip())
    return pages


def extract_text(input_path: Path) -> str:
    pages = extract_text_by_page(input_path)
    chunks = [f"--- Page {index} ---\n{text}".strip() for index, text in enumerate(pages, start=1)]
    return "\n\n".join(chunks).strip()


def encrypt_pdf(input_path: Path, output: Path, user_password: str, owner_password: str | None = None) -> None:
    reader = PdfReader(str(input_path))
    if reader.is_encrypted:
        raise ValidationError("PDF is already encrypted.")
    writer = new_writer_from_reader(reader)
    writer.encrypt(user_password=user_password, owner_password=owner_password or user_password)
    ensure_parent_dir(output)
    with output.open("wb") as handle:
        writer.write(handle)


def decrypt_pdf(input_path: Path, output: Path, password: str) -> None:
    reader = PdfReader(str(input_path))
    if not reader.is_encrypted:
        raise ValidationError("PDF is not encrypted.")
    if reader.decrypt(password) == 0:
        raise ValidationError("Incorrect password.")
    writer = new_writer_from_reader(reader)
    ensure_parent_dir(output)
    with output.open("wb") as handle:
        writer.write(handle)


def build_text_overlay(width: float, height: float, draw_callback) -> BytesIO:
    stream = BytesIO()
    pdf = canvas.Canvas(stream, pagesize=(width, height))
    draw_callback(pdf, width, height)
    pdf.showPage()
    pdf.save()
    stream.seek(0)
    return stream


def _draw_watermark(pdf: canvas.Canvas, width: float, height: float, text: str, font_size: int, opacity: float) -> None:
    pdf.setTitle("PDF Toolkit Watermark")
    try:
        pdf.setFillAlpha(opacity)
    except AttributeError:
        pass
    pdf.setFillColor(Color(0.5, 0.5, 0.5, alpha=opacity))
    pdf.setFont("Helvetica-Bold", font_size)
    pdf.saveState()
    pdf.translate(width / 2, height / 2)
    pdf.rotate(35)
    pdf.drawCentredString(0, 0, text)
    pdf.restoreState()


def stamp_text(input_path: Path, output: Path, text: str, font_size: int = 48, opacity: float = 0.2, page_spec: str | None = None) -> None:
    reader = read_unencrypted(input_path)
    selected = set(parse_page_spec(page_spec, len(reader.pages), allow_duplicates=False)) if page_spec else None
    writer = PdfWriter()
    for index, page in enumerate(reader.pages):
        if selected is None or index in selected:
            overlay_stream = build_text_overlay(
                float(page.mediabox.width),
                float(page.mediabox.height),
                lambda pdf, width, height: _draw_watermark(pdf, width, height, text, font_size, opacity),
            )
            overlay_page = PdfReader(overlay_stream).pages[0]
            page.merge_page(overlay_page)
        writer.add_page(page)
    ensure_parent_dir(output)
    with output.open("wb") as handle:
        writer.write(handle)


def set_metadata(input_path: Path, output: Path, metadata: dict[str, str], *, clear_existing: bool = False) -> None:
    reader = read_unencrypted(input_path)
    writer = new_writer_from_reader(reader, include_metadata=not clear_existing)
    writer.add_metadata({(key if key.startswith("/") else f"/{key}"): value for key, value in metadata.items()})
    ensure_parent_dir(output)
    with output.open("wb") as handle:
        writer.write(handle)


def clear_metadata(input_path: Path, output: Path) -> None:
    reader = read_unencrypted(input_path)
    writer = new_writer_from_reader(reader, include_metadata=False)
    ensure_parent_dir(output)
    with output.open("wb") as handle:
        writer.write(handle)


def compress_pdf(input_path: Path, output: Path) -> None:
    reader = read_unencrypted(input_path)
    writer = new_writer_from_reader(reader)
    for page in writer.pages:
        page.compress_content_streams()
    writer.compress_identical_objects(remove_identicals=True, remove_orphans=True)
    ensure_parent_dir(output)
    with output.open("wb") as handle:
        writer.write(handle)


def _draw_page_number(pdf: canvas.Canvas, width: float, height: float, text: str, position: str, margin: float, font_size: int, opacity: float) -> None:
    try:
        pdf.setFillAlpha(opacity)
    except AttributeError:
        pass
    pdf.setFillColor(Color(0.15, 0.15, 0.15, alpha=opacity))
    pdf.setFont("Helvetica", font_size)
    if position == "bottom-right":
        pdf.drawRightString(width - margin, margin, text)
    elif position == "bottom-center":
        pdf.drawCentredString(width / 2, margin, text)
    elif position == "bottom-left":
        pdf.drawString(margin, margin, text)
    elif position == "top-right":
        pdf.drawRightString(width - margin, height - margin, text)
    elif position == "top-center":
        pdf.drawCentredString(width / 2, height - margin, text)
    elif position == "top-left":
        pdf.drawString(margin, height - margin, text)
    else:
        raise ValidationError(f"Unsupported position '{position}'.")


def number_pages(input_path: Path, output: Path, *, format_text: str = "Page {page} of {total}", start_number: int = 1, page_spec: str | None = None, position: str = "bottom-right", margin: float = 36, font_size: int = 10, opacity: float = 0.85) -> None:
    reader = read_unencrypted(input_path)
    writer = PdfWriter()
    selected = set(parse_page_spec(page_spec, len(reader.pages), allow_duplicates=False)) if page_spec else None
    for index, page in enumerate(reader.pages):
        if selected is None or index in selected:
            number_text = format_text.format(page=index + start_number, total=len(reader.pages))
            overlay_stream = build_text_overlay(
                float(page.mediabox.width),
                float(page.mediabox.height),
                lambda pdf, width, height: _draw_page_number(pdf, width, height, number_text, position, margin, font_size, opacity),
            )
            overlay_page = PdfReader(overlay_stream).pages[0]
            page.merge_page(overlay_page)
        writer.add_page(page)
    ensure_parent_dir(output)
    with output.open("wb") as handle:
        writer.write(handle)


def crop_pdf(input_path: Path, output: Path, *, left: float = 0, right: float = 0, top: float = 0, bottom: float = 0, page_spec: str | None = None) -> None:
    reader = read_unencrypted(input_path)
    writer = PdfWriter()
    selected = set(parse_page_spec(page_spec, len(reader.pages), allow_duplicates=False)) if page_spec else None
    for index, page in enumerate(reader.pages):
        if selected is None or index in selected:
            old_box = page.mediabox
            new_left = float(old_box.left) + left
            new_bottom = float(old_box.bottom) + bottom
            new_right = float(old_box.right) - right
            new_top = float(old_box.top) - top
            if new_left >= new_right or new_bottom >= new_top:
                raise ValidationError(f"Crop margins are too large for page {index + 1}.")
            new_box = RectangleObject((new_left, new_bottom, new_right, new_top))
            page.mediabox = new_box
            page.cropbox = new_box
        writer.add_page(page)
    ensure_parent_dir(output)
    with output.open("wb") as handle:
        writer.write(handle)


def render_pdf(input_path: Path, output_dir: Path, *, dpi: int = 150, page_spec: str | None = None, image_format: str = "png") -> list[Path]:
    ensure_dir(output_dir)
    doc = pdfium.PdfDocument(str(input_path))
    selected = set(parse_page_spec(page_spec, len(doc), allow_duplicates=False)) if page_spec else None
    normalized_format = image_format.lower()
    save_format = {"jpg": "JPEG", "jpeg": "JPEG", "png": "PNG"}.get(normalized_format)
    if save_format is None:
        raise ValidationError("image_format must be png, jpg, or jpeg.")
    outputs: list[Path] = []
    try:
        for index in range(len(doc)):
            if selected is not None and index not in selected:
                continue
            page = doc[index]
            try:
                bitmap = page.render(scale=dpi / 72)
                image = bitmap.to_pil()
                destination = output_dir / f"{input_path.stem}-page-{index + 1:03d}.{normalized_format}"
                image.save(destination, format=save_format)
                image.close()
                outputs.append(destination)
            finally:
                page.close()
    finally:
        doc.close()
    return outputs


def extract_images(input_path: Path, output_dir: Path, *, page_spec: str | None = None) -> list[Path]:
    reader = read_unencrypted(input_path)
    selected = set(parse_page_spec(page_spec, len(reader.pages), allow_duplicates=False)) if page_spec else None
    ensure_dir(output_dir)
    outputs: list[Path] = []
    for page_index, page in enumerate(reader.pages, start=1):
        if selected is not None and page_index - 1 not in selected:
            continue
        for image_index, image in enumerate(page.images, start=1):
            suffix = Path(image.name).suffix or ".png"
            destination = output_dir / f"{input_path.stem}-page-{page_index:03d}-image-{image_index:02d}{suffix}"
            destination.write_bytes(image.data)
            outputs.append(destination)
    return outputs


def images_to_pdf(inputs: list[Path], output: Path) -> None:
    if not inputs:
        raise ValidationError("Provide at least one input image.")
    converted: list[Image.Image] = []
    try:
        for path in inputs:
            converted.append(Image.open(path).convert("RGB"))
        first, *rest = converted
        ensure_parent_dir(output)
        first.save(output, save_all=True, append_images=rest)
    finally:
        for image in converted:
            image.close()


def list_attachments(input_path: Path) -> list[PdfAttachmentInfo]:
    reader = PdfReader(str(input_path))
    return [
        PdfAttachmentInfo(name=attachment.name, size=attachment.size or len(attachment.content), description=attachment.description)
        for attachment in reader.attachment_list
    ]


def add_attachments(input_path: Path, output: Path, attachments: list[Path]) -> None:
    reader = read_unencrypted(input_path)
    writer = new_writer_from_reader(reader)
    for attachment_path in attachments:
        writer.add_attachment(attachment_path.name, attachment_path.read_bytes())
    ensure_parent_dir(output)
    with output.open("wb") as handle:
        writer.write(handle)


def extract_attachments(input_path: Path, output_dir: Path) -> list[Path]:
    reader = PdfReader(str(input_path))
    ensure_dir(output_dir)
    outputs: list[Path] = []
    for index, attachment in enumerate(reader.attachment_list, start=1):
        destination = output_dir / sanitize_filename(attachment.name, f"attachment-{index:02d}.bin")
        destination.write_bytes(attachment.content)
        outputs.append(destination)
    return outputs


def list_form_fields(input_path: Path) -> list[PdfFormFieldInfo]:
    reader = PdfReader(str(input_path))
    fields = reader.get_fields() or {}
    type_names = {"/Tx": "text", "/Btn": "button", "/Ch": "choice", "/Sig": "signature"}
    result: list[PdfFormFieldInfo] = []
    for name, details in fields.items():
        raw_type = str(details.get("/FT", "unknown"))
        result.append(PdfFormFieldInfo(name=name, field_type=type_names.get(raw_type, raw_type.strip("/")), value=str(details.get("/V", ""))))
    return result


def fill_form(input_path: Path, output: Path, values: dict[str, str]) -> None:
    reader = read_unencrypted(input_path)
    writer = PdfWriter()
    writer.clone_document_from_reader(reader)
    writer.set_need_appearances_writer(True)
    for page in writer.pages:
        writer.update_page_form_field_values(page, values, auto_regenerate=False)
    ensure_parent_dir(output)
    with output.open("wb") as handle:
        writer.write(handle)


def list_bookmarks(input_path: Path) -> list[str]:
    reader = PdfReader(str(input_path))
    try:
        outline = reader.outline
    except Exception:
        return []
    bookmarks: list[str] = []

    def walk(items) -> None:
        for item in items:
            if isinstance(item, list):
                walk(item)
            else:
                title = getattr(item, "title", None)
                if title:
                    bookmarks.append(str(title))

    if isinstance(outline, list):
        walk(outline)
    return bookmarks


def remove_annotations(input_path: Path, output: Path, *, page_spec: str | None = None) -> None:
    reader = read_unencrypted(input_path)
    writer = PdfWriter()
    selected = set(parse_page_spec(page_spec, len(reader.pages), allow_duplicates=False)) if page_spec else None
    for index, page in enumerate(reader.pages):
        if selected is None or index in selected:
            if "/Annots" in page:
                del page[NameObject("/Annots")]
        writer.add_page(page)
    ensure_parent_dir(output)
    with output.open("wb") as handle:
        writer.write(handle)
