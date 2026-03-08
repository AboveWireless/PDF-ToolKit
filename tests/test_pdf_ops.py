from __future__ import annotations

from pathlib import Path

from pypdf import PdfReader

from pdf_toolkit.pdf_ops import (
    add_attachments,
    clear_metadata,
    compress_pdf,
    crop_pdf,
    decrypt_pdf,
    encrypt_pdf,
    extract_attachments,
    extract_images,
    extract_text,
    fill_form,
    images_to_pdf,
    inspect_pdf,
    list_attachments,
    list_bookmarks,
    list_form_fields,
    merge_pdfs,
    number_pages,
    parse_page_spec,
    remove_annotations,
    render_pdf,
    rotate_pdf,
    select_pages,
    set_metadata,
    split_pdf,
    stamp_text,
)


def test_inspect_pdf(sample_pdf: Path) -> None:
    info = inspect_pdf(sample_pdf)
    assert info.page_count == 3
    assert info.is_encrypted is False
    assert len(info.page_sizes) == 3
    assert info.metadata["Title"] == "Toolkit Sample"


def test_parse_page_spec_preserves_order() -> None:
    assert parse_page_spec("3,1-2", 3) == [2, 0, 1]


def test_merge_pdfs(sample_pdf: Path, tmp_path: Path) -> None:
    output = tmp_path / "merged.pdf"
    merge_pdfs([sample_pdf, sample_pdf], output)
    reader = PdfReader(str(output))
    assert len(reader.pages) == 6


def test_split_every_page(sample_pdf: Path, tmp_path: Path) -> None:
    outputs = split_pdf(sample_pdf, tmp_path / "split", every_page=True)
    assert len(outputs) == 3
    assert all(path.exists() for path in outputs)


def test_select_pages(sample_pdf: Path, tmp_path: Path) -> None:
    output = tmp_path / "selected.pdf"
    select_pages(sample_pdf, output, "3,1")
    reader = PdfReader(str(output))
    assert len(reader.pages) == 2


def test_rotate_pdf(sample_pdf: Path, tmp_path: Path) -> None:
    output = tmp_path / "rotated.pdf"
    rotate_pdf(sample_pdf, output, 90, "1")
    reader = PdfReader(str(output))
    assert reader.pages[0].rotation == 90
    assert reader.pages[1].rotation in (0, None)


def test_extract_text(sample_pdf: Path) -> None:
    text = extract_text(sample_pdf)
    assert "Hello from page 1" in text
    assert "Hello from page 3" in text


def test_encrypt_and_decrypt(sample_pdf: Path, tmp_path: Path) -> None:
    encrypted = tmp_path / "encrypted.pdf"
    decrypted = tmp_path / "decrypted.pdf"
    encrypt_pdf(sample_pdf, encrypted, "secret123")
    reader = PdfReader(str(encrypted))
    assert reader.is_encrypted is True
    decrypt_pdf(encrypted, decrypted, "secret123")
    decrypted_reader = PdfReader(str(decrypted))
    assert len(decrypted_reader.pages) == 3


def test_stamp_text(sample_pdf: Path, tmp_path: Path) -> None:
    output = tmp_path / "stamped.pdf"
    stamp_text(sample_pdf, output, "CONFIDENTIAL")
    reader = PdfReader(str(output))
    assert len(reader.pages) == 3


def test_set_and_clear_metadata(sample_pdf: Path, tmp_path: Path) -> None:
    updated = tmp_path / "metadata.pdf"
    cleared = tmp_path / "cleared.pdf"
    set_metadata(sample_pdf, updated, {"Author": "Jeff", "Subject": "PDF Toolkit"})
    info = inspect_pdf(updated)
    assert info.metadata["Author"] == "Jeff"
    assert info.metadata["Subject"] == "PDF Toolkit"
    clear_metadata(updated, cleared)
    cleared_info = inspect_pdf(cleared)
    assert "Author" not in cleared_info.metadata


def test_compress_pdf(sample_pdf: Path, tmp_path: Path) -> None:
    output = tmp_path / "compressed.pdf"
    compress_pdf(sample_pdf, output)
    reader = PdfReader(str(output))
    assert len(reader.pages) == 3


def test_number_pages(sample_pdf: Path, tmp_path: Path) -> None:
    output = tmp_path / "numbered.pdf"
    number_pages(sample_pdf, output, position="top-right")
    reader = PdfReader(str(output))
    assert len(reader.pages) == 3


def test_crop_pdf(sample_pdf: Path, tmp_path: Path) -> None:
    output = tmp_path / "cropped.pdf"
    crop_pdf(sample_pdf, output, left=18, right=18, top=18, bottom=18)
    reader = PdfReader(str(output))
    page = reader.pages[0]
    assert float(page.mediabox.width) < 595
    assert float(page.mediabox.height) < 842


def test_render_pdf(sample_pdf: Path, tmp_path: Path) -> None:
    output_dir = tmp_path / "rendered"
    outputs = render_pdf(sample_pdf, output_dir, dpi=96)
    assert len(outputs) == 3
    assert all(path.exists() for path in outputs)


def test_extract_images(image_pdf: Path, tmp_path: Path) -> None:
    output_dir = tmp_path / "images"
    outputs = extract_images(image_pdf, output_dir)
    assert len(outputs) >= 1
    assert outputs[0].exists()


def test_images_to_pdf(sample_images: list[Path], tmp_path: Path) -> None:
    output = tmp_path / "images.pdf"
    images_to_pdf(sample_images, output)
    reader = PdfReader(str(output))
    assert len(reader.pages) == 2


def test_attachments_round_trip(sample_pdf: Path, tmp_path: Path) -> None:
    attachment = tmp_path / "note.txt"
    attachment.write_text("hello attachment", encoding="utf-8")
    attached_pdf = tmp_path / "attached.pdf"
    add_attachments(sample_pdf, attached_pdf, [attachment])
    attachments = list_attachments(attached_pdf)
    assert attachments[0].name == "note.txt"

    extracted_dir = tmp_path / "attachments"
    extracted = extract_attachments(attached_pdf, extracted_dir)
    assert extracted[0].read_text(encoding="utf-8") == "hello attachment"


def test_list_and_fill_form_fields(form_pdf: Path, tmp_path: Path) -> None:
    fields = list_form_fields(form_pdf)
    assert {field.name for field in fields} == {"name", "department"}

    output = tmp_path / "filled-form.pdf"
    fill_form(form_pdf, output, {"name": "Jeff", "department": "Operations"})
    reader = PdfReader(str(output))
    form_fields = reader.get_fields() or {}
    assert form_fields["name"]["/V"] == "Jeff"
    assert form_fields["department"]["/V"] == "Operations"


def test_list_bookmarks(bookmarked_pdf: Path) -> None:
    bookmarks = list_bookmarks(bookmarked_pdf)
    assert bookmarks == ["Intro", "Summary"]


def test_remove_annotations(form_pdf: Path, tmp_path: Path) -> None:
    output = tmp_path / "clean.pdf"
    remove_annotations(form_pdf, output)
    reader = PdfReader(str(output))
    assert "/Annots" not in reader.pages[0]
