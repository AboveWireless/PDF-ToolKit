from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw
import pytest
from reportlab.pdfgen import canvas
from reportlab.lib.colors import black


@pytest.fixture()
def sample_pdf(tmp_path: Path) -> Path:
    destination = tmp_path / "sample.pdf"
    pdf = canvas.Canvas(str(destination))
    pdf.setTitle("Toolkit Sample")
    pdf.drawString(72, 720, "Hello from page 1")
    pdf.showPage()
    pdf.drawString(72, 720, "Hello from page 2")
    pdf.showPage()
    pdf.drawString(72, 720, "Hello from page 3")
    pdf.showPage()
    pdf.save()
    return destination


@pytest.fixture()
def sample_images(tmp_path: Path) -> list[Path]:
    red = tmp_path / "red.png"
    blue = tmp_path / "blue.png"
    Image.new("RGB", (120, 90), color="red").save(red)
    Image.new("RGB", (120, 90), color="blue").save(blue)
    return [red, blue]


@pytest.fixture()
def image_pdf(tmp_path: Path, sample_images: list[Path]) -> Path:
    destination = tmp_path / "image.pdf"
    pdf = canvas.Canvas(str(destination))
    pdf.drawImage(str(sample_images[0]), 72, 620, width=140, height=100)
    pdf.drawString(72, 740, "PDF with embedded image")
    pdf.showPage()
    pdf.save()
    return destination


@pytest.fixture()
def form_pdf(tmp_path: Path) -> Path:
    destination = tmp_path / "form.pdf"
    pdf = canvas.Canvas(str(destination))
    pdf.drawString(72, 720, "Name:")
    pdf.acroForm.textfield(name="name", x=120, y=710, width=200, height=20, value="")
    pdf.drawString(72, 670, "Department:")
    pdf.acroForm.textfield(name="department", x=160, y=660, width=200, height=20, value="")
    pdf.save()
    return destination


@pytest.fixture()
def bookmarked_pdf(tmp_path: Path) -> Path:
    destination = tmp_path / "bookmarked.pdf"
    pdf = canvas.Canvas(str(destination))
    pdf.bookmarkPage("intro")
    pdf.addOutlineEntry("Intro", "intro", 0, closed=False)
    pdf.drawString(72, 720, "Introduction")
    pdf.showPage()
    pdf.bookmarkPage("summary")
    pdf.addOutlineEntry("Summary", "summary", 0, closed=False)
    pdf.drawString(72, 720, "Summary")
    pdf.save()
    return destination


@pytest.fixture()
def scanned_image_pdf(tmp_path: Path) -> Path:
    image_path = tmp_path / "scan.png"
    image = Image.new("RGB", (800, 1000), color="white")
    draw = ImageDraw.Draw(image)
    draw.text((80, 120), "SCANNED PAGE 1", fill="black")
    image.save(image_path)

    destination = tmp_path / "scanned.pdf"
    pdf = canvas.Canvas(str(destination))
    pdf.drawImage(str(image_path), 0, 0, width=595, height=842)
    pdf.showPage()
    pdf.save()
    return destination


@pytest.fixture()
def mixed_pdf(tmp_path: Path, scanned_image_pdf: Path) -> Path:
    destination = tmp_path / "mixed.pdf"
    pdf = canvas.Canvas(str(destination))
    pdf.drawString(72, 720, "Native text page")
    pdf.showPage()
    pdf.drawImage(str(tmp_path / "scan.png"), 0, 0, width=595, height=842)
    pdf.showPage()
    pdf.save()
    return destination


@pytest.fixture()
def table_pdf(tmp_path: Path) -> Path:
    destination = tmp_path / "table.pdf"
    pdf = canvas.Canvas(str(destination))
    start_x = 72
    start_y = 720
    row_height = 28
    col_width = 120
    rows = [
        ["Name", "Dept", "Score"],
        ["Jeff", "Ops", "95"],
        ["Ana", "Finance", "91"],
    ]
    for row_index in range(len(rows) + 1):
        y = start_y - row_index * row_height
        pdf.setStrokeColor(black)
        pdf.line(start_x, y, start_x + col_width * 3, y)
    for col_index in range(4):
        x = start_x + col_index * col_width
        pdf.line(x, start_y, x, start_y - row_height * len(rows))
    for row_index, row in enumerate(rows):
        for col_index, value in enumerate(row):
            pdf.drawString(start_x + 8 + col_index * col_width, start_y - 20 - row_index * row_height, value)
    pdf.showPage()
    pdf.save()
    return destination
