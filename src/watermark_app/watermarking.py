"""Watermark utilities for supported document types."""

from __future__ import annotations

import io
from pathlib import Path

SUPPORTED_EXTENSIONS = {".docx", ".docm", ".xlsx", ".xlsm", ".pptx", ".pptm", ".pdf"}


def is_supported_extension(filename: str) -> bool:
    return Path(filename).suffix.lower() in SUPPORTED_EXTENSIONS


def apply_watermark(source_path: Path, output_path: Path, watermark_png_path: Path) -> None:
    ext = source_path.suffix.lower()
    if ext in {".docx", ".docm"}:
        _watermark_word(source_path, output_path, watermark_png_path)
        return
    if ext in {".xlsx", ".xlsm"}:
        _watermark_excel(source_path, output_path, watermark_png_path)
        return
    if ext in {".pptx", ".pptm"}:
        _watermark_powerpoint(source_path, output_path, watermark_png_path)
        return
    if ext == ".pdf":
        _watermark_pdf(source_path, output_path, watermark_png_path)
        return
    raise ValueError(f"Unsupported file extension: {ext}")


def _watermark_word(source_path: Path, output_path: Path, watermark_png_path: Path) -> None:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Inches

    document = Document(str(source_path))
    for section in document.sections:
        header = section.header
        paragraph = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run()
        run.add_picture(str(watermark_png_path), width=Inches(6.0))
    document.save(str(output_path))


def _watermark_excel(source_path: Path, output_path: Path, watermark_png_path: Path) -> None:
    from openpyxl import load_workbook
    from openpyxl.drawing.image import Image as XLImage

    keep_vba = source_path.suffix.lower() == ".xlsm"
    workbook = load_workbook(filename=str(source_path), keep_vba=keep_vba)
    for sheet in workbook.worksheets:
        image = XLImage(str(watermark_png_path))
        sheet.add_image(image, "A1")
    workbook.save(str(output_path))


def _watermark_powerpoint(source_path: Path, output_path: Path, watermark_png_path: Path) -> None:
    from pptx import Presentation

    presentation = Presentation(str(source_path))
    for slide in presentation.slides:
        # Scale watermark to ~60% of slide width and center it.
        target_width = int(presentation.slide_width * 0.6)
        picture = slide.shapes.add_picture(
            str(watermark_png_path),
            left=0,
            top=0,
            width=target_width,
        )
        picture.left = int((presentation.slide_width - picture.width) / 2)
        picture.top = int((presentation.slide_height - picture.height) / 2)
    presentation.save(str(output_path))


def _watermark_pdf(source_path: Path, output_path: Path, watermark_png_path: Path) -> None:
    from pypdf import PdfReader, PdfWriter
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas

    def _overlay_page_bytes(page_width: float, page_height: float) -> bytes:
        buffer = io.BytesIO()
        pdf = canvas.Canvas(buffer, pagesize=(page_width, page_height))
        img = ImageReader(str(watermark_png_path))
        img_width, img_height = img.getSize()

        # Fit watermark to max 70% of page width/height while preserving aspect ratio.
        max_width = page_width * 0.7
        max_height = page_height * 0.7
        scale = min(max_width / img_width, max_height / img_height)
        draw_width = img_width * scale
        draw_height = img_height * scale
        x = (page_width - draw_width) / 2
        y = (page_height - draw_height) / 2

        pdf.drawImage(
            str(watermark_png_path),
            x,
            y,
            width=draw_width,
            height=draw_height,
            mask="auto",
            preserveAspectRatio=True,
        )
        pdf.showPage()
        pdf.save()
        buffer.seek(0)
        return buffer.read()

    reader = PdfReader(str(source_path))
    writer = PdfWriter()
    for page in reader.pages:
        width = float(page.mediabox.width)
        height = float(page.mediabox.height)
        overlay_reader = PdfReader(io.BytesIO(_overlay_page_bytes(width, height)))
        page.merge_page(overlay_reader.pages[0])
        writer.add_page(page)

    with output_path.open("wb") as out_file:
        writer.write(out_file)
