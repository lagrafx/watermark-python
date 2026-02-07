"""Word/Excel watermark utilities."""

from __future__ import annotations

from pathlib import Path


SUPPORTED_EXTENSIONS = {".docx", ".docm", ".xlsx", ".xlsm"}


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
