"""Resume file extraction helpers.

This module provides one extraction entry point for supported resume formats
while keeping the PDF, DOCX, and OCR implementations independent and reusable.
"""

import os
from typing import Optional, Type

from docx import Document
from pypdf import PdfReader

from analyzer.ocr import (
    OCRError,
    extract_images_from_page,
    extract_ocr_from_images,
)


class UnsupportedFileFormatError(ValueError):
    """Raised when a resume format is not supported by the application."""


# Threshold in characters: pages with less selectable text are checked for scanned images
MIN_PAGE_TEXT_CHARS = 30


def extract_pdf_text(
    filepath: str,
    reader_class: Optional[Type[PdfReader]] = None,
    enable_ocr: bool = True,
    min_text_chars: int = MIN_PAGE_TEXT_CHARS,
) -> str:
    """Extract text from every page of a PDF file, applying OCR to scanned pages."""
    reader = (reader_class or PdfReader)(filepath)
    text = ""

    try:
        for page in reader.pages:
            page_text = page.extract_text() or ""

            # Normal text extraction for pages with sufficient text
            if len(page_text.strip()) >= min_text_chars or not enable_ocr:
                if page_text:
                    text += page_text + "\n"
                continue

            # Page has insufficient extractable text: check for scanned images to OCR
            images = extract_images_from_page(page)
            if images:
                ocr_text = extract_ocr_from_images(images)
                parts = [p.strip() for p in (page_text, ocr_text) if p and p.strip()]
                if parts:
                    text += "\n".join(parts) + "\n"
            else:
                if page_text:
                    text += page_text + "\n"
    finally:
        # Ensure file streams opened by PdfReader are explicitly closed
        if hasattr(reader, "close"):
            try:
                reader.close()
            except Exception:
                pass
        stream = getattr(reader, "stream", None)
        if stream and hasattr(stream, "close"):
            try:
                stream.close()
            except Exception:
                pass

    return text


def extract_docx_text(filepath: str) -> str:
    """Extract paragraph and table text from a DOCX file."""
    document = Document(filepath)
    paragraphs = [paragraph.text for paragraph in document.paragraphs]
    tables = [
        "\n".join(
            cell.text
            for row in table.rows
            for cell in row.cells
        )
        for table in document.tables
    ]
    return "\n".join(paragraphs + tables)


def extract_resume_text(filepath: str) -> str:
    """Extract resume text based on the file extension.

    Raises:
        UnsupportedFileFormatError: if the path does not use a supported
            resume extension.
        OCRError: if OCR processing of scanned pages fails.
    """
    extension = os.path.splitext(filepath)[1].lower()
    if extension == ".pdf":
        return extract_pdf_text(filepath)
    if extension == ".docx":
        return extract_docx_text(filepath)

    raise UnsupportedFileFormatError("Only PDF and DOCX files are supported.")
