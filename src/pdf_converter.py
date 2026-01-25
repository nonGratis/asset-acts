import os
from io import BytesIO

import pymupdf as fitz
from PIL import Image
from docx2pdf import convert as docx2pdf_convert

from .config import OUTPUT_LOCAL_DIR_DOC, OUTPUT_LOCAL_DIR_PDF, OUTPUT_LOCAL_DIR_JPEG


def convert_to_pdf(docx_path: str) -> str:
    pdf_path = docx_path.replace(".docx", ".pdf")
    pdf_dir = os.path.dirname(pdf_path).replace(OUTPUT_LOCAL_DIR_DOC, OUTPUT_LOCAL_DIR_PDF)
    pdf_path = os.path.join(pdf_dir, os.path.basename(pdf_path))

    os.makedirs(pdf_dir, exist_ok=True)

    try:
        docx2pdf_convert(docx_path, pdf_path)
        return pdf_path
    except Exception as e:
        raise RuntimeError(f"PDF conversion failed: {e}")


def convert_to_jpeg(pdf_path: str) -> str:
    jpeg_path = pdf_path.replace(".pdf", ".jpg")
    jpeg_dir = os.path.dirname(jpeg_path).replace(OUTPUT_LOCAL_DIR_PDF, OUTPUT_LOCAL_DIR_JPEG)
    jpeg_path = os.path.join(jpeg_dir, os.path.basename(jpeg_path))

    os.makedirs(jpeg_dir, exist_ok=True)

    try:
        doc = fitz.open(pdf_path)
        page = doc.load_page(0)

        rect = page.rect
        zoom_x = 1280 / max(rect.width, rect.height)
        mat = fitz.Matrix(zoom_x, zoom_x)

        pix = page.get_pixmap(matrix=mat, alpha=False)
        img_data = pix.tobytes("ppm")

        img = Image.open(BytesIO(img_data))

        img.save(jpeg_path, "JPEG", quality=100, optimize=False)
        doc.close()

        return jpeg_path

    except Exception as e:
        raise RuntimeError(f"JPEG conversion failed: {e}")
