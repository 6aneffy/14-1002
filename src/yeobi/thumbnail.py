from __future__ import annotations

import base64
import io
from pathlib import Path

import fitz  # pymupdf
from PIL import Image

THUMBNAIL_MAX = 240  # px
VIEW_MAX = 1000  # px (링크로 볼 때 적당한 크기)
JPEG_QUALITY = 78


def _pdf_first_page_png(data: bytes, dpi: int = 110) -> bytes:
    doc = fitz.open(stream=data, filetype="pdf")
    try:
        page = doc[0]
        pix = page.get_pixmap(dpi=dpi)
        return pix.tobytes("png")
    finally:
        doc.close()


def _load_thumbnail_image(filename: str, data: bytes, max_px: int) -> Image.Image:
    ext = Path(filename).suffix.lower()
    if ext == ".pdf":
        raw = _pdf_first_page_png(data)
        img = Image.open(io.BytesIO(raw))
    else:
        img = Image.open(io.BytesIO(data))
    img.thumbnail((max_px, max_px))
    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")
    return img


def make_thumbnail_data_url(filename: str, data: bytes) -> str:
    """파일 bytes → 썸네일 base64 data URL (data_editor ImageColumn에 직접 사용)."""
    img = _load_thumbnail_image(filename, data, THUMBNAIL_MAX)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=JPEG_QUALITY, optimize=True)
    b64 = base64.standard_b64encode(buf.getvalue()).decode("ascii")
    return f"data:image/jpeg;base64,{b64}"


def make_thumbnail_png_bytes(filename: str, data: bytes, max_px: int = 200) -> bytes:
    """파일 bytes → PNG 썸네일 bytes (xlsx 삽입용, 축소)."""
    img = _load_thumbnail_image(filename, data, max_px)
    buf = io.BytesIO()
    img.save(buf, format="PNG", optimize=True)
    return buf.getvalue()


def make_view_data_url(filename: str, data: bytes) -> str:
    """링크로 열어볼 때 쓰는 뷰용 data URL (최대 VIEW_MAX px JPEG)."""
    img = _load_thumbnail_image(filename, data, VIEW_MAX)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=85, optimize=True)
    b64 = base64.standard_b64encode(buf.getvalue()).decode("ascii")
    return f"data:image/jpeg;base64,{b64}"


def make_full_image_png_bytes(filename: str, data: bytes) -> bytes:
    """xlsx 삽입용 원본 크기 PNG bytes. PDF는 DPI 200으로 첫 페이지 렌더."""
    ext = Path(filename).suffix.lower()
    if ext == ".pdf":
        return _pdf_first_page_png(data, dpi=200)
    img = Image.open(io.BytesIO(data))
    if img.mode not in ("RGB", "RGBA", "L"):
        img = img.convert("RGB")
    buf = io.BytesIO()
    img.save(buf, format="PNG", optimize=True)
    return buf.getvalue()
