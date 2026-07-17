#!/usr/bin/env python3
"""
Shared QR code resolution for invoices. Presentation-agnostic (returns image
bytes) so both the PDF (ReportLab) and Excel (openpyxl) generators can embed
the same image without duplicating QR-building or resolution logic.

Controlled by three optional clinic_config.json fields (not required — a
config without them behaves as if show_qr were false):
  "show_qr": true/false
  "qr_image_path": "templates/zelle_qr.jpg"  (a static, pre-made QR image,
      e.g. a bank/Zelle-issued QR — takes priority when set, since a
      real payment QR like Zelle's isn't just a URL a generator can
      reproduce)
  "qr_content": "https://..."  (falls back to the clinic's "website" if
      omitted; used to generate a QR only when qr_image_path is unset or
      the file can't be found)
"""
import io
from pathlib import Path
from typing import Optional

import qrcode

_REPO_ROOT = Path(__file__).parent


def generate_qr_png_bytes(content: str, box_size: int = 10, border: int = 2) -> bytes:
    """Generate a QR code encoding `content`, returned as PNG bytes."""
    img = qrcode.make(content, box_size=box_size, border=border)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def qr_settings(clinic: dict) -> tuple:
    """Resolve (show_qr, qr_content, qr_image_path) from a loaded clinic
    config dict. All three fields are optional — a config that predates
    this feature is treated as show_qr=False, not an error."""
    show_qr = bool(clinic.get("show_qr", False))
    qr_content = clinic.get("qr_content") or clinic.get("website", "")
    qr_image_path = clinic.get("qr_image_path", "")
    return show_qr, qr_content, qr_image_path


def resolve_qr_image_bytes(clinic: dict) -> Optional[bytes]:
    """Resolve the image bytes to embed for this clinic's QR, or None if
    show_qr is off / nothing usable is configured. Prefers a static
    pre-made image (qr_image_path, resolved relative to the repo root) over
    generating one from qr_content, since a real payment QR (e.g. Zelle's)
    encodes bank-specific data a generator can't reproduce from a URL."""
    show_qr, qr_content, qr_image_path = qr_settings(clinic)
    if not show_qr:
        return None
    if qr_image_path:
        image_path = _REPO_ROOT / qr_image_path
        if image_path.is_file():
            return image_path.read_bytes()
    if qr_content:
        return generate_qr_png_bytes(qr_content)
    return None
