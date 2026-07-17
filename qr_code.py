#!/usr/bin/env python3
"""
Shared QR code generation for invoices. Presentation-agnostic (returns PNG
bytes) so both the PDF (ReportLab) and Excel (openpyxl) generators can embed
the same image without duplicating QR-building logic.

Controlled by two optional clinic_config.json fields (not required — a
config without them behaves as if show_qr were false):
  "show_qr": true/false
  "qr_content": "https://..."  (falls back to the clinic's "website" if omitted)
"""
import io

import qrcode


def generate_qr_png_bytes(content: str, box_size: int = 10, border: int = 2) -> bytes:
    """Generate a QR code encoding `content`, returned as PNG bytes."""
    img = qrcode.make(content, box_size=box_size, border=border)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def qr_settings(clinic: dict) -> tuple:
    """Resolve (show_qr, qr_content) from a loaded clinic config dict.
    Both fields are optional — a config that predates this feature is
    treated as show_qr=False, not an error."""
    show_qr = bool(clinic.get("show_qr", False))
    qr_content = clinic.get("qr_content") or clinic.get("website", "")
    return show_qr, qr_content
