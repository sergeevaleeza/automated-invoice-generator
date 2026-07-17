#!/usr/bin/env python3
"""
Superbill PDF generator — a separate, single-patient export containing the
standard fields a patient needs to seek out-of-network reimbursement:
provider identity (name, NPI, tax ID/EIN, practice address), patient
identity (name, DOB, address), diagnosis codes (ICD-10), and per-service
lines (date, CPT code, charge, payment).

Deliberately a clean letterhead-styled document, not a pixel-perfect
CMS-1500 form — CMS-1500 is a fixed-field form designed for direct payer
submission, and reproducing its exact box layout in ReportLab wouldn't be
"trivial" as the spec allows skipping. This covers the same information a
superbill needs to convey for patient self-submission.

Consumes PatientData + SuperbillServiceLine (from invoice_models.py) and a
loaded clinic config dict — the same shapes the main invoice generators
use — so no billing/resolution logic is duplicated here; this module is
presentation only. CPT/ICD-10 resolution lives in
PatientInvoiceGenerator.resolve_superbill_service_lines() /
resolve_default_icd10_codes().
"""
from datetime import datetime
from pathlib import Path
from typing import List

from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

from invoice_models import PatientData, SuperbillServiceLine

CONFIG = {
    "title": "SUPERBILL",
    "subtitle": "Statement for Insurance Reimbursement",
    "table_headers": ["Date", "CPT Code", "Description", "Charge", "Payment"],
    "font_family": "Helvetica",
    "margins": dict(top=0.6 * inch, bottom=0.6 * inch, left=0.65 * inch, right=0.65 * inch),
}


def _clean_postal_code(postal_code: str) -> str:
    if postal_code and '.' in postal_code:
        return postal_code.split('.')[0]
    return postal_code


def generate_superbill_pdf(patient: PatientData, clinic: dict, service_lines: List[SuperbillServiceLine],
                            icd10_codes: List[str], statement_date: datetime, output_path: Path) -> None:
    """Generate a single-patient superbill PDF. Totals are a plain sum of
    the given service_lines — not recomputed via any amount-due strategy —
    since a superbill reports what was charged/paid per line, not an
    "amount due" balance."""
    doc = SimpleDocTemplate(
        str(output_path), pagesize=letter,
        topMargin=CONFIG["margins"]["top"], bottomMargin=CONFIG["margins"]["bottom"],
        leftMargin=CONFIG["margins"]["left"], rightMargin=CONFIG["margins"]["right"],
    )
    styles = getSampleStyleSheet()
    story = []

    header_style = ParagraphStyle(
        'SBHeader', parent=styles['Normal'], fontSize=12, alignment=TA_CENTER,
        spaceAfter=2, fontName=f"{CONFIG['font_family']}-Bold",
    )
    header2_style = ParagraphStyle(
        'SBHeader2', parent=styles['Normal'], fontSize=9, alignment=TA_CENTER,
        spaceAfter=2, fontName=CONFIG['font_family'],
    )
    title_style = ParagraphStyle(
        'SBTitle', parent=styles['Normal'], fontSize=14, alignment=TA_CENTER,
        spaceAfter=2, fontName=f"{CONFIG['font_family']}-Bold",
    )
    subtitle_style = ParagraphStyle(
        'SBSubtitle', parent=styles['Normal'], fontSize=10, alignment=TA_CENTER,
        spaceAfter=14, fontName=CONFIG['font_family'],
    )
    label_style = ParagraphStyle(
        'SBLabel', parent=styles['Normal'], fontSize=9, alignment=TA_LEFT,
        fontName=f"{CONFIG['font_family']}-Bold", leading=13,
    )
    body_style = ParagraphStyle(
        'SBBody', parent=styles['Normal'], fontSize=9, alignment=TA_LEFT,
        fontName=CONFIG['font_family'], leading=13,
    )

    # --- Practice letterhead ---
    story.append(Paragraph(clinic['clinic_name'], header_style))
    story.append(Paragraph(clinic['doctor_name'], header_style))
    story.append(Paragraph(clinic['specialty'], header_style))
    story.append(Paragraph(clinic['office_address'], header2_style))
    story.append(Spacer(1, 10))
    story.append(Paragraph(CONFIG["title"], title_style))
    story.append(Paragraph(CONFIG["subtitle"], subtitle_style))

    # --- Provider / patient info, side by side ---
    display_postal = _clean_postal_code(patient.postal_code)
    patient_address = f"{patient.address_line1}"
    if patient.address_line2:
        patient_address += f", {patient.address_line2}"
    patient_address += f"<br/>{patient.city}, {patient.state} {display_postal}"

    provider_info = (
        f"<b>Provider:</b> {clinic['doctor_name']}<br/>"
        f"<b>NPI:</b> {clinic['npi']}<br/>"
        f"<b>Tax ID (EIN):</b> {clinic['ein']}<br/>"
        f"<b>Practice Address:</b> {clinic['office_address']}"
    )
    patient_info = (
        f"<b>Patient:</b> {patient.first_name} {patient.last_name}<br/>"
        f"<b>DOB:</b> {patient.dob or 'N/A'}<br/>"
        f"<b>Address:</b> {patient_address}<br/>"
        f"<b>PRN:</b> {patient.prn or 'N/A'}"
    )
    info_table = Table(
        [[Paragraph(provider_info, body_style), Paragraph(patient_info, body_style)]],
        colWidths=[3.35 * inch, 3.35 * inch],
    )
    info_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
        ('BOX', (0, 0), (0, 0), 0.5, colors.grey),
        ('BOX', (1, 0), (1, 0), 0.5, colors.grey),
        ('LEFTPADDING', (1, 0), (1, 0), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]))
    story.append(info_table)
    story.append(Spacer(1, 12))

    # --- Diagnosis codes ---
    icd_text = ", ".join(icd10_codes) if icd10_codes else "Not specified"
    story.append(Paragraph(f"<b>Diagnosis Code(s) (ICD-10):</b> {icd_text}", label_style))
    story.append(Spacer(1, 4))
    story.append(Paragraph(f"<b>Statement Date:</b> {statement_date.strftime('%m/%d/%Y')}", label_style))
    story.append(Spacer(1, 12))

    # --- Service lines table ---
    table_data = [CONFIG["table_headers"]]
    total_charge = 0.0
    total_payment = 0.0
    for line in service_lines:
        table_data.append([
            line.service_date, line.cpt_code or "—", line.description,
            f"${line.charge:.2f}", f"${line.payment:.2f}",
        ])
        total_charge += line.charge
        total_payment += line.payment
    table_data.append(["", "", "TOTAL", f"${total_charge:.2f}", f"${total_payment:.2f}"])

    service_table = Table(table_data, colWidths=[0.9 * inch, 0.8 * inch, 2.9 * inch, 1 * inch, 1 * inch])
    service_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('FONTNAME', (0, 0), (-1, 0), f"{CONFIG['font_family']}-Bold"),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('ALIGN', (0, 0), (2, -1), 'LEFT'),
        ('ALIGN', (3, 0), (4, -1), 'RIGHT'),
        ('FONTNAME', (0, -1), (-1, -1), f"{CONFIG['font_family']}-Bold"),
        ('LINEABOVE', (0, -1), (-1, -1), 1, colors.black),
        ('GRID', (0, 0), (-1, -2), 0.5, colors.black),
        ('BOX', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]))
    story.append(service_table)
    story.append(Spacer(1, 24))

    # --- Signature ---
    sig_style = ParagraphStyle(
        'SBSig', parent=styles['Normal'], fontSize=10, alignment=TA_RIGHT,
        fontName=f"{CONFIG['font_family']}-Bold",
    )
    story.append(Paragraph("_________________________________", sig_style))
    story.append(Spacer(1, 6))
    story.append(Paragraph(f"Provider Signature - {clinic['provider_name_for_signature']}", sig_style))

    doc.build(story)
