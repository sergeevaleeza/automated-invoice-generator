#!/usr/bin/env python3
"""
Excel Invoice Generator
Produces a print-ready .xlsx patient invoice that visually mirrors the PDF
invoice built by PatientInvoiceGenerator._generate_pdf_invoice(). Consumes the
exact same inputs (PatientData, InvoiceLine list, total_due, the per-patient
DataFrame, statement/payment-due dates, and the has_cpt flag) so no billing
business logic is duplicated here — this module is presentation only.
"""
import math
from pathlib import Path
from datetime import datetime
from typing import List, Tuple

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter

from invoice_models import PatientData, InvoiceLine, format_date_for_display

# --- Configuration block: fonts, labels, widths, margins, layout tiers -----
CONFIG = {
    "clinic_name": "ACCESS MULTI-SPECIALTY MEDICAL CLINIC, INC.",
    "doctor_name": "MICHAEL U. LEVINSON, MD, PH D.",
    "specialty": "BOARD CERTIFIED PSYCHIATRIST",
    "ein_npi": "EIN: 94-3368586    NPI: 1245365782",
    "office_address": "OFFICE ADDRESS: 25 EDWARDS COURT, SUITE 101, BURLINGAME, CA 94010",
    "mailing_address": "MAILING ADDRESS: PO BOX 351, BURLINGAME, CA 94011",
    "email": "EMAIL: ACCESS.MSMC@GMAIL.COM",
    "website": "WEBSITE: https://accessmultispecialty.com/",
    "payment_instructions": [
        "Please note we do not accept credit cards.",
        "1. Zelle access.msmc@gmail.com (IRA Billing and Mgmt)",
        "2. Check payable to: Michael Levinson, MD",
        "   PO Box 351, Burlingame, CA 94011",
    ],
    "statement_title": "PATIENT STATEMENT",
    "table_headers": ["Service Date(s)", "Description", "Amount Paid", "Copay/Deductible"],
    "amount_due_label": "YOUR PORTION DUE:",
    "amount_enclosed_label": "AMOUNT ENCLOSED:",
    "signature_label": "Provider Signature - Michael Levinson, MD",
    "footer_line1": "If you have questions regarding your bill, please contact us at (415)857-1151.",
    "footer_line2": "For current pricing, please visit: https://accessmultispecialty.com/pricing.html",
    "default_service_description": "Psychotherapy and/or Med Management",

    "font_family": "Arial",
    "col_widths": [15, 42, 15, 17],  # Service Date, Description, Amount Paid, Copay/Deductible
    "header_fill_color": "D9D9D9",

    "margins": dict(left=0.65, right=0.65, top=0.4, bottom=0.6, header=0.2, footer=0.2),

    # Layout tiers applied in order (mirrors the PDF's _LAYOUT_TIERS approach)
    # until the accumulated row height fits one printed Letter page.
    "tiers": [
        dict(font_header=12, font_header2=10, font_title=13, font_body=10, font_table=9,
             row_h=15, header_row_h=18, spacer_h=6),
        dict(font_header=11, font_header2=9, font_title=12, font_body=9, font_table=8,
             row_h=13, header_row_h=16, spacer_h=5),
        dict(font_header=10, font_header2=8, font_title=11, font_body=8, font_table=7,
             row_h=11, header_row_h=14, spacer_h=4),
    ],
}

_LETTER_HEIGHT_IN = 11.0
CONFIG["usable_height_pt"] = (
    _LETTER_HEIGHT_IN - CONFIG["margins"]["top"] - CONFIG["margins"]["bottom"]
) * 72

CURRENCY_FMT = '$#,##0.00;-$#,##0.00;"$ -"'
THIN = Side(style="thin", color="000000")
CELL_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _clean_postal_code(postal_code: str) -> str:
    if postal_code and '.' in postal_code:
        return postal_code.split('.')[0]
    return postal_code


def _estimate_chars_per_line(col_width_units: float, font_size: int) -> int:
    """Estimate how many characters fit on one wrapped line inside a merged
    range of the given total column width (Excel width units) and font size.
    An Excel width unit is roughly one character of the workbook's default
    ~11pt font; scale for the actual font size and subtract a small margin
    for cell padding."""
    base_chars = max(col_width_units - 2, 1)
    scale = 11.0 / max(font_size, 6)
    return max(int(base_chars * scale), 8)


def _count_wrapped_lines(text_lines: List[str], col_width_units: float, font_size: int) -> int:
    """Estimate the total visual (wrapped) line count for a list of logical
    text lines once Excel wraps them inside a merged cell of the given width.
    Excel does not auto-grow merged-cell row heights for wrapped text, so
    row spans must be sized from this estimate rather than the raw line
    count, or trailing lines get visually clipped."""
    chars_per_line = _estimate_chars_per_line(col_width_units, font_size)
    total = 0
    for line in text_lines:
        if not line:
            total += 1
            continue
        total += max(1, math.ceil(len(line) / chars_per_line))
    return total


def apply_box_border(ws: Worksheet, min_row: int, min_col: int, max_row: int, max_col: int,
                      style: str = "thin") -> None:
    """Draw a single outline box around a cell range (merged or not),
    setting only the true perimeter edge on each boundary cell — no internal
    divider lines. Applied explicitly to every cell in the range rather than
    relying on openpyxl's implicit top-left-cell border propagation for
    merges, which other spreadsheet applications may not honor."""
    side = Side(style=style, color="000000")
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            cell = ws.cell(row=r, column=c)
            existing = cell.border
            cell.border = Border(
                left=side if c == min_col else existing.left,
                right=side if c == max_col else existing.right,
                top=side if r == min_row else existing.top,
                bottom=side if r == max_row else existing.bottom,
            )


def _build_workbook(patient: PatientData, total_due: float, patient_df: pd.DataFrame,
                     statement_date: datetime, payment_due_date: datetime,
                     has_cpt: bool, tier: dict) -> Tuple[Workbook, Worksheet, float, int, int]:
    """Build one full attempt of the invoice at a given layout tier.

    Returns (workbook, worksheet, accumulated_row_height_pt, header_row_index, last_row_index).
    """
    cfg = CONFIG
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoice"

    height_accum = 0.0
    row = 1

    def set_row_height(r: int, h: float):
        nonlocal height_accum
        ws.row_dimensions[r].height = h
        height_accum += h

    def merged(r1, c1, r2, c2, value, font, alignment):
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
        cell = ws.cell(row=r1, column=c1, value=value)
        cell.font = font
        cell.alignment = alignment
        return cell

    header_font = Font(name=cfg["font_family"], size=tier["font_header"], bold=True)
    header2_font = Font(name=cfg["font_family"], size=tier["font_header2"], bold=True)
    title_font = Font(name=cfg["font_family"], size=tier["font_title"], bold=True)
    body_font = Font(name=cfg["font_family"], size=tier["font_body"])
    body_bold_font = Font(name=cfg["font_family"], size=tier["font_body"], bold=True)
    table_font = Font(name=cfg["font_family"], size=tier["font_table"])
    table_bold_font = Font(name=cfg["font_family"], size=tier["font_body"], bold=True)
    sig_font = Font(name=cfg["font_family"], size=tier["font_body"] + 1, bold=True)

    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="top", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center")

    # --- Clinic header block ---
    for text in (cfg["clinic_name"], cfg["doctor_name"], cfg["specialty"]):
        merged(row, 1, row, 4, text, header_font, center)
        set_row_height(row, tier["header_row_h"])
        row += 1
    if has_cpt:
        merged(row, 1, row, 4, cfg["ein_npi"], header_font, center)
        set_row_height(row, tier["header_row_h"])
        row += 1
    set_row_height(row, tier["spacer_h"])
    row += 1

    # --- Contact info block ---
    for text in (cfg["office_address"], cfg["mailing_address"], cfg["email"], cfg["website"]):
        merged(row, 1, row, 4, text, header2_font, center)
        set_row_height(row, tier["header_row_h"] - 2)
        row += 1
    set_row_height(row, tier["spacer_h"] * 2)
    row += 1

    # --- Patient info + payment instructions side by side ---
    display_postal = _clean_postal_code(patient.postal_code)
    patient_lines = [f"{patient.first_name.upper()} {patient.last_name.upper()}"]
    if patient.address_line1:
        patient_lines.append(patient.address_line1)
    if patient.address_line2:
        patient_lines.append(patient.address_line2)
    patient_lines.append(f"{patient.city}, {patient.state} {display_postal}")
    patient_text = "\n".join(patient_lines)
    payment_text = "\n".join(cfg["payment_instructions"])

    patient_col_width = cfg["col_widths"][0] + cfg["col_widths"][1]
    payment_col_width = cfg["col_widths"][2] + cfg["col_widths"][3]
    wrapped_patient_lines = _count_wrapped_lines(patient_lines, patient_col_width, tier["font_body"])
    wrapped_payment_lines = _count_wrapped_lines(cfg["payment_instructions"], payment_col_width, tier["font_body"])

    info_row_start = row
    n_info_lines = max(wrapped_patient_lines, wrapped_payment_lines)
    info_row_end = info_row_start + n_info_lines - 1

    p_cell = merged(info_row_start, 1, info_row_end, 2, patient_text, body_bold_font, left)
    pay_cell = merged(info_row_start, 3, info_row_end, 4, payment_text, body_bold_font, left)
    apply_box_border(ws, info_row_start, 3, info_row_end, 4)
    for r in range(info_row_start, info_row_end + 1):
        set_row_height(r, tier["row_h"])
    row = info_row_end + 1
    set_row_height(row, tier["spacer_h"] * 2)
    row += 1

    # --- Title ---
    merged(row, 1, row, 4, cfg["statement_title"], title_font, center)
    set_row_height(row, tier["header_row_h"] + 2)
    row += 1
    set_row_height(row, tier["spacer_h"])
    row += 1

    # --- Statement date table ---
    merged(row, 1, row, 2, "STATEMENT DATE:", body_bold_font, Alignment(horizontal="left", vertical="center"))
    merged(row, 3, row, 4, "Payment due date:", body_bold_font, Alignment(horizontal="left", vertical="center"))
    set_row_height(row, tier["row_h"])
    row += 1
    merged(row, 1, row, 2, statement_date.strftime('%m/%d/%Y'), body_bold_font, Alignment(horizontal="left", vertical="center"))
    merged(row, 3, row, 4, payment_due_date.strftime('%m/%d/%Y'), body_bold_font, Alignment(horizontal="left", vertical="center"))
    set_row_height(row, tier["row_h"])
    row += 1
    set_row_height(row, tier["spacer_h"] * 2)
    row += 1

    # --- Line-item table ---
    header_row_idx = row
    for col_idx, text in enumerate(cfg["table_headers"], start=1):
        cell = ws.cell(row=row, column=col_idx, value=text)
        cell.font = table_bold_font
        cell.alignment = Alignment(horizontal="left" if col_idx <= 2 else "right", vertical="center")
        cell.fill = PatternFill("solid", fgColor=cfg["header_fill_color"])
        cell.border = CELL_BORDER
    set_row_height(row, tier["header_row_h"])
    row += 1

    def add_item_row(date_val, desc, paid_val, copay_val, bold=False, top_border_style="thin"):
        nonlocal row
        font = table_bold_font if bold else table_font
        border = Border(left=THIN, right=THIN, bottom=THIN, top=Side(style=top_border_style, color="000000"))
        c1 = ws.cell(row=row, column=1, value=date_val)
        c2 = ws.cell(row=row, column=2, value=desc)
        c3 = ws.cell(row=row, column=3, value=paid_val)
        c4 = ws.cell(row=row, column=4, value=copay_val)
        for c in (c1, c2, c3, c4):
            c.font = font
            c.border = border
        c1.alignment = Alignment(horizontal="left", vertical="center")
        c2.alignment = Alignment(horizontal="left", vertical="center")
        c3.alignment = right
        c4.alignment = right
        if paid_val is not None:
            c3.number_format = CURRENCY_FMT
        if copay_val is not None:
            c4.number_format = CURRENCY_FMT
        set_row_height(row, tier["row_h"])
        row += 1
        return (c1, c2, c3, c4)

    previous_balance = float(patient_df.iloc[0].get('previous_balance', 0))
    total_paid_display = 0.0
    total_copay_display = 0.0

    if previous_balance > 0:
        add_item_row("", "Previous Balance", 0.0, previous_balance)
        total_copay_display += previous_balance
    elif previous_balance < 0:
        add_item_row("", "Previous Balance (Overpaid)", abs(previous_balance), 0.0)
        total_paid_display += abs(previous_balance)

    for _, prow in patient_df.iterrows():
        paid_amount = float(prow.get('paid', 0))
        copay_amount = float(prow.get('copay', 0))
        if paid_amount > 0 or copay_amount > 0:
            display_date = format_date_for_display(prow['visit_date'])
            service_type = str(prow.get('type_of_service', '')).strip() or cfg["default_service_description"]
            add_item_row(display_date, service_type, paid_amount, copay_amount)
            total_paid_display += paid_amount
            total_copay_display += copay_amount

    add_item_row("", "SUBTOTAL", total_paid_display, total_copay_display, bold=True, top_border_style="medium")
    add_item_row("", "TOTAL", None, total_due, bold=True, top_border_style="double")
    total_row_idx = row - 1
    apply_box_border(ws, header_row_idx, 1, total_row_idx, 4)

    set_row_height(row, tier["spacer_h"] * 2)
    row += 1

    # --- Amount due section ---
    label_align = Alignment(horizontal="left", vertical="center")
    amount_box_row_start = row
    l1 = merged(row, 1, row, 2, cfg["amount_due_label"], body_bold_font, label_align)
    l2 = merged(row, 3, row, 4, cfg["amount_enclosed_label"], body_bold_font, label_align)
    set_row_height(row, tier["row_h"] + 2)
    row += 1
    v1 = merged(row, 1, row, 2, total_due, body_bold_font, label_align)
    v1.number_format = CURRENCY_FMT
    v2 = merged(row, 3, row, 4, "", body_bold_font, label_align)
    set_row_height(row, tier["row_h"] + 2)
    amount_box_row_end = row
    row += 1
    apply_box_border(ws, amount_box_row_start, 1, amount_box_row_end, 2)
    apply_box_border(ws, amount_box_row_start, 3, amount_box_row_end, 4)
    set_row_height(row, tier["spacer_h"] * 3)
    row += 1

    # --- Provider signature ---
    merged(row, 1, row, 4, "_________________________________", sig_font, right)
    set_row_height(row, tier["row_h"])
    row += 1
    merged(row, 1, row, 4, cfg["signature_label"], sig_font, right)
    set_row_height(row, tier["row_h"])
    row += 1

    last_row = row - 1
    return wb, ws, height_accum, header_row_idx, last_row


def _apply_page_setup(ws: Worksheet, last_row: int, header_row_idx: int, fits_one_page: bool):
    cfg = CONFIG
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1 if fits_one_page else 0
    ws.page_margins = PageMargins(**cfg["margins"])
    ws.print_area = f"A1:D{last_row}"
    ws.print_title_rows = f"{header_row_idx}:{header_row_idx}"
    ws.sheet_view.showGridLines = False

    for i, width in enumerate(cfg["col_widths"], start=1):
        ws.column_dimensions[get_column_letter(i)].width = width

    ws.oddFooter.center.text = f"{cfg['footer_line1']}\n{cfg['footer_line2']}"
    ws.oddFooter.center.size = 8
    ws.oddFooter.center.font = "Arial,Regular"


def generate_excel_invoice(patient: PatientData, lines: List[InvoiceLine], total_due: float,
                            patient_df: pd.DataFrame, statement_date: datetime,
                            payment_due_date: datetime, has_cpt: bool, output_path: Path) -> None:
    """Generate a print-ready Excel invoice mirroring the PDF invoice layout.

    Takes the same inputs already assembled for _generate_pdf_invoice (lines,
    total_due, patient_df, has_cpt) — no billing logic is recomputed here.
    """
    tiers = CONFIG["tiers"]
    usable_height = CONFIG["usable_height_pt"]

    chosen = None
    for tier in tiers:
        wb, ws, height, header_row_idx, last_row = _build_workbook(
            patient, total_due, patient_df, statement_date, payment_due_date, has_cpt, tier
        )
        chosen = (wb, ws, height, header_row_idx, last_row)
        if height <= usable_height:
            break

    wb, ws, height, header_row_idx, last_row = chosen
    fits_one_page = height <= usable_height
    _apply_page_setup(ws, last_row, header_row_idx, fits_one_page)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(output_path))
