# Changelog

## [Unreleased] — 2026-06-26

### Changed

#### 1. Envelope files — DOCX only, no date suffix
- Removed all PDF envelope generation from `generate_invoices()`. `_generate_envelope_pdf()` is retained in the codebase but is no longer called.
- Renamed DOCX envelope output from `LastName_Year_Envelope_mmddyyyy.docx` to `LastName_Envelope.docx`. The absence of a date suffix means each run overwrites the previous copy in the patient folder, ensuring only one current envelope exists per patient.

#### 2. PDF header — reduced font size
- Reduced the main clinic header font size from 12pt to 10pt (base tier). The three header lines — `ACCESS MULTI-SPECIALTY MEDICAL CLINIC, INC.`, `MICHAEL U. LEVINSON, MD, PH D.`, and `BOARD CERTIFIED PSYCHIATRIST` — now render at 10pt by default.

#### 3. PDF header — added WEBSITE line
- Added `WEBSITE: https://accessmultispecialty.com/` as a fourth line in the contact info block, below the existing EMAIL line.

#### 4. Conditional EIN and NPI in header
- Added `_has_cpt_codes()` method that checks whether any value in the `type_of_service` column matches the regex `r'^\d{5}$'` (exactly 5 digits).
- When CPT codes are detected, a fourth line `EIN: 94-3368586    NPI: 1245365782` is appended to the three-line clinic header. When no CPT codes are present the header is unchanged.

#### 5. Two-line footer
- `add_optimized_footer()` now draws two centered lines at the bottom of every page, both in Helvetica 8pt:
  - Line 1 at y = 0.55 in: `If you have questions regarding your bill, please contact us at (415)857-1151.`
  - Line 2 at y = 0.35 in: `For current pricing, please visit: https://accessmultispecialty.com/pricing.html`
- Removed the dynamic `footer_font` instance variable; font size is now fixed at 8pt.

#### 6. Negative previous balance support
- `InvoiceLine` dataclass gained a new `is_credit: bool = False` field.
- `_generate_invoice_lines()` now handles three cases for previous balance:
  - `> 0`: creates a normal `InvoiceLine` with `is_credit=False` (owed from prior period)
  - `< 0`: creates a credit `InvoiceLine` with `is_credit=True` and `description="Previous Balance (Overpaid)"`
  - `== 0`: no line created
- In the PDF table, positive previous balance appears in the **Copay/Deductible** column; a negative (overpaid) balance appears in the **Amount Paid** column as a credit.
- The `total_due` calculation (`max(0, sum(copay) + previous_balance - sum(paid))`) already correctly reduces the total when `previous_balance` is negative — no arithmetic change was needed.
- `_generate_csv_export()` updated to route credit lines into the `Amount` column instead of `Copay/Deductible`.

#### 7. Force single-page PDF
- Added `_LAYOUT_TIERS` class attribute: a list of 6 parameter dicts applied in order to compress content until it fits on one page.
- Added `_count_pdf_pages()` method that counts `/Type /Page` objects in raw PDF bytes (distinguishes individual pages from the `/Pages` catalog object).
- `_generate_pdf_invoice()` now builds the PDF into a `BytesIO` buffer, checks the page count, and advances to the next tier only if the content overflows. The first tier that produces a single page is written to disk; Tier 5 is written regardless as a last resort.
- Compression order across tiers:
  1. Reduce Spacer heights by ~25%
  2. Reduce table `TOPPADDING`/`BOTTOMPADDING` from 6pt to 3pt
  3. Reduce body font from 9pt to 8pt
  4. Reduce header font from 10pt to 9pt
  5. Reduce all spacers and padding further; reduce body font to 7pt
- Fixed document margins: `topMargin=0.4 in`, `bottomMargin=0.6 in`, `leftMargin=0.65 in`, `rightMargin=0.65 in`.

#### 8. Amount Paid column fix
- Fixed a display bug where the **Copay/Deductible** column previously showed the raw copay amount instead of the amount still owed.
- Per-row display is now:
  - **Amount Paid** = value from the `paid` column (`$ -` when zero)
  - **Copay/Deductible** = `max(0, copay − paid)` (`$ -` when zero)
- SUBTOTAL row sums each column independently.
- TOTAL row shows `total_due` as computed by `_generate_invoice_lines()`.

### Added
- `_has_cpt_codes(patient_df)` — detects 5-digit CPT codes in `type_of_service` column.
- `_count_pdf_pages(pdf_bytes)` — counts pages in a ReportLab-generated PDF from raw bytes.
- `_LAYOUT_TIERS` class attribute — six progressive layout configurations used for single-page fitting.
- `InvoiceLine.is_credit` field — distinguishes overpayment credits from standard previous-balance charges.
- `from io import BytesIO` import added at module level.

### Output folder structure (unchanged)
```
output/LastName_FirstName_PRN/
  LastName_Year_Invoice_mmddyyyy.pdf       ← date kept
  LastName_Envelope.docx                   ← no date (overwrites on each run)
  LastName_Year_InvoiceItems_mmddyyyy.csv  ← date kept
```
