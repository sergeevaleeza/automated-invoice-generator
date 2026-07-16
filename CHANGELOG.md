# Changelog

## [Unreleased] — 2026-07-16

### Added

#### 1. Excel (.xlsx) invoice export, alongside PDF
- Added `excel_invoice_generator.py`, a new module using `openpyxl` that generates a print-ready `.xlsx` invoice visually mirroring the PDF layout: clinic header (with conditional EIN/NPI line), contact info block, patient info + payment instructions, `PATIENT STATEMENT` title, statement/payment-due dates, the line-item table (Service Date(s), Description, Amount Paid, Copay/Deductible), previous-balance/credit rows, SUBTOTAL/TOTAL rows, the "YOUR PORTION DUE" section, provider signature line, and the two-line footer.
- All fonts, labels, column widths, margins, and the layout-tier scale live in a single `CONFIG` dict at the top of the module for easy adjustment.
- `generate_excel_invoice(patient, lines, total_due, patient_df, statement_date, payment_due_date, has_cpt, output_path)` takes the exact same inputs already assembled for `_generate_pdf_invoice()` (same `lines`/`total_due`/`patient_df` from `_generate_invoice_lines()`, same `has_cpt` from `_has_cpt_codes()`) — no billing logic (amount-due calculation, previous-balance/credit classification, zero-balance vs. skip handling) is duplicated; this module is presentation-only.
- Print setup is US Letter, portrait, `fitToWidth=1`/`fitToHeight=1` (forced single page, matching the PDF's single-page compression behavior) with margins (0.65/0.65/0.4/0.6 in) matching the PDF's margins, explicit column widths/row heights (no autofit), currency number format (`$#,##0.00`, with a literal `"$ -"` for zero — matching the PDF's `$ -` placeholder), grid borders on the item table, and an explicit print area set to the used range.
- Uses the same 3-tier font/row-height scaling approach as the PDF's `_LAYOUT_TIERS`: the sheet is built tier-by-tier, measuring actual accumulated row height, until one fits a single printed page. If even the smallest tier doesn't fit (very long item lists), falls back to `fitToHeight=0` (multi-page) with the table header row set to repeat on every page via `print_title_rows`.

#### 2. Invoice export format selector in the UI
- Added a radio selector ("PDF only" / "Excel only" / "Both PDF & Excel") on the "Generate Reports" tab, directly above the "Generate All Reports" button. Defaults to "PDF only" to preserve existing behavior.
- `generate_invoices()` gained an `export_format: str = "pdf"` parameter (`"pdf"`, `"excel"`, or `"both"`) controlling which invoice file(s) are written per patient. The existing generic zip-download logic (which already zips everything under `output_dir`) needed no changes — selecting "Both" naturally bundles PDF and Excel invoices together in the same download. Filenames share the same base name across formats (e.g. `LastName_2026_Invoice_mmddyyyy.pdf` / `.xlsx`).

#### 3. Shared invoice data structures extracted to `invoice_models.py`
- Moved `PatientData`, `InvoiceLine`, `ProcessingSummary`, and the date-formatting helper (`_format_date_for_display`, now a module-level `format_date_for_display()`) out of `complete_patient_invoice_generator.py` into a new `invoice_models.py`, so both the PDF and Excel generators import the same definitions instead of duplicating them. `PatientInvoiceGenerator._format_date_for_display()` is now a thin wrapper around the shared function; behavior is unchanged.

#### 4. Bundled default cover letter template, with optional override
- Added `templates/Access_Multi_Letter_Cover.docx`, loaded automatically on app start instead of requiring a re-upload every session. (Bundled here as a clearly-labeled placeholder with the correct merge placeholders — replace with the real letter content, or use "Save as new default" below.)
- The uploader in the "Upload Files" tab is now optional, moved into a "Replace cover letter template (optional)" expander. An uploaded file overrides the bundled default for the session; the active template (bundled vs. uploaded) is shown via `st.info`. If the bundled default is missing, the app shows a warning and asks for an upload instead of crashing.
- Added a "Save as new default template" button (behind a confirmation checkbox) that overwrites `templates/Access_Multi_Letter_Cover.docx` with the uploaded file. UI text notes this won't survive a Streamlit Cloud redeploy and the file should also be committed to the repo.
- Added `TEMPLATE_CONFIG` (top-level dict in `invoice_app.py`) holding the default template path and required placeholder list.

#### 5. Cover letter template validation
- Added `REQUIRED_TEMPLATE_PLACEHOLDERS` (the 9 merge placeholders `_generate_cover_letter()` fills in) and `validate_cover_letter_template()` to `invoice_models.py` — opens a `.docx` (path or uploaded file object) with `python-docx` and reports which required placeholders are missing, scanning both paragraphs and table cells. Used to validate both the bundled default and any uploaded override; missing placeholders show as a warning without blocking generation.
- `_generate_cover_letter()`'s `replacements` dict is now built from `REQUIRED_TEMPLATE_PLACEHOLDERS` (`dict(zip(...))`) instead of a separately hand-typed literal, so the generator and the validator can't drift apart.

### Fixed

#### 6. Excel invoice — payment notice text was cut off
- The payment-instructions box (merged cell, columns C:D) sized its row span from the raw count of text items (4), not from how many visual lines Excel would actually wrap them into. Because the column width is narrower than several of the instruction lines (e.g. the 53-character Zelle line), Excel wrapped them into 2 lines each — needing ~8 rows of height, not 4 — and the last 1-2 lines were clipped. Excel does not auto-grow a merged cell's row height for wrapped text, so this required an explicit fix rather than just "turning on autofit".
- Added `_count_wrapped_lines()` / `_estimate_chars_per_line()` to `excel_invoice_generator.py`, which estimate the wrapped line count from text length, column width, and font size, and size the merged block's row span from that estimate instead of the raw item count.

#### 7. Excel invoice — bottom boxes rendered without a right border in some viewers
- The "YOUR PORTION DUE" and "AMOUNT ENCLOSED" boxes only had their border set on the top-left cell of each merged range. openpyxl 3.1.5 happens to auto-propagate that border to the rest of the merge at save time, but this is an implicit, version-dependent behavior other spreadsheet applications (LibreOffice, Google Sheets) aren't guaranteed to replicate the same way.
- Added a reusable `apply_box_border(ws, min_row, min_col, max_row, max_col)` helper that explicitly sets only the true perimeter edge on every boundary cell in a range (no relying on merge-copy behavior, no stray internal divider line). Used it for the payment-notice box (which also gained a border it didn't have before, per the reported "bordered cell block" description), both bottom amount boxes (now a single unified box per label+value pair instead of two independently-bordered rows), and defensively on the line-item table's outer perimeter.

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
- Added `_has_cpt_codes()` method that checks whether any value in the `type_of_service` column contains a 5-digit CPT code.
- When CPT codes are detected, a fourth line `EIN: 94-3368586    NPI: 1245365782` is appended to the three-line clinic header. When no CPT codes are present the header is unchanged.

#### 4a. Bugfix — CPT detection missed codes embedded in descriptions
- The initial regex `r'^\d{5}$'` required the entire cell value to be exactly 5 digits, so descriptions like `"Med Management (CPT Code 99213)"` were not detected and the EIN/NPI line was never added.
- Changed to `re.search(r'\b\d{5}\b', ...)`, which finds a 5-digit number anywhere in the string while `\b` word boundaries still prevent false matches from 4- or 6-digit numbers.

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
- `_has_cpt_codes(patient_df)` — detects 5-digit CPT codes in `type_of_service` column using `re.search(r'\b\d{5}\b', ...)`, matching both bare codes (`90837`) and codes embedded in descriptions (`Med Management (CPT Code 99213)`).
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
