# Changelog

## [Unreleased] — Phase 1 item 1: HIPAA/open-source hygiene sweep (2026-07-16)

### Security / Findings

#### 1. Repo/history audit — no additional real data found
- Checked every `.xlsx`/`.xls`/`.csv`/`.docx` ever added across all branches (`git log --all --diff-filter=A`): only two files have ever existed in this repo's history — `templates/Access_Multi_Letter_Cover.docx` and `tests/fixtures/Example_2026_Invoice_07162026.xlsx` — both already accounted for (the latter is the real-PHI incident from the previous entry, already remediated). No other data files, in history or the current tree.
- Purged the previous incident's dangling git objects from local storage too (`git reflog expire --expire=now --all && git gc --prune=now --aggressive`) as defense in depth, on top of the earlier remote-side fix. Confirmed zero unreachable objects remain.

#### 2. PHI removed from application logs and UI exception output
- `self.logger.*()` calls in `complete_patient_invoice_generator.py` no longer include patient names or full file paths (paths contain the patient's last name via the folder-naming scheme) — found leaking in: a full roster-row dump on load, fuzzy-match logging (patient names + PRN), ambiguous/credit-balance/zero-balance/error logging per patient, and "Generated X" confirmations that included the identifying output path. PRN is kept in the fuzzy-match log line (useful for debugging, materially less sensitive than name+DOB+address) — everything else identifying was dropped or genericized.
- Removed `st.exception(e)` from `invoice_app.py`'s error handler — a full traceback can echo patient data from local variables in the call stack; `st.error(str(e))` (already present) is kept for basic diagnosis.
- Deliberately did NOT touch the "Successfully Processed/Skipped/Errors" patient lists shown in the UI or the `Processing_Summary_*.txt` file — showing staff which patients were processed, by name, is the app's core authorized function, not a leak. The distinction drawn here is server-side application logs (potentially retained/visible outside the practice, e.g. via hosting-provider log aggregation) vs. the in-app results the staff generating these specific invoices are already authorized to see.
- Confirmed no filenames are written outside `output_dir` containing patient data (uploaded files are saved under generic names in a temp dir; the only root-level file in `output_dir` is the genericly-named summary report).

### Added

#### 3. Clinic identity moved to `clinic_config.json` (gitignored) + example template
- Added `clinic_config.py`: `load_clinic_config()` reads `clinic_config.json`, validates all required fields are present, and raises `ClinicConfigError` (a message safe to show directly in the UI — never contains patient data) if the file is missing or incomplete. Deliberately does **not** silently fall back to placeholder data on its own — generating a real-looking invoice with fake clinic identity baked in would be worse than failing loudly.
- Added `clinic_config.example.json` (committed) as the placeholder template, and a local, gitignored `clinic_config.json` populated with the practice's real values so the running app is unaffected.
- `excel_invoice_generator.py`: split the old `CONFIG` dict into `LAYOUT_CONFIG` (static: fonts, widths, margins, row heights, non-identity labels) and `_clinic_derived_config(clinic)` (builds the identity-derived display strings — header lines, payment-notice text, signature label, footer — from a loaded clinic dict). `generate_excel_invoice()` now takes an optional `clinic` parameter (defaults to `load_clinic_config()`) so callers/tests can inject a config instead of depending on the real file.
- `complete_patient_invoice_generator.py`: `PatientInvoiceGenerator.__init__()` now loads clinic config into `self.clinic` (also injectable via a new `clinic_config` constructor param) and every hardcoded clinic string in `_generate_pdf_invoice()`, `add_optimized_footer()`, and the unused-but-still-public `_generate_envelope_pdf()` now reads from it instead.
- `invoice_app.py`: added a proactive check on the "Generate Reports" tab that shows a clear `st.error()` (reusing `ClinicConfigError`'s message) and blocks generation if `clinic_config.json` is missing/invalid, mirroring the existing cover-letter-template pattern.
- Tests never touch the real `clinic_config.json` — `tests/conftest.py` loads `clinic_config.example.json` instead and injects it explicitly, so the suite is fully self-contained and runs in a fresh clone/CI with no real business data present (verified by temporarily removing the local `clinic_config.json` and re-running the suite). The golden fixture was regenerated using the placeholder identity, so it no longer shows the real clinic's name — a deliberate, correct change, not a regression.

#### 4. `.gitignore` added (none existed before)
- Covers: `clinic_config.json`; `data/` and `output/` (runtime PHI); `*.xlsx`/`*.xls`/`*.csv`/`*.docx` everywhere, with explicit negations for the three files that should stay tracked (`clinic_config.example.json`, `templates/**/*.docx`, `tests/fixtures/**/*.xlsx`/`*.csv`/`*.docx`); `__pycache__/`, `*.pyc`/`*.pyo`; `.pytest_cache/`; common venv/editor cruft.
- Untracked the 7 `.pyc` files that were previously committed (`git rm -r --cached __pycache__ tests/__pycache__`) — they'll regenerate locally as needed and no longer show up as noise in every diff.

## [Unreleased] — Phase 0: Excel layout match (2026-07-16)

### Security

#### 0. Real patient data removed from git history
- A commit (`a7f35ba`) briefly added a test fixture containing a real patient's name and address instead of synthetic data, on the public `main` branch. Remediated same-day: repo set to private, the offending commit dropped from `main` (it was the branch tip, so no history rebase was needed — a clean reset), and the fixture recommitted with fabricated placeholder data. See `tests/conftest.py` for the synthetic-only policy going forward. If you're reading this after a `git clone`, the removed commit is not reachable from any branch.

### Changed

#### 1. Excel invoice layout rewritten to match the approved fixture exactly
- `excel_invoice_generator.py` is substantially rewritten: the sheet is now **5 columns (A–E)**, not 4 — column C is a spacer between the patient-address block and the payment-notice box on rows 10–14, and item-table rows merge B:C for Description (giving it a wider span) while A/D/E stay single-column (Service Date / Amount Paid / Copay-Deductible).
- Column widths: A=15, C=32.140625, D=20.140625 (exact fixture values); B and E are deliberately left unset, matching the fixture's use of Excel's implicit default width rather than an explicit one.
- Replaced the previous 3-tier font/row-height auto-shrink system with the fixture's fixed sizing: Arial 12 bold (clinic header), Arial 10 bold (contact info/statement labels/subtotal/total/portion-due), Arial 13 bold (title), Arial 9 (line items), Arial 11 bold (signature). Row heights are the fixture's exact values (18/18/18/6/15.95×4/12/15×5/12/20.1/6/15×2/12/18/15/…). Item rows below the fixed header extend downward 1:1 with item count, shifting SUBTOTAL/TOTAL/bottom-boxes/signature down by the same amount — everything above the item table (header through the item-table header row) stays fixed.
- Dropped the multi-tier print-scaling logic entirely; a single `fitToPage=True` (no explicit `paperSize`/`fitToWidth`/`fitToHeight`, matching the fixture's implicit-default style) now handles fitting arbitrary content onto one printed Letter page via Excel's own print-time scaling, rather than programmatically shrinking fonts.
- The payment-notice box (D10:E14) now uses the approved **fixed 5-line literal text** (with a deliberate mid-sentence line break before "(IRA Billing and Mgmt)" and specific leading whitespace) instead of the previous 4-line, auto-wrap-estimated version — sized to exactly fit 5 rows by design, per the approved copy. The patient-address box still uses dynamic wrap-line estimation (via `_count_wrapped_lines()`) since address content varies, floored at 5 rows to match the fixture.
- Per the fixture: the payment-notice box has **no border** (despite an earlier ask for one — the committed fixture is the ground truth for Phase 0 and doesn't have it). The "YOUR PORTION DUE" box (column C only) has **top-only / bottom-only** borders with no left/right — different from the "AMOUNT ENCLOSED" box (D:E merged), which gets a full `apply_box_border()` outline. Both are intentional, fixture-verified shapes, not oversights.
- Normalized two apparent authoring artifacts in the fixture rather than reproducing them: item rows 22–28 have explicit 15.0pt height instead of the fixture's unset/default (row 29, the last item, does have an explicit height — the inconsistency looks unintentional), and the SUBTOTAL row's B:C cell no longer gets a stray "double" bottom border that the rest of the row (A/D/E) doesn't have (would otherwise print as a broken half-underline under just the word "SUBTOTAL"). Both are documented in `tests/test_excel_invoice_generator.py`.
- Bugfix found while building this: `Border(top=THIN)`-style partial construction in openpyxl leaves the other sides as raw Python `None` rather than an empty `Side()`, which reads back differently than "never styled." Fixed by explicitly passing `Side()` for every omitted side.

#### 3. Column widths and "YOUR PORTION DUE" border corrected after visual review
- After seeing the Phase 0 output actually rendered, two of the fixture-derived decisions above turned out wrong in practice: column widths were revised to explicit on-screen pixel targets (A=105px, B=64px, C=225px, D=140px, E=140px), giving the payment-notice box's D:E span much more room (E was previously left at the ~64px implicit default, which was too narrow for the notice text to wrap the way it was authored). All five columns are now set explicitly (B and E are no longer left at Excel's implicit default).
- The "YOUR PORTION DUE" box (column C) now gets a complete `apply_box_border()` outline matching "AMOUNT ENCLOSED," reversing the earlier "match the fixture's top/bottom-only border exactly" call — the fixture's missing left border was confirmed by direct visual review to look broken, not intentional.
- `tests/fixtures/Example_2026_Invoice_07162026.xlsx` was regenerated from the exact synthetic inputs in `tests/conftest.py::golden_invoice_inputs`, using the corrected generator, so it now matches the generator byte-for-byte with no documented exceptions — the two "authoring artifact" workarounds from item #1 above (unset row heights on 22–28, stray double border on B30) are no longer needed since the fixture is self-generated rather than hand-built in Excel.

#### 4. Column-width pixel conversion formula corrected
- Item #3's initial pixel→width conversion used `width = (pixels − 5) / 7`, which measured 5px short of every target on screen (A: 100px not 105, B: 59px not 64, C: 220px not 225, D/E: 135px not 140 — a uniform -5px offset across all five columns). The correct formula, confirmed against all four measurements, is simply `width = pixels / 7` (no padding offset): A=15.0, B=9.14, C=32.14, D=20.0, E=20.0. Fixture regenerated again from the corrected values.

### Added

#### 2. Golden-file test suite for the Excel generator
- Added `pytest` to `requirements.txt` and a `pyproject.toml` with `[tool.pytest.ini_options]` (no test framework existed in the repo before this).
- Added `tests/fixtures/Example_2026_Invoice_07162026.xlsx` (synthetic data only) as the approved reference layout, and `tests/test_excel_invoice_generator.py`, which generates an invoice from synthetic inputs matching the fixture's data and asserts dimensions, merges, column widths, row heights, cell values, key-cell styling (font/format/alignment/borders), print setup, and correct row-shifting behavior when there are more line items than the fixture's 8 — all against the fixture, cell-for-cell.

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
