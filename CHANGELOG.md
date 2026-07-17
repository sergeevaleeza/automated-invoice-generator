# Changelog

## [Unreleased] — Use the real Zelle QR image instead of a generated one (2026-07-16)

### Changed

#### 1. Static, pre-made QR image takes priority over the generated one
- A real payment QR like Zelle's isn't just a URL — it's bank/Zelle-issued and encodes data a generator can't reproduce from a plain string. Added `templates/zelle_qr.jpg` (the actual clinic-issued Zelle QR image, committed to the repo the same way `templates/Access_Multi_Letter_Cover.docx` already is) and a new optional `qr_image_path` clinic config field.
- `qr_code.py`: added `resolve_qr_image_bytes(clinic)`, the single place both generators call. Resolution order: (1) `qr_image_path` (resolved relative to the repo root) if the file exists, (2) generate from `qr_content` via the `qrcode` library (unchanged from before), (3) `None` if `show_qr` is off or nothing usable is configured. `qr_settings()` now returns a 3-tuple `(show_qr, qr_content, qr_image_path)`.
- PDF (`add_optimized_footer()`) and Excel (`_clinic_derived_config()`/embedding block) both now call `resolve_qr_image_bytes()` directly instead of generating from `qr_content` themselves — no priority logic duplicated between the two.
- Real `clinic_config.json`: `qr_image_path` set to `templates/zelle_qr.jpg` (the static image used in practice); `qr_content` updated to the clinic's actual Zelle enrollment link (decoded from the QR: `{"name":"IRA BILLING AND MANAGEMENT, INC","action":"payment","token":"access.msmc@gmail.com"}`) so a *working* Zelle QR still generates as a fallback if the static image file is ever missing, instead of falling back to a plain pricing-page link.
- `tests/test_qr_code.py` extended with a `TestResolveQrImageBytes` class (priority order, missing-file fallback, disabled/unconfigured → `None`, and a sanity check against the real committed `templates/zelle_qr.jpg`) and split the PDF/Excel embedding tests into explicit static-image and generated-fallback cases. Full suite (68/68) and a manual pipeline run against the real `clinic_config.json` confirmed the static image embeds correctly in both formats (Excel anchors at row 9, column C — matching the payment-notice box position).

## [Unreleased] — Superbill export for a selected patient (2026-07-16)

### Added

#### 1. Single-patient superbill PDF, separate from the batch invoice flow
- Added `superbill_generator.py`: `generate_superbill_pdf(patient, clinic, service_lines, icd10_codes, statement_date, output_path)` — a clean letterhead-styled document (not a pixel-perfect CMS-1500 form, per the spec) with practice header, "SUPERBILL" title, side-by-side provider/patient info (name, DOB, address, NPI, tax ID/EIN, practice address), diagnosis codes (ICD-10), a per-service-line table (date, CPT code, description, charge, payment) with a TOTAL row, and a signature line. Reuses `PatientData`/`SuperbillServiceLine` (from `invoice_models.py`) — no billing/matching logic duplicated; this module is presentation only.
- `invoice_app.py`: new "🧾 Superbill" tab (4th tab). Once the roster and invoice files are uploaded, a patient `st.selectbox` lets the user pick one record; matched service lines and default ICD-10 codes are shown in an editable `st.data_editor` / `st.text_input` (CPT/ICD data is often not in the workbook, so the spec called for editable fields before generating) before the "🧾 Generate Superbill" button produces a download.
- Parsed roster/invoice data is cached in `st.session_state` keyed on `(roster_file.file_id, invoice_file.file_id, amount_strategy, statement_date)`, so switching patients in the selectbox doesn't re-parse the workbook on every rerun.

#### 2. Three-tier CPT code resolution and ICD-10 default resolution
- `PatientInvoiceGenerator.resolve_superbill_service_lines(patient_df)`: for each service line, resolves a CPT code in priority order — (1) an explicit `cpt_code` workbook column if present, (2) a 5-digit code embedded in the `type_of_service` text (e.g. "Med Management (CPT Code 99213)"), via the new shared `invoice_models.extract_embedded_cpt_code()` (also used to de-duplicate the existing `_has_cpt_codes()` check), (3) `clinic_config.json`'s new optional `default_cpt_by_service_type` mapping (case-insensitive match on service type), else blank — always meant to be reviewed/edited in the UI before generating.
- `PatientInvoiceGenerator.resolve_default_icd10_codes(patient_df)`: unique, sorted `icd10_code` workbook values if present, else `clinic_config.json`'s new optional `default_icd10_codes` list, else empty.
- New optional `clinic_config.json` fields: `default_icd10_codes` and `default_cpt_by_service_type`. Not added to `REQUIRED_KEYS` — existing configs keep working unchanged. Documented with placeholder values in the committed `clinic_config.example.json` template (NPI/EIN and the mapping are clinic-specific and must not be hardcoded in source since the repo is public).
- Added `tests/test_superbill.py` (12 tests): all three CPT resolution tiers, ICD-10 workbook-priority and clinic-default fallback (including the true "nothing configured" case), charge/payment/date pass-through, and PDF generation (single page, key identifying fields actually present in the decompressed content stream — ReportLab's default output is FlateDecode+ASCII85 compressed, so a raw-byte search wouldn't find embedded text).

#### 3. Streamlit tab execution-order fix
- Found while wiring up the new tab: `st.stop()` halts the *entire* script, not just the calling `with tabX:` block's rendering. The existing "📊 Generate Reports" tab calls `st.stop()` when its own prerequisites (e.g. a cover-letter template) aren't ready — which would have silently prevented the Superbill tab from ever rendering, since Streamlit tab bodies all execute top-to-bottom in script order regardless of which tab is visually selected. Fixed by moving the Superbill tab's body earlier in the script (before the Reports tab), with a comment explaining why — `st.tabs()` controls visual order independently of execution order.
- While relocating code, found `custom_mapping` (the roster/invoice column-mapping dict) was actually being built inside the Reports tab despite a comment claiming it was "defined" earlier — which would have made it undefined for the now-earlier-running Superbill tab. Moved its construction to the Settings tab (immediately after the column-mapping UI it depends on) so both tabs can rely on it.

## [Unreleased] — QR code on invoices (2026-07-16)

### Added

#### 1. Optional QR code, PDF and Excel, config-gated
- Added `qrcode[pil]` to `requirements.txt` and `qr_code.py`: `generate_qr_png_bytes(content)` (presentation-agnostic PNG bytes, shared by both generators — no QR-building logic duplicated) and `qr_settings(clinic)` resolving `(show_qr, qr_content)` from a loaded clinic config, defaulting to disabled and falling back to the clinic's `website` if `qr_content` is omitted.
- Two new **optional** `clinic_config.json` fields: `show_qr` (bool) and `qr_content` (string, e.g. a payment/pricing page URL). Deliberately not added to `REQUIRED_KEYS` — an existing config from before this feature keeps working unchanged (QR just stays off).
- **PDF**: embedded in `add_optimized_footer()`, bottom-right corner, ~0.9in, sized to not overlap the centered footer text (verified the footer's longest line leaves ~1.75in clear on the right at this length). Uses `reportlab.lib.utils.ImageReader` wrapping the PNG bytes.
- **Excel**: embedded as a floating `openpyxl` image anchored at column C of the patient-info/payment-notice row block — a column that's *always* blank there by design (the deliberate spacer between the two boxes, regardless of how many rows they span for a longer address), so it can't overlap existing text or disturb the Phase 0 golden-fixture layout. Verified this holds: the golden-fixture test suite (which uses `clinic_config.example.json`, `show_qr: false`) is completely unaffected. Sized via `width=height=86` (openpyxl converts image pixel dimensions to the saved anchor's EMU extent at 96dpi, confirmed empirically by round-tripping through save/reload — 86px ≈ 0.9in; an earlier 65px guess, by analogy to the PDF's 72dpi point-based sizing, actually produced 0.677in).
- Added `tests/test_qr_code.py` (9 tests): PNG validity, `qr_settings()` defaults/fallback/override, PDF embedding present/absent by toggle (checking for a real `/Subtype /Image` XObject in the raw PDF bytes), Excel embedding present/absent by toggle and anchor position, and an explicit guard test asserting the shared test clinic config keeps `show_qr` off so no other Excel test silently starts embedding an unexpected image.
- Note on numbering: this is "QR code on invoices" from the original Phase 1 feature list (item 5 there) — it got mislabeled "item 6" mid-session after Phase 1 was resequenced by dependency order (hygiene sweep moved to the front), which actually corresponds to the Superbill export. Flagged and corrected; Superbill is next.

## [Unreleased] — Phase 1 item 4: Enhanced batch summary (2026-07-16)

### Added

#### 1. Structured per-patient records replace the plain-string summary list
- `invoice_models.py`: added `ProcessedPatientRecord` (display_name, service_date_start/end, amount_due, amount_paid). `ProcessingSummary.processed_patients: List[str]` is replaced with `processed_records: List[ProcessedPatientRecord]` — the old field held pre-formatted strings like `"Name - Due: $X, Paid: $Y"`, which couldn't carry the service-date range the spec asked for without fragile string parsing. Also added `total_amount_paid: float` (the existing `total_amount_due` is now documented as "total outstanding").
- `generate_invoices()` populates each record from data it already computes per patient — the service-date range reuses the same `_service_date_range()` helper added for duplicate detection, so nothing is duplicated.

#### 2. `Processing_Summary_*.txt` extended with financial totals, service dates, and validation warnings
- `_generate_summary_report()` is split into `_generate_summary_report_text()` (pure text, mirrors the `_generate_validation_report_text()` pattern from Phase 1 item 2) and a thin file-writing wrapper — so the exact same content can be shown in the UI, not just written to disk.
- New "Total Invoiced (billed this period)" (`total_amount_due + total_amount_paid`), "Total Outstanding (amount due)", and "Total Already Paid" lines. Per-patient lines now show the service-date range (e.g. "01/05/2026 to 02/10/2026") alongside due/paid amounts.
- `generate_invoices()` gained an optional `validation_report` parameter — when the caller already ran `validate_before_generation()` (as `invoice_app.py` always does now, gated behind the review checkbox), its warnings are appended to the summary under a new "VALIDATION WARNINGS" section, giving one combined report instead of two separate ones. Purely informational — doesn't change what gets generated.
- Skipped-patient reasons were already tracked per-patient (`skipped_patients: List[Tuple[str, str]]`) and needed no changes — currently the only reason that occurs is "Skipped by user (duplicate invoice)" from Phase 1 item 3; earlier-considered reasons like "credit balance" or "filtered insurance" don't apply since credit-balance patients are invoiced (not skipped, per the correction two entries back) and no insurance-filtering feature exists in this codebase.

#### 3. Same summary rendered in the UI after a run
- `invoice_app.py`: the post-generation results section now shows the financial breakdown as three additional metrics (Total Invoiced / Outstanding / Already Paid), lists processed patients with their service-date range, and adds a "📋 Full Summary Report" expander showing the identical text that gets written to `Processing_Summary_*.txt` (via `st.code`), plus a standalone download button for it.
- Added `tests/test_summary_report.py` (5 tests): structured record fields are populated correctly, the three financial totals reconcile, the text report contains the expected service-date and totals lines, validation warnings are included when a report is passed, and the file actually gets written to `output_dir`. Full suite (37/37) and the PDF/Excel/both regression pipeline (re-checked with the scratch test script updated for the renamed field) confirmed unaffected otherwise.

## [Unreleased] — Phase 1 item 3: Duplicate invoice protection (2026-07-16)

### Added

#### 1. Local SQLite run-history store
- Added `run_history.py`: a small SQLite-backed store (`data/run_history.db`, gitignored — operational data about real patients) recording, per successfully-generated invoice, the patient's identity, service-date range, statement (invoice) date, and output filenames. `patient_key(prn, first_name, last_name)` prefers PRN (stable across differently-formatted invoice workbooks) and falls back to a normalized name for unmatched patients. `find_overlapping_runs()` does a standard inclusive date-range overlap query; `record_invoice_run()` inserts a new record.
- Caveat noted in the module docstring and the UI: this is a local file — it will **not** persist across a Streamlit Cloud redeploy, only within one running container's lifetime.

#### 2. Duplicate-invoice detection wired into pre-flight validation
- `validate_before_generation()` now also flags a `duplicate_invoice` category: for each patient group, computes the service-date range (`_service_date_range()`, new helper — earliest/latest parseable `visit_date`) and checks `run_history` for an overlapping prior run, producing messages in the exact format originally requested: "Already invoiced for 01/12 to 04/30 on 2026-06-15." (ISO dates; the UI shows the patient name alongside).
- `invoice_app.py`: duplicate-invoice issues render as per-patient checkboxes ("Skip **Name** — already invoiced...") instead of plain text, checked (skip) by default, in their own section of the Pre-Flight Validation panel.

#### 3. `generate_invoices()` respects skip choices and records successful runs
- New `skip_patient_names` parameter: patients named here are skipped entirely (added to `skipped_patients` with reason "Skipped by user (duplicate invoice)"), never generated, never recorded to `run_history`. `invoice_app.py` collects this set from the checkbox states of whatever duplicate issues the current validation report found, and passes it through when "Generate All Reports" runs.
- Every patient that's actually generated (not skipped, no error) gets recorded to `run_history` afterward — service-date range, the statement date used, and the filenames actually written (only the formats that were actually generated, e.g. no `.xlsx` entry when `export_format="pdf"`).
- Added `tests/test_run_history.py` (12 tests: patient_key derivation, overlap boundary cases, cross-patient isolation, most-recent-first ordering) and `tests/test_duplicate_detection.py` (4 integration tests: no false positive on a first run, a second batch is correctly flagged after the first run recorded history, `skip_patient_names` excludes generation and doesn't record a phantom run, and a later *non-overlapping* service period for the same patient is correctly not flagged).
- Found and fixed a test-isolation bug while building this: two `generate_invoices()` calls in `tests/test_credit_balance.py` (from the previous entry) didn't pass `run_history_db_path`, so they wrote to the real `data/run_history.db` — causing a later, unrelated validation test to fail against polluted state. Fixed by threading an explicit `tmp_path`-based DB through every test that touches run history (in that file and `tests/test_validation.py`), and deleted the accidentally-created `data/run_history.db` (gitignored, contained only test artifacts). Full suite (32/32) confirmed clean of any default-path usage after the fix.

## [Unreleased] — Correction: don't skip overpaid patients after all (2026-07-16)

### Changed

#### 1. Overpaid (credit-balance) patients are invoiced, not skipped
- Reverses part of the previous entry: that fix made the dead-code `total_due < 0` skip check reachable again, on the assumption its original intent (skip credit-balance patients) was the desired behavior. Per clarification, it wasn't — a patient who overpaid should still receive an invoice showing their credit and $0.00 due, not be silently excluded from the batch. They already have a "Previous Balance (Overpaid)" line item showing the credit amount; the only thing that changes is whether an invoice is generated at all.
- `generate_invoices()` no longer skips on `total_due < 0`. It now floors the value for display (`total_due = max(0, raw_total_due)` — nobody's invoice should show a negative "amount due") while still generating the PDF/Excel/DOCX/CSV normally. Tracks a `credit_balance_count` on the summary (informational, mirroring the existing `zero_balance_count`) instead of adding to `skipped_patients`.
- `_generate_invoice_lines()` keeps returning the unfloored `total_due` (from the previous fix) — that part was correct and is still needed so callers can tell a true credit apart from an exact zero balance; only `generate_invoices()`'s *use* of a negative value (skip vs. floor-and-invoice) was wrong.
- `validate_before_generation()`'s `negative_balance` warning text updated: no longer says the patient "will be skipped," says they'll "still be invoiced, showing the credit and $0.00 due."
- `tests/test_credit_balance.py` rewritten to assert the corrected behavior: a pure-credit patient is invoiced (not skipped, `Due: $0.00`), its line items actually contain the credit amount, an offsetting case still nets correctly, and validation's message no longer implies a skip. Full suite (16/16) and the PDF/Excel/both regression pipeline confirmed unaffected otherwise.

## [Unreleased] — Fix credit-balance skip logic (2026-07-16)

### Fixed

#### 1. Credit-balance patients are now actually skipped, not $0.00-invoiced
- Follow-up to the previous entry's "found, not fixed" item: `_generate_invoice_lines()` no longer floors `total_due` at `max(0, ...)` before returning it. That floor made `generate_invoices()`'s `if total_due < 0: skip as credit balance` permanently unreachable — a patient with a negative previous balance and no offsetting charges this period fell through to the zero-balance branch and got a **$0.00 invoice generated** instead of being skipped. The floor is no longer needed: `generate_invoices()` already checks `total_due < 0` and `continue`s immediately after calling `_generate_invoice_lines()`, before any generation happens, so a negative value is never actually used to produce an invoice — it now reaches the (already-correct) skip check instead.
- `validate_before_generation()`'s `negative_balance` check previously used `previous_balance < 0` alone as a proxy — inaccurate, since same-period charges can bring the net balance back to zero or positive even when the previous balance was a credit. It now calls `_generate_invoice_lines()` directly and checks the real `total_due < 0` condition, so validation can't drift from what generation actually does. Updated the warning text accordingly ("will be skipped entirely" instead of "still generates a $0.00 invoice").
- Added `tests/test_credit_balance.py` (3 tests): a pure-credit patient is skipped, not zero-invoiced; a negative previous balance fully offset by this period's charges is correctly **not** skipped (the exact case the old proxy check got wrong); and validation agrees with generation in both directions. Full suite (15/15) and PDF/Excel/both regression pipeline confirmed unaffected otherwise.

## [Unreleased] — Phase 1 item 2: Pre-flight validation report (2026-07-16)

### Added

#### 1. `validate_before_generation()` — read-only scan before any files are generated
- Added `ValidationIssue`/`ValidationReport` dataclasses and a `VALIDATION_CATEGORIES` label map to `invoice_models.py`.
- Added `PatientInvoiceGenerator.validate_before_generation(roster_file, invoice_file, custom_mapping=None) -> ValidationReport` — reuses `load_patient_roster()`/`load_invoice_data()`/`_match_patient()` (no parsing/matching logic duplicated) to scan for: patients with no roster match or a low-confidence/ambiguous fuzzy match (shows the best match + confidence score), missing or malformed addresses (blank street/city/state, or a postal code that doesn't look like a valid US ZIP), service lines with no visit date, charges/payments with no service description, and negative (credit) previous balances. Generates zero files.
- `_match_patient()` now returns a third value, `confidence_score` (1.0 for an exact match, the fuzzy score otherwise, 0.0 for no match) — needed to detect low-confidence matches. Its one existing caller in `generate_invoices()` was updated; behavior there is unchanged.
- Added `PatientInvoiceGenerator._generate_validation_report_text()` (a `@staticmethod`, mirrors `_generate_summary_report()`'s style) to render a `ValidationReport` as plain text for export alongside `Processing_Summary_*.txt`.
- `invoice_app.py`: new "Pre-Flight Validation" section between the export-format selector and the Generate button. A "🔍 Run Validation" button parses the current roster/invoice into a scratch temp dir, runs the scan, and stores the report in `st.session_state` (first real use of session state in this app). Issues are grouped into expanders by category with error/warning icons, plus a "📄 Download Validation Report (.txt)" button. The "🚀 Generate All Reports" button is `disabled` until a checkbox ("I've reviewed the validation results...") is checked — and that checkbox, along with any prior report, is automatically reset whenever the uploaded roster/invoice files change (tracked via `UploadedFile.file_id`), so a stale review can't wave through different data.
- Added `tests/test_validation.py` (4 tests): synthetic data covering all 7 issue categories, error/warning severity counts, a clean-data case producing zero issues, and the text export.

#### 2. Found (not fixed): credit-balance skip logic in `generate_invoices()` is dead code
- `_generate_invoice_lines()` always floors `total_due` at `max(0, ...)`, so `generate_invoices()`'s `if total_due < 0: skip as credit balance` can never trigger — a patient with a negative previous balance and no charges this period currently falls through to the zero-balance branch and gets a **$0.00 invoice generated** instead of being skipped. Confirmed empirically, not a misread. Per discussion, left generation behavior unchanged for this item — `validate_before_generation()` computes the pre-floor balance itself and surfaces it as a `negative_balance` warning (noting the current behavior) so it's visible before generating rather than fixed as a side effect. Revisit as its own deliberate change if the skip behavior should actually work.

## [Unreleased] — Phase 1 item 1 follow-up: Streamlit secrets fallback for clinic config (2026-07-16)

### Fixed

#### 1. clinic_config.json is gitignored, so it never exists on a fresh Streamlit Cloud deploy
- `clinic_config.py` now resolves clinic identity from either of two sources, in order: (a) `clinic_config.json` if present (local dev), or (b) a `[clinic_config]` table in Streamlit secrets (`st.secrets`) if the file isn't there (cloud deploys). Only raises `ClinicConfigError` if neither source has a complete config — the error message now mentions both options and points to the new `docs/DEPLOY.md`.
- Centralized in `load_clinic_config()` (unchanged signature/return shape — still a plain dict regardless of source) plus a new `get_clinic_config_source()` for display purposes, both built on a shared internal `_resolve_clinic_config()` so the two can't drift out of sync. Every existing caller (`excel_invoice_generator.py`, `complete_patient_invoice_generator.py`, `invoice_app.py`) already went through `load_clinic_config()`, so no call-site changes were needed beyond `invoice_app.py`'s new status line.
- `st.secrets` returns `AttrDict`-like objects, not plain dicts — added `_attrdict_to_dict()` to recursively convert (duck-typed on `.items()` rather than importing Streamlit's internal `AttrDict` class, so it's not tied to a specific Streamlit version). Accessing `st.secrets` when no secrets are configured at all raises `StreamlitSecretNotFoundError`; caught broadly and treated as "not available," not an error — only the final "neither source worked" case raises.
- `invoice_app.py`'s "Generate Reports" tab now shows `Config: local file` or `Config: Streamlit secrets` via `st.caption()` so it's clear at a glance which source is active.
- Error/validation messages only ever name missing *keys* (e.g. "missing required field(s): npi"), never values, from either source.
- Added `docs/DEPLOY.md` with the exact `[clinic_config]` TOML block to paste into Streamlit Cloud's Secrets panel (same fields as `clinic_config.example.json`, placeholder values only), plus `.streamlit/secrets.toml` to `.gitignore` for local secrets-path testing.
- Verified locally: (1) real `clinic_config.json` present → `local file`; (2) file renamed away with a mock `.streamlit/secrets.toml` in place → `Streamlit secrets`, full `generate_invoices()` pipeline confirmed working through this path; (3) both missing → the clear, updated `ClinicConfigError`, surfaced correctly through `invoice_app.py` with no exceptions (checked via `AppTest`). Full pytest suite (8/8) and PDF/Excel/both regression checks unaffected throughout.

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
