# Deploying to Streamlit Cloud

## Clinic configuration

The app needs your practice's identity (name, doctor, addresses, EIN/NPI,
payment info) to generate invoices. It's never hardcoded in this repo —
`clinic_config.py` loads it from one of two places, in order:

1. **`clinic_config.json`** in the project root, if present. This file is
   gitignored — copy `clinic_config.example.json` to `clinic_config.json`
   and fill in your real details. Works for local development; it will
   **not** exist on a fresh Streamlit Cloud deploy, since it's never
   committed.
2. **Streamlit secrets**, if `clinic_config.json` isn't present. On
   Streamlit Cloud: open your app → **Settings → Secrets**, and paste in a
   `[clinic_config]` table with the same fields, filled in with your real
   details:

   ```toml
   [clinic_config]
   clinic_name = "YOUR CLINIC NAME, INC."
   doctor_name = "JANE A. DOE, MD"
   specialty = "BOARD CERTIFIED PHYSICIAN"
   ein = "00-0000000"
   npi = "0000000000"
   office_address = "123 MAIN STREET, SUITE 100, YOUR CITY, ST 00000"
   mailing_address = "PO BOX 000, YOUR CITY, ST 00000"
   email = "BILLING@YOURCLINIC.COM"
   website = "https://yourclinic.example.com/"
   pricing_page_url = "https://yourclinic.example.com/pricing.html"
   phone = "(000)000-0000"
   zelle_email = "billing@yourclinic.com"
   check_payable_to = "Jane A. Doe, MD"
   provider_name_for_signature = "Jane A. Doe, MD"

   # --- Optional fields below — omit entirely if you don't use these features ---

   # QR code on invoices (PDF + Excel). qr_image_path takes priority when
   # set and the file exists in the repo (e.g. a real bank/Zelle-issued QR
   # image committed under templates/) — falls back to generating a QR
   # from qr_content (a URL/text) otherwise.
   show_qr = false
   qr_image_path = "templates/zelle_qr.jpg"
   qr_content = "https://yourclinic.example.com/pricing.html"

   # Superbill export defaults (still editable per-patient in the UI).
   default_icd10_codes = ["Z00.00"]

   [clinic_config.default_cpt_by_service_type]
   psychotherapy = "90837"
   "med management" = "99213"
   ```

   The values above are placeholders — replace them with your practice's
   real information. Field names must match exactly (same keys as
   `clinic_config.example.json`). **Note:** every field your local
   `clinic_config.json` has, including these optional ones, needs to be
   copied into Secrets separately — they're two independent config sources,
   so adding a field locally (e.g. turning on `show_qr`) does nothing on
   Streamlit Cloud until you paste the same change into Secrets too.

   For **local testing** of the secrets path, save the same block to
   `.streamlit/secrets.toml` in the project root (already gitignored —
   never commit this file).

If neither source is available, or a source is missing required fields,
the app shows a clear error on the "Generate Reports" tab naming which
fields are missing (never their values) and how to fix it — it will not
silently generate invoices with placeholder or incomplete clinic identity.

Once configured, the "Generate Reports" tab shows a small status line —
`Config: local file` or `Config: Streamlit secrets` — so it's always clear
at a glance which source is active.
