# 🏥 **[Medical Invoice Generator](https://automated-invoice-generator.streamlit.app/)**

Generate per‑patient **PDF invoices**, **DOCX cover letters**, optional **CSV line items**, and a **summary report** from a patient roster and a billing spreadsheet. Includes fuzzy patient matching, previous‑balance handling, and flexible column mapping — all via a Streamlit UI.

Built around a reusable `PatientInvoiceGenerator` class with a simple Streamlit front end.

## ✨ Features
- **Streamlit UI** with tabs for Uploads → Settings → Generate, plus a one‑click **Download ZIP** of all outputs.
- **PDF invoices** (ReportLab), **DOCX cover letters** (python‑docx), and **CSV** export per patient.
- **Fuzzy name matching** between invoice names and roster (handles typos/variants).
- **Previous balance** line item + **amount‑due strategies** (`auto`, `copay_minus_paid`, `total_minus_paid`).
- **Column aliasing & custom mapping** so you don’t have to rename your spreadsheet headers.
- **Due date** auto‑calculated from statement date (weekend‑aware).
- **Comprehensive summary report** and clean patient‑specific folder naming.

## 🧱 Project structure
```
.
├─ invoice_app.py                      # Streamlit UI (upload, settings, run, download zip)
├─ complete_patient_invoice_generator.py  # Core class: PDF/DOCX/CSV generation, matching, summary
└─ output/                              # Created at runtime with per‑patient folders/files
```

## 📥 Inputs

### 1) Patient roster (CSV)
Common headers detected include: `Patient Record Number/PRN`, `First name`, `Last name`, `DOB`, `Address Line 1/2`, `City`, `State`, `Postal Code`. There is a fallback parser for odd, space‑separated exports.

### 2) Invoice data (Excel, Sheet1)
**Required** (or aliased) columns:
- `Name` (e.g., `LastName, FirstName`, variants supported)
- `Visit Date` (aka `Service Date`, `Date of Service`, `DOS`)
- `Total amount`
- `copay`
- `Paid`

**Optional**:
- `Previous Balance`
- `Insurance`

Aliases and **custom mapping** are supported in the UI’s “Advanced: Map Excel columns” expander.

### 3) Cover letter template (DOCX)
Use placeholders such as `[First Name]`, `[Last Name]`, `[Full Name]`, `[Address Line 1]`, `[City]`, `[State]`, `[Postal Code]`, `[Patient Record Number]`. The generator replaces them across paragraphs/tables while preserving formatting.

## 📦 Installation

```bash
# Python 3.10+ (3.12 OK). Use a virtual environment.
python3 -m venv .venv
source .venv/bin/activate

pip install --upgrade pip
pip install -r requirements.txt
```

> On Ubuntu/WSL, if your browser doesn’t auto‑open, visit `http://localhost:8501` manually.

## ▶️ Running the app

```bash
streamlit run invoice_app.py
```

**UI steps**
1. **Upload Files**: roster CSV, invoice Excel, DOCX cover letter template.  
2. **Settings**: choose statement date, amount‑due strategy, optional column mapping, CSV export toggle.  
3. **Generate Reports** → wait for processing → **Download ZIP** of all outputs.

## 🗂️ Outputs
For each processed patient a folder like:
```
output/
  LastName_FirstName_PRN/
    LastName_YYYY_Invoice_MMDDYYYY.pdf
    LastName_YYYY_Envelope_MMDDYYYY.docx
    LastName_YYYY_InvoiceItems_MMDDYYYY.csv   # if CSV export enabled
Processing_Summary_YYYYMMDD.txt               # top‑level
```
Totals reflect **copay vs paid** and **previous balance** where applicable; “No open balance” patients are skipped with a reason.

## 🧠 How amount due is computed
- **auto**: use `copay - paid` if copay exists, else `total - paid`  
- **copay_minus_paid**: `copay - paid`  
- **total_minus_paid**: `total - paid`  
Values are floored at 0.

## 🔩 Programmatic use

```python
from complete_patient_invoice_generator import PatientInvoiceGenerator

gen = PatientInvoiceGenerator(amount_due_strategy="auto", statement_date="2025-09-12")
summary = gen.generate_invoices(
    roster_file="PatientListReport_active.csv",
    invoice_file="inv1.xlsx",
    template_file="CoverTemplate.docx",
    output_dir="output",
    custom_mapping=None,      # or {'name': 'Patient Name', 'copay': 'Co-pay', ...}
    generate_csv=True
)
print(summary.total_processed, summary.total_amount_due)
```

## 🧪 Tips & Troubleshooting
- **`ModuleNotFoundError: docx`** → Ensure you installed **`python-docx`** (not `docx`).  
- **PEP 668 / “externally‑managed”** → Always install with `pip` **inside your venv**.  
- **Browser didn’t open on WSL** → The app is running; open `http://localhost:8501` manually.  
- **Columns not found** → Use the **Advanced column mapping** expander in Settings.

## 🔐 PHI notice
This tool may process protected health information (PHI). Use secure storage and access controls appropriate to your environment.

## 📄 License
MIT (see `LICENSE`). If you need another license, replace accordingly.
