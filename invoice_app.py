import streamlit as st
import pandas as pd
import zipfile
import tempfile
import shutil
from pathlib import Path
from datetime import datetime
import io
import os

# Import your existing class
from complete_patient_invoice_generator import PatientInvoiceGenerator
from invoice_models import REQUIRED_TEMPLATE_PLACEHOLDERS, validate_cover_letter_template, VALIDATION_CATEGORIES
from clinic_config import get_clinic_config_source, ClinicConfigError

# --- Config: cover letter template path + required placeholders ---
TEMPLATE_CONFIG = {
    "default_template_path": Path(__file__).parent / "templates" / "Access_Multi_Letter_Cover.docx",
    "required_placeholders": REQUIRED_TEMPLATE_PLACEHOLDERS,
}

st.set_page_config(
    page_title="Medical Invoice Generator",
    page_icon="🏥",
    layout="wide"
)

st.title("🏥 Medical Invoice Generator")
st.markdown("Generate patient invoices, cover letters, and reports automatically")

# Create tabs for different sections
tab1, tab2, tab3 = st.tabs(["📁 Upload Files", "⚙️ Settings", "📊 Generate Reports"])

with tab1:
    st.header("Upload Required Files")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.subheader("Patient Roster")
        roster_file = st.file_uploader(
            "Upload Patient Roster CSV",
            type=['csv'],
            help="CSV file containing patient information"
        )
        
    with col2:
        st.subheader("Invoice Data")
        invoice_file = st.file_uploader(
            "Upload Invoice Excel File",
            type=['xlsx', 'xls'],
            help="Excel file with patient billing data"
        )
        
    with col3:
        st.subheader("Cover Letter Template")

        default_template_path = TEMPLATE_CONFIG["default_template_path"]
        default_template_exists = default_template_path.exists()

        with st.expander("Replace cover letter template (optional)"):
            template_upload = st.file_uploader(
                "Upload a .docx to use instead of the bundled default",
                type=['docx'],
                help="Overrides the bundled template for this session only, unless saved as the new default below."
            )

            if template_upload is not None:
                template_upload.seek(0)
                try:
                    missing = validate_cover_letter_template(template_upload)
                    if missing:
                        st.warning(f"Uploaded template is missing placeholders: {', '.join(missing)}")
                    else:
                        st.success("Uploaded template validated — all required placeholders found.")
                except Exception as e:
                    st.error(f"Could not read uploaded template: {e}")
                finally:
                    template_upload.seek(0)

                st.caption(
                    "⚠️ Saving overwrites the bundled default template file. On Streamlit Cloud this "
                    "change will NOT survive a redeploy — also commit the updated file to the repo "
                    "to make it permanent."
                )
                confirm_save = st.checkbox("I understand this overwrites the bundled default template")
                if st.button("💾 Save as new default template", disabled=not confirm_save):
                    default_template_path.parent.mkdir(parents=True, exist_ok=True)
                    template_upload.seek(0)
                    default_template_path.write_bytes(template_upload.read())
                    template_upload.seek(0)
                    st.success(f"Saved as new default: {default_template_path.name}")
                    st.rerun()

        # Resolve which template is active for this run: uploaded override wins
        if template_upload is not None:
            active_template_label = f"Uploaded override: {template_upload.name}"
            active_template_source = template_upload
        elif default_template_exists:
            active_template_label = f"Bundled default: {default_template_path.name}"
            active_template_source = default_template_path
        else:
            active_template_label = None
            active_template_source = None

        if active_template_source is None:
            st.warning("No cover letter template available. Upload one above to continue.")
        else:
            st.info(f"Active template: {active_template_label}")
            try:
                if hasattr(active_template_source, 'seek'):
                    active_template_source.seek(0)
                missing = validate_cover_letter_template(active_template_source)
                if hasattr(active_template_source, 'seek'):
                    active_template_source.seek(0)
                if missing:
                    st.warning(f"Active template is missing placeholders: {', '.join(missing)}")
            except Exception as e:
                st.error(f"Could not validate active template: {e}")

with tab2:
    st.header("Invoice Settings")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Statement Date")
        statement_date = st.date_input(
            "Statement Date",
            value=datetime.now().date(),
            help="Date to appear on the invoice"
        )
        
        amount_strategy = st.selectbox(
            "Amount Due Calculation",
            options=["auto", "copay_minus_paid", "total_minus_paid"],
            index=0,
            help="How to calculate the amount due"
        )
        
    with col2:
        st.subheader("Options")
        generate_csv = st.checkbox("Generate CSV exports", value=True)
        
        # Custom column mapping (optional)
        st.subheader("Custom Column Mapping (Optional)")
        with st.expander("Advanced: Map Excel columns"):
            st.info("Only fill these if your Excel file has different column names")
            name_col = st.text_input("Patient Name Column", placeholder="e.g., Patient Name")
            visit_date_col = st.text_input("Visit Date Column", placeholder="e.g., Service Date")
            total_amount_col = st.text_input("Total Amount Column", placeholder="e.g., Billed Amount")
            copay_col = st.text_input("Copay Column", placeholder="e.g., Co-pay")
            paid_col = st.text_input("Paid Column", placeholder="e.g., Patient Paid")

with tab3:
    st.header("Generate Reports")

    clinic_config_error = None
    try:
        clinic_config_source = get_clinic_config_source()
        st.caption(f"Config: {clinic_config_source}")
    except ClinicConfigError as e:
        clinic_config_error = str(e)

    if clinic_config_error:
        st.error(f"⚠️ Clinic configuration problem: {clinic_config_error}")

    # Check if all required files are uploaded / available
    files_ready = all([roster_file, invoice_file]) and active_template_source is not None and not clinic_config_error

    if not files_ready:
        if not clinic_config_error:
            st.warning("Please upload all required files in the 'Upload Files' tab before generating reports.")
        st.stop()
    
    st.success("All required files uploaded successfully!")

    st.subheader("Invoice Export Format")
    export_format_label = st.radio(
        "Choose which invoice file format(s) to generate",
        options=["PDF only", "Excel only", "Both PDF & Excel"],
        index=0,
        horizontal=True,
        help="Excel invoices are print-ready (US Letter, one page) and mirror the PDF layout."
    )
    export_format = {
        "PDF only": "pdf",
        "Excel only": "excel",
        "Both PDF & Excel": "both",
    }[export_format_label]

    st.subheader("Pre-Flight Validation")
    st.caption(
        "Scans the roster and invoice data for issues — unmatched or low-confidence "
        "patient matches, missing/malformed addresses, missing service dates, charges "
        "with no description, credit balances, and possible duplicate invoices (an "
        "overlapping service-date range already invoiced) — before anything is generated. "
        "Duplicate history is a local file and won't survive a Streamlit Cloud redeploy."
    )

    # Custom mapping is defined in tab2 but used here and by validation below.
    custom_mapping = {}
    if name_col: custom_mapping['name'] = name_col
    if visit_date_col: custom_mapping['visit_date'] = visit_date_col
    if total_amount_col: custom_mapping['total_amount'] = total_amount_col
    if copay_col: custom_mapping['copay'] = copay_col
    if paid_col: custom_mapping['paid'] = paid_col

    current_files_fingerprint = (roster_file.file_id, invoice_file.file_id)
    if st.session_state.get("validation_files_fingerprint") != current_files_fingerprint:
        # Roster/invoice changed since the last validation run (or this is
        # the first run) — any prior review no longer applies.
        st.session_state.pop("validation_report", None)
        st.session_state["validation_reviewed"] = False

    if st.button("🔍 Run Validation", use_container_width=True):
        with st.spinner("Validating..."):
            with tempfile.TemporaryDirectory() as val_temp_dir:
                val_temp_path = Path(val_temp_dir)
                val_roster_path = val_temp_path / "roster.csv"
                val_invoice_path = val_temp_path / "invoice.xlsx"
                with open(val_roster_path, "wb") as f:
                    f.write(roster_file.getbuffer())
                with open(val_invoice_path, "wb") as f:
                    f.write(invoice_file.getbuffer())

                try:
                    validator = PatientInvoiceGenerator(
                        amount_due_strategy=amount_strategy,
                        statement_date=statement_date.strftime("%Y-%m-%d")
                    )
                    st.session_state["validation_report"] = validator.validate_before_generation(
                        roster_file=str(val_roster_path),
                        invoice_file=str(val_invoice_path),
                        custom_mapping=custom_mapping if custom_mapping else None,
                    )
                    st.session_state["validation_files_fingerprint"] = current_files_fingerprint
                    st.session_state["validation_reviewed"] = False
                except Exception as e:
                    st.error(f"Validation failed: {str(e)}")
                    st.session_state.pop("validation_report", None)

    validation_report = st.session_state.get("validation_report")
    if validation_report is not None:
        if validation_report.issues:
            st.warning(
                f"⚠️ {validation_report.error_count} error(s), {validation_report.warning_count} "
                f"warning(s) found across {validation_report.total_patient_groups} patient group(s)."
            )
            issues_by_category = {}
            for issue in validation_report.issues:
                issues_by_category.setdefault(issue.category, []).append(issue)
            for category, label in VALIDATION_CATEGORIES.items():
                category_issues = issues_by_category.get(category)
                if not category_issues:
                    continue
                with st.expander(f"{label} ({len(category_issues)})"):
                    if category == "duplicate_invoice":
                        st.caption(
                            "Checked = skip this patient in the run below. Uncheck to "
                            "regenerate anyway (e.g. a legitimate correction/reprint)."
                        )
                        for issue in category_issues:
                            st.checkbox(
                                f"Skip **{issue.patient_name}** — {issue.detail}",
                                value=True,
                                key=f"dup_skip__{issue.patient_name}",
                            )
                    else:
                        for issue in category_issues:
                            icon = "🛑" if issue.severity == "error" else "⚠️"
                            st.write(f"{icon} **{issue.patient_name}** — {issue.detail}")
        else:
            st.success(f"✅ No issues found across {validation_report.total_patient_groups} patient group(s).")

        validation_text = PatientInvoiceGenerator._generate_validation_report_text(validation_report)
        st.download_button(
            "📄 Download Validation Report (.txt)",
            data=validation_text,
            file_name=f"Validation_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
            mime="text/plain",
        )

        st.checkbox(
            "I've reviewed the validation results above and want to proceed with generation",
            key="validation_reviewed",
        )
    else:
        st.info("Run validation before generating reports.")

    # Patients whose "Skip" checkbox (rendered above, in the duplicate-invoice
    # expander) is checked — collected from whatever duplicate issues the
    # current validation_report actually found.
    skip_patient_names = set()
    if validation_report is not None:
        for issue in validation_report.issues:
            if issue.category == "duplicate_invoice" and st.session_state.get(f"dup_skip__{issue.patient_name}", True):
                skip_patient_names.add(issue.patient_name)

    generation_blocked = not st.session_state.get("validation_reviewed", False)
    if generation_blocked:
        st.caption("⬆️ Run validation and check the box above to enable generation.")

    if st.button("🚀 Generate All Reports", type="primary", use_container_width=True, disabled=generation_blocked):
        try:
            with st.spinner("Processing invoices... This may take a few minutes."):
                
                # Create temporary directory for processing
                with tempfile.TemporaryDirectory() as temp_dir:
                    temp_path = Path(temp_dir)
                    
                    # Save uploaded files to temp directory
                    roster_path = temp_path / "roster.csv"
                    invoice_path = temp_path / "invoice.xlsx"
                    template_path = temp_path / "template.docx"
                    
                    with open(roster_path, "wb") as f:
                        f.write(roster_file.getbuffer())
                    with open(invoice_path, "wb") as f:
                        f.write(invoice_file.getbuffer())

                    if hasattr(active_template_source, 'getbuffer'):
                        # Uploaded file object (session override)
                        active_template_source.seek(0)
                        with open(template_path, "wb") as f:
                            f.write(active_template_source.getbuffer())
                        active_template_source.seek(0)
                    else:
                        # Path to the bundled default template
                        shutil.copy(active_template_source, template_path)

                    # custom_mapping was already built above, alongside validation

                    # Initialize generator
                    generator = PatientInvoiceGenerator(
                        amount_due_strategy=amount_strategy,
                        statement_date=statement_date.strftime("%Y-%m-%d")
                    )
                    
                    # Generate invoices
                    output_dir = temp_path / "output"
                    summary = generator.generate_invoices(
                        roster_file=str(roster_path),
                        invoice_file=str(invoice_path),
                        template_file=str(template_path),
                        output_dir=str(output_dir),
                        custom_mapping=custom_mapping if custom_mapping else None,
                        generate_csv=generate_csv,
                        export_format=export_format,
                        skip_patient_names=skip_patient_names,
                    )
                    
                    # Display results
                    st.success("✅ Invoice generation completed!")
                    
                    # Show summary
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Processed", summary.total_processed)
                    with col2:
                        st.metric("Skipped", summary.total_skipped)
                    with col3:
                        st.metric("Errors", summary.total_errors)
                    with col4:
                        st.metric("Total Amount Due", f"${summary.total_amount_due:.2f}")
                    
                    # Show processed patients
                    if summary.processed_patients:
                        st.subheader("Successfully Processed Patients")
                        for patient in summary.processed_patients:
                            st.write(f"✅ {patient}")
                    
                    # Show skipped patients
                    if summary.skipped_patients:
                        st.subheader("Skipped Patients")
                        for patient, reason in summary.skipped_patients:
                            st.write(f"⏭️ {patient} - {reason}")
                    
                    # Show errors
                    if summary.errors:
                        st.subheader("Errors")
                        for patient, error in summary.errors:
                            st.write(f"❌ {patient} - {error}")
                    
                    # Create downloadable zip file
                    if output_dir.exists():
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                            for file_path in output_dir.rglob('*'):
                                if file_path.is_file():
                                    # Create relative path for zip
                                    relative_path = file_path.relative_to(output_dir)
                                    zip_file.write(file_path, relative_path)
                        
                        zip_buffer.seek(0)

                        format_suffix = {"pdf": "", "excel": "_excel", "both": "_pdf_excel"}[export_format]
                        st.download_button(
                            label=f"📥 Download All Generated Files ({export_format_label})",
                            data=zip_buffer.getvalue(),
                            file_name=f"invoices{format_suffix}_{statement_date.strftime('%Y%m%d')}.zip",
                            mime="application/zip",
                            use_container_width=True
                        )
        
        except Exception as e:
            # No st.exception(e) here: a full traceback can echo back
            # patient data from local variables in the call stack, which
            # shouldn't be rendered in the UI.
            st.error(f"Error generating invoices: {str(e)}")

# Sidebar with instructions
with st.sidebar:
    st.header("📋 Instructions")
    st.markdown("""
    ### Step 1: Upload Files
    - **Patient Roster**: CSV with patient information
    - **Invoice Data**: Excel file with billing data
    - **Cover Letter Template**: uses the bundled default automatically — upload a `.docx` in "Replace cover letter template (optional)" only if you want to override it for this session (or save it as the new default)
    
    ### Step 2: Configure Settings
    - Set the statement date
    - Choose calculation method
    - Map custom columns if needed
    
    ### Step 3: Generate Reports
    - Choose an invoice export format (PDF, Excel, or both)
    - Click "Generate All Reports"
    - Download the zip file with all invoices

    ### Output Files
    Each patient gets:
    - PDF and/or Excel Invoice
    - Word Cover Letter
    - CSV Line Items (optional)
    """)
    
    st.header("📞 Support")
    st.info("If you encounter issues, check that your Excel file has the expected column names: Name, Visit Date, Total amount, copay, Paid")
