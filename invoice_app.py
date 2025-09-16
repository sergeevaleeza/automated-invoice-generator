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
        template_file = st.file_uploader(
            "Upload Cover Letter Template",
            type=['docx'],
            help="Word document template for cover letters"
        )

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
    
    # Check if all required files are uploaded
    files_ready = all([roster_file, invoice_file, template_file])
    
    if not files_ready:
        st.warning("Please upload all required files in the 'Upload Files' tab before generating reports.")
        st.stop()
    
    st.success("All required files uploaded successfully!")
    
    if st.button("🚀 Generate All Reports", type="primary", use_container_width=True):
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
                    with open(template_path, "wb") as f:
                        f.write(template_file.getbuffer())
                    
                    # Create custom mapping if provided
                    custom_mapping = {}
                    if name_col: custom_mapping['name'] = name_col
                    if visit_date_col: custom_mapping['visit_date'] = visit_date_col
                    if total_amount_col: custom_mapping['total_amount'] = total_amount_col
                    if copay_col: custom_mapping['copay'] = copay_col
                    if paid_col: custom_mapping['paid'] = paid_col
                    
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
                        generate_csv=generate_csv
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
                        
                        # Provide download button
                        st.download_button(
                            label="📥 Download All Generated Files",
                            data=zip_buffer.getvalue(),
                            file_name=f"invoices_{statement_date.strftime('%Y%m%d')}.zip",
                            mime="application/zip",
                            use_container_width=True
                        )
        
        except Exception as e:
            st.error(f"Error generating invoices: {str(e)}")
            st.exception(e)

# Sidebar with instructions
with st.sidebar:
    st.header("📋 Instructions")
    st.markdown("""
    ### Step 1: Upload Files
    - **Patient Roster**: CSV with patient information
    - **Invoice Data**: Excel file with billing data
    - **Cover Letter Template**: Word document template
    
    ### Step 2: Configure Settings
    - Set the statement date
    - Choose calculation method
    - Map custom columns if needed
    
    ### Step 3: Generate Reports
    - Click "Generate All Reports"
    - Download the zip file with all invoices
    
    ### Output Files
    Each patient gets:
    - PDF Invoice
    - Word Cover Letter
    - CSV Line Items (optional)
    """)
    
    st.header("📞 Support")
    st.info("If you encounter issues, check that your Excel file has the expected column names: Name, Visit Date, Total amount, copay, Paid")
