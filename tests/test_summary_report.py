"""Tests for the enhanced batch summary: structured per-patient records,
total invoiced/outstanding/paid breakdown, and inclusion of a pre-flight
ValidationReport's warnings. All patient/roster data here is synthetic."""
import pandas as pd
import pytest
from docx import Document

from complete_patient_invoice_generator import PatientInvoiceGenerator
from tests.conftest import TEST_CLINIC_CONFIG


@pytest.fixture
def generator():
    return PatientInvoiceGenerator(
        amount_due_strategy="auto", statement_date="2026-07-16",
        clinic_config=TEST_CLINIC_CONFIG,
    )


@pytest.fixture
def minimal_template(tmp_path):
    path = tmp_path / "template.docx"
    Document().save(path)
    return str(path)


@pytest.fixture
def roster_and_invoice(tmp_path):
    roster_rows = [dict(**{
        "Patient Identifier": "1", "Patient First Name": "Alice", "Patient Last Name": "Anderson",
        "DOB": "1980-01-01", "Address Line 1": "1 Main St", "Address Line 2": "",
        "City": "Testville", "State": "CA", "Postal Code": "94000",
    })]
    roster_path = tmp_path / "roster.csv"
    pd.DataFrame(roster_rows).to_csv(roster_path, index=False)

    invoice_rows = [
        dict(Name="Anderson, Alice", **{
            "Visit Date": "2026-01-05", "Total amount": 200, "Copay": 40, "Paid": 10,
            "Previous Balance": "", "Type Of Service": "Psychotherapy",
        }),
        dict(Name="Anderson, Alice", **{
            "Visit Date": "2026-02-10", "Total amount": 200, "Copay": 40, "Paid": 10,
            "Previous Balance": "", "Type Of Service": "Psychotherapy",
        }),
    ]
    invoice_path = tmp_path / "invoice.xlsx"
    pd.DataFrame(invoice_rows).to_excel(invoice_path, sheet_name="Sheet1", index=False)
    return str(roster_path), str(invoice_path)


def test_processed_records_have_service_date_range_and_amounts(generator, minimal_template, roster_and_invoice, tmp_path):
    roster_path, invoice_path = roster_and_invoice
    summary = generator.generate_invoices(
        roster_file=roster_path, invoice_file=invoice_path, template_file=minimal_template,
        output_dir=str(tmp_path / "output"), generate_csv=False, export_format="pdf",
        run_history_db_path=tmp_path / "run_history.db",
    )
    assert len(summary.processed_records) == 1
    record = summary.processed_records[0]
    assert record.display_name.startswith("Alice Anderson")
    assert record.service_date_start == "2026-01-05"
    assert record.service_date_end == "2026-02-10"
    assert record.amount_due == 60.0  # (40-10) + (40-10) copay-minus-paid per row
    assert record.amount_paid == 20.0  # 10 + 10 paid


def test_summary_totals_invoiced_outstanding_paid(generator, minimal_template, roster_and_invoice, tmp_path):
    roster_path, invoice_path = roster_and_invoice
    summary = generator.generate_invoices(
        roster_file=roster_path, invoice_file=invoice_path, template_file=minimal_template,
        output_dir=str(tmp_path / "output"), generate_csv=False, export_format="pdf",
        run_history_db_path=tmp_path / "run_history.db",
    )
    assert summary.total_amount_due == 60.0       # outstanding
    assert summary.total_amount_paid == 20.0       # already paid
    total_invoiced = summary.total_amount_due + summary.total_amount_paid
    assert total_invoiced == 80.0                  # billed this period (== copay sum here)


def test_summary_text_includes_service_dates_and_totals(generator, minimal_template, roster_and_invoice, tmp_path):
    roster_path, invoice_path = roster_and_invoice
    summary = generator.generate_invoices(
        roster_file=roster_path, invoice_file=invoice_path, template_file=minimal_template,
        output_dir=str(tmp_path / "output"), generate_csv=False, export_format="pdf",
        run_history_db_path=tmp_path / "run_history.db",
    )
    text = generator._generate_summary_report_text(summary)
    assert "Total Invoiced (billed this period): $80.00" in text
    assert "Total Outstanding (amount due): $60.00" in text
    assert "Total Already Paid: $20.00" in text
    assert "01/05/2026 to 02/10/2026" in text
    assert "Alice Anderson" in text


def test_summary_text_includes_validation_warnings(generator, minimal_template, roster_and_invoice, tmp_path):
    roster_path, invoice_path = roster_and_invoice
    db_path = tmp_path / "run_history.db"
    report = generator.validate_before_generation(
        roster_file=roster_path, invoice_file=invoice_path, run_history_db_path=db_path,
    )
    summary = generator.generate_invoices(
        roster_file=roster_path, invoice_file=invoice_path, template_file=minimal_template,
        output_dir=str(tmp_path / "output"), generate_csv=False, export_format="pdf",
        run_history_db_path=db_path, validation_report=report,
    )
    text = generator._generate_summary_report_text(summary, validation_report=report)
    if report.issues:
        assert "VALIDATION WARNINGS" in text
    else:
        # Clean data for this fixture - assert the section is correctly omitted, not broken
        assert "VALIDATION WARNINGS" not in text


def test_summary_file_written_to_output_dir(generator, minimal_template, roster_and_invoice, tmp_path):
    roster_path, invoice_path = roster_and_invoice
    output_dir = tmp_path / "output"
    generator.generate_invoices(
        roster_file=roster_path, invoice_file=invoice_path, template_file=minimal_template,
        output_dir=str(output_dir), generate_csv=False, export_format="pdf",
        run_history_db_path=tmp_path / "run_history.db",
    )
    summary_files = list(output_dir.glob("Processing_Summary_*.txt"))
    assert len(summary_files) == 1
    content = summary_files[0].read_text()
    assert "Total Invoiced" in content
