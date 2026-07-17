"""Integration tests for duplicate-invoice detection: validate_before_generation()
must flag an overlapping service-date range already recorded in run_history,
and generate_invoices() must skip a patient named in skip_patient_names while
recording every successfully-generated invoice for future overlap checks.
All patient/roster data here is synthetic."""
import pandas as pd
import pytest
from docx import Document

import run_history
from complete_patient_invoice_generator import PatientInvoiceGenerator
from tests.conftest import TEST_CLINIC_CONFIG


@pytest.fixture
def generator():
    return PatientInvoiceGenerator(
        amount_due_strategy="auto", statement_date="2026-06-15",
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
        "Patient Identifier": "1", "Patient First Name": "Alex", "Patient Last Name": "Guzhavin",
        "DOB": "1980-01-01", "Address Line 1": "1 Main St", "Address Line 2": "",
        "City": "Testville", "State": "CA", "Postal Code": "94000",
    })]
    roster_path = tmp_path / "roster.csv"
    pd.DataFrame(roster_rows).to_csv(roster_path, index=False)

    invoice_rows = [dict(Name="Guzhavin, Alex", **{
        "Visit Date": "2026-01-12", "Total amount": 200, "Copay": 40, "Paid": 0,
        "Previous Balance": "", "Type Of Service": "Psychotherapy",
    })]
    invoice_path = tmp_path / "invoice.xlsx"
    pd.DataFrame(invoice_rows).to_excel(invoice_path, sheet_name="Sheet1", index=False)
    return str(roster_path), str(invoice_path)


def test_first_run_generates_no_duplicate_warning(generator, roster_and_invoice, tmp_path):
    roster_path, invoice_path = roster_and_invoice
    db_path = tmp_path / "run_history.db"
    report = generator.validate_before_generation(
        roster_file=roster_path, invoice_file=invoice_path, run_history_db_path=db_path,
    )
    assert not any(i.category == "duplicate_invoice" for i in report.issues)


def test_generation_records_run_and_second_batch_is_flagged(generator, minimal_template, roster_and_invoice, tmp_path):
    roster_path, invoice_path = roster_and_invoice
    db_path = tmp_path / "run_history.db"

    summary = generator.generate_invoices(
        roster_file=roster_path, invoice_file=invoice_path, template_file=minimal_template,
        output_dir=str(tmp_path / "output1"), generate_csv=True, export_format="pdf",
        run_history_db_path=db_path,
    )
    assert summary.total_processed == 1

    later_generator = PatientInvoiceGenerator(
        amount_due_strategy="auto", statement_date="2026-07-16", clinic_config=TEST_CLINIC_CONFIG,
    )
    report = later_generator.validate_before_generation(
        roster_file=roster_path, invoice_file=invoice_path, run_history_db_path=db_path,
    )
    dup_issues = [i for i in report.issues if i.category == "duplicate_invoice"]
    assert len(dup_issues) == 1
    assert dup_issues[0].patient_name == "Guzhavin, Alex"
    assert "2026-01-12" in dup_issues[0].detail
    assert "2026-06-15" in dup_issues[0].detail


def test_skip_patient_names_excludes_from_generation(generator, minimal_template, roster_and_invoice, tmp_path):
    roster_path, invoice_path = roster_and_invoice
    db_path = tmp_path / "run_history.db"

    summary = generator.generate_invoices(
        roster_file=roster_path, invoice_file=invoice_path, template_file=minimal_template,
        output_dir=str(tmp_path / "output"), generate_csv=False, export_format="pdf",
        run_history_db_path=db_path, skip_patient_names={"Guzhavin, Alex"},
    )
    assert summary.total_processed == 0
    assert summary.total_skipped == 1
    assert summary.skipped_patients == [("Guzhavin, Alex", "Skipped by user (duplicate invoice)")]

    # Skipping must not record a (nonexistent) run in history.
    key = run_history.patient_key("1", "Alex", "Guzhavin")
    assert run_history.find_overlapping_runs(key, "2026-01-01", "2026-01-31", db_path=db_path) == []


def test_non_overlapping_second_run_is_not_flagged(generator, minimal_template, roster_and_invoice, tmp_path):
    """A later invoice for a different (non-overlapping) service period for
    the same patient should not be treated as a duplicate."""
    roster_path, invoice_path = roster_and_invoice
    db_path = tmp_path / "run_history.db"

    generator.generate_invoices(
        roster_file=roster_path, invoice_file=invoice_path, template_file=minimal_template,
        output_dir=str(tmp_path / "output1"), generate_csv=False, export_format="pdf",
        run_history_db_path=db_path,
    )

    # New invoice workbook: same patient, a later, non-overlapping visit.
    invoice_rows2 = [dict(Name="Guzhavin, Alex", **{
        "Visit Date": "2026-03-01", "Total amount": 200, "Copay": 40, "Paid": 0,
        "Previous Balance": "", "Type Of Service": "Psychotherapy",
    })]
    invoice_path2 = tmp_path / "invoice2.xlsx"
    pd.DataFrame(invoice_rows2).to_excel(invoice_path2, sheet_name="Sheet1", index=False)

    report = generator.validate_before_generation(
        roster_file=roster_path, invoice_file=str(invoice_path2), run_history_db_path=db_path,
    )
    assert not any(i.category == "duplicate_invoice" for i in report.issues)
