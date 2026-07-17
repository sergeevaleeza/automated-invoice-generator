"""Regression tests for the credit-balance skip fix: generate_invoices()
must actually skip patients with a net negative balance instead of
generating a $0.00 invoice, and validate_before_generation()'s
negative_balance check must match that real condition exactly (not just
`previous_balance < 0` in isolation, which can be misleading when
same-period charges bring the net balance back to zero or positive).
All patient/roster data here is synthetic."""
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


def _write_roster_invoice(tmp_path, invoice_rows):
    roster_rows = [dict(**{
        "Patient Identifier": "1", "Patient First Name": "Credit", "Patient Last Name": "Patient",
        "DOB": "1980-01-01", "Address Line 1": "1 Main St", "Address Line 2": "",
        "City": "Testville", "State": "CA", "Postal Code": "94000",
    })]
    roster_path = tmp_path / "roster.csv"
    pd.DataFrame(roster_rows).to_csv(roster_path, index=False)

    invoice_path = tmp_path / "invoice.xlsx"
    pd.DataFrame(invoice_rows).to_excel(invoice_path, sheet_name="Sheet1", index=False)
    return str(roster_path), str(invoice_path)


def test_pure_credit_balance_is_skipped_not_zero_invoiced(generator, minimal_template, tmp_path):
    """A patient with a negative previous balance and no offsetting charges
    this period must be skipped entirely, not given a $0.00 invoice."""
    roster_path, invoice_path = _write_roster_invoice(tmp_path, [
        dict(Name="Patient, Credit", **{
            "Visit Date": "2026-01-10", "Total amount": 0, "Copay": 0, "Paid": 0,
            "Previous Balance": -215.35, "Type Of Service": "Psychotherapy",
        }),
    ])
    summary = generator.generate_invoices(
        roster_file=roster_path, invoice_file=invoice_path, template_file=minimal_template,
        output_dir=str(tmp_path / "output"), generate_csv=False, export_format="pdf",
    )
    assert summary.total_processed == 0
    assert summary.total_skipped == 1
    assert summary.skipped_patients == [("Patient, Credit", "Credit balance (overpaid)")]


def test_negative_previous_balance_offset_by_charges_is_not_skipped(generator, minimal_template, tmp_path):
    """A negative previous balance alone doesn't mean the patient nets to a
    credit — same-period charges can bring the total back to positive, in
    which case the patient should still be invoiced normally."""
    roster_path, invoice_path = _write_roster_invoice(tmp_path, [
        dict(Name="Patient, Credit", **{
            "Visit Date": "2026-01-10", "Total amount": 300, "Copay": 300, "Paid": 0,
            "Previous Balance": -50, "Type Of Service": "Psychotherapy",
        }),
    ])
    summary = generator.generate_invoices(
        roster_file=roster_path, invoice_file=invoice_path, template_file=minimal_template,
        output_dir=str(tmp_path / "output"), generate_csv=False, export_format="pdf",
    )
    assert summary.total_skipped == 0
    assert summary.total_processed == 1
    # net = 300 (copay) - 50 (previous credit) - 0 (paid) = 250
    assert "Due: $250.00" in summary.processed_patients[0]


def test_validation_matches_real_skip_condition(generator, tmp_path):
    """validate_before_generation()'s negative_balance check must agree
    with generate_invoices()'s actual skip decision in both directions."""
    # Case 1: pure credit -> validation flags it, generation skips it.
    roster_path, invoice_path = _write_roster_invoice(tmp_path, [
        dict(Name="Patient, Credit", **{
            "Visit Date": "2026-01-10", "Total amount": 0, "Copay": 0, "Paid": 0,
            "Previous Balance": -215.35, "Type Of Service": "Psychotherapy",
        }),
    ])
    report = generator.validate_before_generation(roster_file=roster_path, invoice_file=invoice_path)
    assert any(i.category == "negative_balance" for i in report.issues)

    # Case 2: negative previous balance fully offset by this period's
    # charges -> validation must NOT flag it (this is the exact scenario
    # the old `previous_balance < 0` proxy check got wrong).
    tmp_path2 = tmp_path / "case2"
    tmp_path2.mkdir()
    roster_path2, invoice_path2 = _write_roster_invoice(tmp_path2, [
        dict(Name="Patient, Credit", **{
            "Visit Date": "2026-01-10", "Total amount": 300, "Copay": 300, "Paid": 0,
            "Previous Balance": -50, "Type Of Service": "Psychotherapy",
        }),
    ])
    report2 = generator.validate_before_generation(roster_file=roster_path2, invoice_file=invoice_path2)
    assert not any(i.category == "negative_balance" for i in report2.issues)
