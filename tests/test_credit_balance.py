"""Regression tests for credit-balance (overpaid) patient handling.
generate_invoices() must NOT skip these patients — they should still
receive an invoice showing their credit (via the "Previous Balance
(Overpaid)" line item) and $0.00 due, not be excluded from the batch.
validate_before_generation()'s negative_balance check flags these for
staff awareness without implying they'll be skipped. All patient/roster
data here is synthetic."""
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


def test_pure_credit_balance_is_invoiced_not_skipped(generator, minimal_template, tmp_path):
    """A patient with a negative previous balance and no offsetting charges
    this period must still be invoiced — not skipped — showing $0.00 due."""
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
    assert summary.total_skipped == 0
    assert summary.skipped_patients == []
    assert summary.total_processed == 1
    assert "Due: $0.00" in summary.processed_patients[0]
    assert (tmp_path / "output" / "Patient_Credit_1").exists()


def test_credit_balance_line_item_shows_the_credit(generator):
    """The invoice's line items must reflect the credit amount, not just a
    $0.00 total, so the patient can see why they owe nothing."""
    df = pd.DataFrame([{
        "visit_date": "2026-01-10", "total_amount": 0, "copay": 0, "paid": 0,
        "previous_balance": -215.35, "type_of_service": "Psychotherapy",
    }])
    lines, total_due, _ = generator._generate_invoice_lines(df)
    assert total_due == -215.35  # raw value, unfloored — see _generate_invoice_lines docstring

    credit_lines = [l for l in lines if l.is_previous_balance and l.is_credit]
    assert len(credit_lines) == 1
    assert credit_lines[0].amount == pytest.approx(215.35)
    assert credit_lines[0].description == "Previous Balance (Overpaid)"


def test_negative_previous_balance_offset_by_charges(generator, minimal_template, tmp_path):
    """A negative previous balance alone doesn't mean the patient nets to a
    credit — same-period charges can bring the total back to positive."""
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


def test_validation_flags_credit_without_implying_skip(generator, tmp_path):
    """validate_before_generation() should flag a net credit balance for
    staff awareness, but the message must not claim the patient is
    skipped — they aren't."""
    roster_path, invoice_path = _write_roster_invoice(tmp_path, [
        dict(Name="Patient, Credit", **{
            "Visit Date": "2026-01-10", "Total amount": 0, "Copay": 0, "Paid": 0,
            "Previous Balance": -215.35, "Type Of Service": "Psychotherapy",
        }),
    ])
    report = generator.validate_before_generation(roster_file=roster_path, invoice_file=invoice_path)
    credit_issues = [i for i in report.issues if i.category == "negative_balance"]
    assert len(credit_issues) == 1
    assert "skip" not in credit_issues[0].detail.lower()
    assert "215.35" in credit_issues[0].detail

    # Offsetting case must not be flagged at all (the scenario the old
    # `previous_balance < 0` proxy check got wrong).
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
