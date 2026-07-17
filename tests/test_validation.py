"""Tests for PatientInvoiceGenerator.validate_before_generation() — the
pre-flight scan that runs before any invoice files are generated. All
patient/roster data here is synthetic."""
import pandas as pd
import pytest

from complete_patient_invoice_generator import PatientInvoiceGenerator
from tests.conftest import TEST_CLINIC_CONFIG


@pytest.fixture
def generator():
    return PatientInvoiceGenerator(
        amount_due_strategy="auto", statement_date="2026-07-16",
        clinic_config=TEST_CLINIC_CONFIG,
    )


@pytest.fixture
def roster_and_invoice(tmp_path):
    roster_rows = [
        dict(**{
            "Patient Identifier": "2001", "Patient First Name": "Roberta", "Patient Last Name": "Nolan",
            "DOB": "1980-01-01", "Address Line 1": "", "Address Line 2": "",
            "City": "", "State": "CA", "Postal Code": "9401",  # missing addr1/city, bad zip
        }),
        dict(**{
            "Patient Identifier": "2002", "Patient First Name": "Sam", "Patient Last Name": "Weller",
            "DOB": "1975-05-12", "Address Line 1": "1 Elm St", "Address Line 2": "",
            "City": "Fresno", "State": "CA", "Postal Code": "93701",
        }),
        dict(**{
            "Patient Identifier": "2003", "Patient First Name": "Samuel", "Patient Last Name": "Weller",
            "DOB": "1990-05-12", "Address Line 1": "2 Elm St", "Address Line 2": "",
            "City": "Fresno", "State": "CA", "Postal Code": "93701",
        }),
    ]
    roster_path = tmp_path / "roster.csv"
    pd.DataFrame(roster_rows).to_csv(roster_path, index=False)

    invoice_rows = [
        # Malformed address (exact match)
        dict(Name="Nolan, Roberta", **{
            "Visit Date": "2026-01-10", "Total amount": 200, "Copay": 40, "Paid": 0,
            "Previous Balance": "", "Type Of Service": "Psychotherapy",
        }),
        # Low-confidence fuzzy match
        dict(Name="Weler, Sammy", **{
            "Visit Date": "2026-01-11", "Total amount": 200, "Copay": 40, "Paid": 0,
            "Previous Balance": "", "Type Of Service": "Psychotherapy",
        }),
        # No roster match at all
        dict(Name="Zephyrine, Quilla", **{
            "Visit Date": "2026-01-12", "Total amount": 200, "Copay": 40, "Paid": 0,
            "Previous Balance": "", "Type Of Service": "Psychotherapy",
        }),
        # Missing service date
        dict(Name="Weller, Sam", **{
            "Visit Date": "", "Total amount": 200, "Copay": 40, "Paid": 0,
            "Previous Balance": "", "Type Of Service": "Psychotherapy",
        }),
        # Charge with no description
        dict(Name="Weller, Sam", **{
            "Visit Date": "2026-01-13", "Total amount": 200, "Copay": 40, "Paid": 0,
            "Previous Balance": "", "Type Of Service": "",
        }),
        # Ambiguous match (score >= 0.85 for two roster entries)
        dict(Name="Weller, Samu", **{
            "Visit Date": "2026-01-14", "Total amount": 200, "Copay": 40, "Paid": 0,
            "Previous Balance": "", "Type Of Service": "Psychotherapy",
        }),
        # Negative (credit) balance, own patient group
        dict(Name="Weller, Samuel", **{
            "Visit Date": "2026-01-15", "Total amount": 0, "Copay": 0, "Paid": 0,
            "Previous Balance": -50, "Type Of Service": "Psychotherapy",
        }),
    ]
    invoice_path = tmp_path / "invoice.xlsx"
    pd.DataFrame(invoice_rows).to_excel(invoice_path, sheet_name="Sheet1", index=False)

    return str(roster_path), str(invoice_path)


def test_validation_finds_all_issue_categories(generator, roster_and_invoice):
    roster_path, invoice_path = roster_and_invoice
    report = generator.validate_before_generation(roster_file=roster_path, invoice_file=invoice_path)

    categories_found = {issue.category for issue in report.issues}
    assert categories_found == {
        "unmatched_patient", "low_confidence_match", "ambiguous_match",
        "malformed_address", "missing_service_date", "missing_description",
        "negative_balance",
    }


def test_validation_counts_and_severities(generator, roster_and_invoice):
    roster_path, invoice_path = roster_and_invoice
    report = generator.validate_before_generation(roster_file=roster_path, invoice_file=invoice_path)

    assert report.total_patient_groups == 6
    # unmatched_patient and missing_service_date are errors; the rest are warnings
    errors = {i.category for i in report.issues if i.severity == "error"}
    warnings = {i.category for i in report.issues if i.severity == "warning"}
    assert errors == {"unmatched_patient", "missing_service_date"}
    assert "malformed_address" in warnings
    assert "negative_balance" in warnings


def test_clean_data_produces_no_issues(generator, tmp_path):
    roster_rows = [dict(**{
        "Patient Identifier": "1", "Patient First Name": "Clean", "Patient Last Name": "Patient",
        "DOB": "1980-01-01", "Address Line 1": "1 Main St", "Address Line 2": "",
        "City": "Testville", "State": "CA", "Postal Code": "94000",
    })]
    roster_path = tmp_path / "roster.csv"
    pd.DataFrame(roster_rows).to_csv(roster_path, index=False)

    invoice_rows = [dict(Name="Patient, Clean", **{
        "Visit Date": "2026-01-10", "Total amount": 200, "Copay": 40, "Paid": 0,
        "Previous Balance": "", "Type Of Service": "Psychotherapy",
    })]
    invoice_path = tmp_path / "invoice.xlsx"
    pd.DataFrame(invoice_rows).to_excel(invoice_path, sheet_name="Sheet1", index=False)

    report = generator.validate_before_generation(roster_file=str(roster_path), invoice_file=str(invoice_path))
    assert report.issues == []
    assert report.error_count == 0
    assert report.warning_count == 0


def test_validation_report_text_export(generator, roster_and_invoice):
    roster_path, invoice_path = roster_and_invoice
    report = generator.validate_before_generation(roster_file=roster_path, invoice_file=invoice_path)
    text = PatientInvoiceGenerator._generate_validation_report_text(report)

    assert "PRE-FLIGHT VALIDATION REPORT" in text
    assert f"Errors: {report.error_count}" in text
    for issue in report.issues:
        assert issue.patient_name in text
