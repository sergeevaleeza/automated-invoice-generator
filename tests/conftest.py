"""Shared test fixtures. All patient data here is synthetic — never use real
patient records in tests or fixtures (see tests/fixtures/README or repo
CHANGELOG for why this matters)."""
from datetime import datetime
from pathlib import Path

import pandas as pd
import pytest

from invoice_models import PatientData

FIXTURES_DIR = Path(__file__).parent / "fixtures"
GOLDEN_INVOICE_PATH = FIXTURES_DIR / "Example_2026_Invoice_07162026.xlsx"


@pytest.fixture
def golden_invoice_inputs():
    """Synthetic patient/invoice data matching tests/fixtures/Example_2026_Invoice_07162026.xlsx
    exactly (same name, address, dates, service lines, and amounts), so a
    generated invoice can be compared cell-for-cell against the approved
    golden file."""
    patient = PatientData(
        prn="9999", first_name="Papa", last_name="Bo", dob="",
        address_line1="1 Hope Ave", address_line2="",
        city="San Francisco", state="CA", postal_code="94000",
    )

    rows = [
        ("2026-01-12", "Psychotherapy", 0, 50),
        ("2026-01-29", "Psychotherapy", 0, 50),
        ("2026-02-12", "Psychotherapy", 500, 250),
        ("2026-03-05", "Psychotherapy", 0, 250),
        ("2026-03-31", "Med Management", 0, 157),
        ("2026-04-02", "Psychotherapy", 0, 122),
        ("2026-04-14", "Med Management", 0, 220),
        ("2026-04-30", "Psychotherapy", 0, 122),
    ]
    patient_df = pd.DataFrame(rows, columns=["visit_date", "type_of_service", "paid", "copay"])
    patient_df["previous_balance"] = 0

    return dict(
        patient=patient,
        lines=[],
        total_due=721.0,
        patient_df=patient_df,
        statement_date=datetime(2026, 7, 16),
        payment_due_date=datetime(2026, 8, 17),
        has_cpt=False,
    )
