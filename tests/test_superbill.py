"""Tests for the superbill export: CPT/ICD-10 resolution logic
(PatientInvoiceGenerator.resolve_superbill_service_lines() /
resolve_default_icd10_codes()) and the PDF generator. All patient data
here is synthetic."""
import base64
import re
import zlib
from pathlib import Path

import pandas as pd
import pytest

from complete_patient_invoice_generator import PatientInvoiceGenerator
from invoice_models import PatientData, SuperbillServiceLine, extract_embedded_cpt_code
from superbill_generator import generate_superbill_pdf
from tests.conftest import TEST_CLINIC_CONFIG


@pytest.fixture
def clinic():
    return {
        **TEST_CLINIC_CONFIG,
        "default_cpt_by_service_type": {"psychotherapy": "90837", "med management": "99213"},
        "default_icd10_codes": ["F41.1"],
    }


@pytest.fixture
def generator(clinic):
    return PatientInvoiceGenerator(amount_due_strategy="auto", statement_date="2026-07-16",
                                    clinic_config=clinic)


def test_extract_embedded_cpt_code():
    assert extract_embedded_cpt_code("Med Management (CPT Code 99213)") == "99213"
    assert extract_embedded_cpt_code("Psychotherapy") is None
    assert extract_embedded_cpt_code("") is None


class TestCptResolution:
    def test_explicit_workbook_column_wins(self, generator):
        df = pd.DataFrame([{
            "visit_date": "2026-01-01", "type_of_service": "Consult", "total_amount": 100,
            "paid": 0, "cpt_code": "99999", "icd10_code": "",
        }])
        lines = generator.resolve_superbill_service_lines(df)
        assert lines[0].cpt_code == "99999"

    def test_embedded_code_used_when_no_workbook_column(self, generator):
        df = pd.DataFrame([{
            "visit_date": "2026-01-01", "type_of_service": "Med Management (CPT Code 99213)",
            "total_amount": 150, "paid": 30, "cpt_code": "", "icd10_code": "",
        }])
        lines = generator.resolve_superbill_service_lines(df)
        assert lines[0].cpt_code == "99213"

    def test_default_mapping_used_as_last_resort(self, generator):
        df = pd.DataFrame([{
            "visit_date": "2026-01-01", "type_of_service": "Psychotherapy",
            "total_amount": 200, "paid": 50, "cpt_code": "", "icd10_code": "",
        }])
        lines = generator.resolve_superbill_service_lines(df)
        assert lines[0].cpt_code == "90837"

    def test_blank_when_nothing_resolves(self, generator):
        df = pd.DataFrame([{
            "visit_date": "2026-01-01", "type_of_service": "Unmapped Service",
            "total_amount": 100, "paid": 0, "cpt_code": "", "icd10_code": "",
        }])
        lines = generator.resolve_superbill_service_lines(df)
        assert lines[0].cpt_code == ""

    def test_charge_and_payment_carried_through(self, generator):
        df = pd.DataFrame([{
            "visit_date": "2026-01-01", "type_of_service": "Psychotherapy",
            "total_amount": 200, "paid": 50, "cpt_code": "", "icd10_code": "",
        }])
        lines = generator.resolve_superbill_service_lines(df)
        assert lines[0].charge == 200.0
        assert lines[0].payment == 50.0
        assert lines[0].service_date == "01/01/2026"


class TestIcd10Resolution:
    def test_workbook_codes_take_priority(self, generator):
        df = pd.DataFrame([
            {"visit_date": "2026-01-01", "type_of_service": "X", "total_amount": 1,
             "paid": 0, "cpt_code": "", "icd10_code": "Z71.1"},
            {"visit_date": "2026-01-08", "type_of_service": "X", "total_amount": 1,
             "paid": 0, "cpt_code": "", "icd10_code": "F32.9"},
        ])
        codes = generator.resolve_default_icd10_codes(df)
        assert codes == ["F32.9", "Z71.1"]  # sorted, unique

    def test_falls_back_to_clinic_default_when_workbook_empty(self, generator):
        df = pd.DataFrame([{
            "visit_date": "2026-01-01", "type_of_service": "X", "total_amount": 1,
            "paid": 0, "cpt_code": "", "icd10_code": "",
        }])
        assert generator.resolve_default_icd10_codes(df) == ["F41.1"]

    def test_empty_when_no_workbook_data_and_no_clinic_default(self):
        clinic_without_default = {k: v for k, v in TEST_CLINIC_CONFIG.items() if k != "default_icd10_codes"}
        gen = PatientInvoiceGenerator(amount_due_strategy="auto", statement_date="2026-07-16",
                                       clinic_config=clinic_without_default)
        df = pd.DataFrame([{
            "visit_date": "2026-01-01", "type_of_service": "X", "total_amount": 1,
            "paid": 0, "cpt_code": "", "icd10_code": "",
        }])
        assert gen.resolve_default_icd10_codes(df) == []


class TestSuperbillPdf:
    def _patient(self):
        return PatientData(prn="1001", first_name="Alice", last_name="Anderson", dob="05/12/1980",
                            address_line1="123 Main St", address_line2="", city="Burlingame",
                            state="CA", postal_code="94010")

    def test_generates_single_page_pdf(self, clinic, tmp_path):
        from datetime import datetime
        lines = [
            SuperbillServiceLine("01/12/2026", "90837", "Psychotherapy", 200.0, 50.0),
            SuperbillServiceLine("01/19/2026", "99213", "Med Management", 150.0, 30.0),
        ]
        out_path = tmp_path / "superbill.pdf"
        generate_superbill_pdf(self._patient(), clinic, lines, ["F41.1", "F33.1"],
                                datetime(2026, 7, 16), out_path)
        data = out_path.read_bytes()
        assert data[:4] == b"%PDF"
        assert len(re.findall(rb"/Type\s*/Page\b", data)) == 1

    def test_pdf_contains_key_fields(self, clinic, tmp_path):
        """Sanity check that the identifying fields actually appear in the
        PDF's text streams. ReportLab compresses content streams with
        FlateDecode (+ ASCII85Decode), so this decodes each stream before
        searching rather than grepping the raw file bytes."""
        from datetime import datetime
        lines = [SuperbillServiceLine("01/12/2026", "90837", "Psychotherapy", 200.0, 50.0)]
        out_path = tmp_path / "superbill.pdf"
        generate_superbill_pdf(self._patient(), clinic, lines, ["F41.1"], datetime(2026, 7, 16), out_path)
        data = out_path.read_bytes()

        full_text = b""
        for match in re.finditer(rb"stream\r?\n(.*?)endstream", data, re.DOTALL):
            stream_data = match.group(1).strip(b"\r\n")
            if stream_data.endswith(b"~>"):
                stream_data = stream_data[:-2]
            try:
                full_text += zlib.decompress(base64.a85decode(stream_data, adobe=False))
            except (ValueError, zlib.error):
                continue

        expected = [b"Anderson", clinic["npi"].encode(), clinic["ein"].encode(), b"90837", b"F41.1"]
        for value in expected:
            assert value in full_text, f"{value!r} not found in decoded PDF text streams"

    def test_empty_icd10_codes_does_not_crash(self, clinic, tmp_path):
        from datetime import datetime
        lines = [SuperbillServiceLine("01/12/2026", "90837", "Psychotherapy", 200.0, 50.0)]
        out_path = tmp_path / "superbill.pdf"
        generate_superbill_pdf(self._patient(), clinic, lines, [], datetime(2026, 7, 16), out_path)
        assert out_path.exists()
