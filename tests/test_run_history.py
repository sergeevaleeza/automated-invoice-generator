"""Tests for run_history.py — the local SQLite store used for
duplicate-invoice detection. All patient data here is synthetic."""
import pytest

import run_history


@pytest.fixture
def db_path(tmp_path):
    return tmp_path / "run_history.db"


def test_patient_key_prefers_prn():
    assert run_history.patient_key("123", "Alex", "Guzhavin") == "prn:123"
    # Case/whitespace-insensitive
    assert run_history.patient_key(" 123 ", "Alex", "Guzhavin") == "prn:123"


def test_patient_key_falls_back_to_name_when_no_prn():
    assert run_history.patient_key(None, "Alex", "Guzhavin") == "name:alex_guzhavin"
    assert run_history.patient_key("", "Alex", "Guzhavin") == "name:alex_guzhavin"
    assert run_history.patient_key("nan", "Alex", "Guzhavin") == "name:alex_guzhavin"


def test_record_and_find_exact_overlap(db_path):
    key = run_history.patient_key("1", "Alex", "Guzhavin")
    run_history.record_invoice_run(key, "Guzhavin, Alex", "2026-01-12", "2026-04-30",
                                    "2026-06-15", ["a.pdf"], db_path=db_path)
    overlaps = run_history.find_overlapping_runs(key, "2026-01-12", "2026-04-30", db_path=db_path)
    assert len(overlaps) == 1
    assert overlaps[0].patient_display_name == "Guzhavin, Alex"
    assert overlaps[0].filenames == ["a.pdf"]


@pytest.mark.parametrize("new_start,new_end,should_overlap", [
    ("2026-03-01", "2026-05-01", True),   # partial overlap, extends past the end
    ("2025-12-01", "2026-02-01", True),   # partial overlap, starts before the beginning
    ("2026-02-01", "2026-03-01", True),   # fully contained within
    ("2025-01-01", "2027-01-01", True),   # fully encompasses
    ("2026-01-12", "2026-01-12", True),   # single day at the exact start
    ("2026-05-01", "2026-06-01", False),  # entirely after, no overlap
    ("2025-01-01", "2025-12-31", False),  # entirely before, no overlap
])
def test_overlap_detection_boundary_cases(db_path, new_start, new_end, should_overlap):
    key = run_history.patient_key("1", "Alex", "Guzhavin")
    run_history.record_invoice_run(key, "Guzhavin, Alex", "2026-01-12", "2026-04-30",
                                    "2026-06-15", ["a.pdf"], db_path=db_path)
    overlaps = run_history.find_overlapping_runs(key, new_start, new_end, db_path=db_path)
    assert bool(overlaps) == should_overlap


def test_different_patients_do_not_cross_contaminate(db_path):
    key_a = run_history.patient_key("1", "Alex", "Guzhavin")
    key_b = run_history.patient_key("2", "Bob", "Smith")
    run_history.record_invoice_run(key_a, "Guzhavin, Alex", "2026-01-12", "2026-04-30",
                                    "2026-06-15", ["a.pdf"], db_path=db_path)
    overlaps = run_history.find_overlapping_runs(key_b, "2026-01-12", "2026-04-30", db_path=db_path)
    assert overlaps == []


def test_most_recent_overlap_returned_first(db_path):
    key = run_history.patient_key("1", "Alex", "Guzhavin")
    run_history.record_invoice_run(key, "Guzhavin, Alex", "2026-01-01", "2026-01-31",
                                    "2026-02-01", ["jan.pdf"], db_path=db_path)
    run_history.record_invoice_run(key, "Guzhavin, Alex", "2026-01-01", "2026-01-31",
                                    "2026-03-01", ["jan_v2.pdf"], db_path=db_path)
    overlaps = run_history.find_overlapping_runs(key, "2026-01-01", "2026-01-31", db_path=db_path)
    assert len(overlaps) == 2
    assert overlaps[0].invoice_date == "2026-03-01"  # most recent first
