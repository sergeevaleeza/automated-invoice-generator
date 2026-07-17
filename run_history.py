#!/usr/bin/env python3
"""
Local run-history store for duplicate-invoice detection.

Records, per successfully-generated invoice: the patient's identity, the
service-date range it covered, the statement (invoice) date, and the
output filenames — so a later batch run can warn if it's about to
re-invoice the same patient for a service-date range that overlaps one
already invoiced.

Stored in data/run_history.db — gitignored, since it's operational data
about real patients, not something to commit. The data/ directory is
created on first use if it doesn't exist.

Caveat: this is a local SQLite file. On Streamlit Cloud (or any ephemeral
filesystem deploy), it will NOT persist across app reboots/redeploys —
only within a single running container's lifetime. See docs/DEPLOY.md.
"""
import json
import sqlite3
from contextlib import closing
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import List, Optional

DEFAULT_DB_PATH = Path(__file__).parent / "data" / "run_history.db"


@dataclass
class RunRecord:
    """One previously-recorded invoice run for a single patient."""
    id: int
    patient_key: str
    patient_display_name: str
    service_date_start: str  # ISO date string, e.g. "2026-01-12"
    service_date_end: str
    invoice_date: str  # ISO date string — the statement date used
    filenames: List[str]
    created_at: str  # ISO timestamp


def patient_key(prn: Optional[str], first_name: str, last_name: str) -> str:
    """Stable identity key for duplicate matching. Prefers PRN (matches the
    roster, most reliable and stable across differently-formatted invoice
    workbooks); falls back to a normalized name for patients with no
    roster match."""
    prn_clean = (prn or "").strip().lower()
    if prn_clean and prn_clean != "nan":
        return f"prn:{prn_clean}"
    return f"name:{first_name.strip().lower()}_{last_name.strip().lower()}"


def _connect(db_path: Path) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(db_path))
    conn.execute("""
        CREATE TABLE IF NOT EXISTS invoice_runs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            patient_key TEXT NOT NULL,
            patient_display_name TEXT NOT NULL,
            service_date_start TEXT NOT NULL,
            service_date_end TEXT NOT NULL,
            invoice_date TEXT NOT NULL,
            filenames TEXT NOT NULL,
            created_at TEXT NOT NULL
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_invoice_runs_patient_key ON invoice_runs(patient_key)")
    conn.commit()
    return conn


def find_overlapping_runs(patient_key_value: str, service_date_start: str, service_date_end: str,
                           db_path: Path = DEFAULT_DB_PATH) -> List[RunRecord]:
    """Find prior recorded runs for this patient whose service-date range
    overlaps [service_date_start, service_date_end] (inclusive, ISO date
    strings 'YYYY-MM-DD'), most recent first."""
    with closing(_connect(db_path)) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute("""
            SELECT * FROM invoice_runs
            WHERE patient_key = ?
              AND NOT (service_date_end < ? OR service_date_start > ?)
            ORDER BY invoice_date DESC, created_at DESC
        """, (patient_key_value, service_date_start, service_date_end)).fetchall()
        return [
            RunRecord(
                id=r["id"], patient_key=r["patient_key"], patient_display_name=r["patient_display_name"],
                service_date_start=r["service_date_start"], service_date_end=r["service_date_end"],
                invoice_date=r["invoice_date"], filenames=json.loads(r["filenames"]), created_at=r["created_at"],
            )
            for r in rows
        ]


def record_invoice_run(patient_key_value: str, patient_display_name: str,
                        service_date_start: str, service_date_end: str, invoice_date: str,
                        filenames: List[str], db_path: Path = DEFAULT_DB_PATH) -> None:
    """Record a successfully-generated invoice for later overlap checks."""
    with closing(_connect(db_path)) as conn:
        conn.execute("""
            INSERT INTO invoice_runs
            (patient_key, patient_display_name, service_date_start, service_date_end, invoice_date, filenames, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            patient_key_value, patient_display_name, service_date_start, service_date_end,
            invoice_date, json.dumps(filenames), datetime.now().isoformat(timespec="seconds"),
        ))
        conn.commit()
