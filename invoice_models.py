#!/usr/bin/env python3
"""
Shared data structures and formatting helpers for invoice generation.
Used by both the PDF (complete_patient_invoice_generator.py) and Excel
(excel_invoice_generator.py) invoice generators so neither duplicates them.
"""
from dataclasses import dataclass
from datetime import datetime
from typing import List, Tuple, Optional
import logging

import pandas as pd


@dataclass
class PatientData:
    """Patient information from roster"""
    prn: str
    first_name: str
    last_name: str
    dob: str
    address_line1: str
    address_line2: str
    city: str
    state: str
    postal_code: str


@dataclass
class InvoiceLine:
    """Individual service line from invoice"""
    service_date: str
    description: str
    amount: float
    is_previous_balance: bool = False
    is_credit: bool = False  # True for negative previous balance (overpayment shown as credit)


@dataclass
class ProcessingSummary:
    """Summary of processing results"""
    processed_patients: List[str]
    skipped_patients: List[Tuple[str, str]]  # (name, reason)
    errors: List[Tuple[str, str]]  # (patient_name, error)
    total_processed: int
    total_skipped: int
    total_errors: int
    total_amount_due: float
    processing_date: str


def format_date_for_display(date_value, logger: Optional[logging.Logger] = None) -> str:
    """Format date value to MM/DD/YYYY format. Shared by PDF and Excel invoice generators."""
    if pd.isna(date_value) or date_value == '' or date_value is None:
        return ''

    try:
        # Convert to string first
        date_str = str(date_value)

        # Try to parse as datetime if it's not already
        if isinstance(date_value, str):
            # Handle different possible input formats
            for fmt in ['%Y-%m-%d', '%m/%d/%Y', '%m-%d-%Y', '%Y/%m/%d']:
                try:
                    parsed_date = datetime.strptime(date_str.split()[0], fmt)  # Take only date part
                    return parsed_date.strftime('%m/%d/%Y')
                except ValueError:
                    continue
        elif hasattr(date_value, 'strftime'):
            # If it's already a datetime object
            return date_value.strftime('%m/%d/%Y')

        # If we can't parse it, try to extract date from string
        date_part = date_str.split()[0] if ' ' in date_str else date_str

        # Try pandas to_datetime as last resort
        parsed_date = pd.to_datetime(date_part, errors='coerce')
        if not pd.isna(parsed_date):
            return parsed_date.strftime('%m/%d/%Y')

        # If all else fails, return the original string
        return date_part

    except Exception as e:
        if logger:
            logger.warning(f"Could not format date '{date_value}': {e}")
        return str(date_value) if date_value else ''
