#!/usr/bin/env python3
"""
Loads clinic identity (name, doctor, addresses, EIN/NPI, payment info, etc.)
from clinic_config.json — kept out of source control (see .gitignore) so a
real practice's business identity never ends up hardcoded in a public repo.

clinic_config.example.json is the committed template: copy it to
clinic_config.json and fill in your practice's real details.
"""
import json
from pathlib import Path
from typing import Optional

DEFAULT_CONFIG_PATH = Path(__file__).parent / "clinic_config.json"
EXAMPLE_CONFIG_PATH = Path(__file__).parent / "clinic_config.example.json"

REQUIRED_KEYS = [
    "clinic_name", "doctor_name", "specialty", "ein", "npi",
    "office_address", "mailing_address", "email", "website",
    "pricing_page_url", "phone", "zelle_email", "check_payable_to",
    "provider_name_for_signature",
]


class ClinicConfigError(Exception):
    """Raised when clinic_config.json is missing or incomplete. The message
    is meant to be shown directly to the user (e.g. via st.error) — it never
    contains patient data, only setup guidance."""


def load_clinic_config(path: Optional[Path] = None) -> dict:
    """Load and validate clinic identity config. Deliberately does NOT fall
    back to clinic_config.example.json's placeholder values on its own —
    silently generating real-looking invoices with fake clinic identity
    would be worse than failing loudly. Callers that need placeholder data
    on purpose (e.g. tests) should pass example data explicitly instead."""
    config_path = path or DEFAULT_CONFIG_PATH
    if not config_path.exists():
        raise ClinicConfigError(
            f"{config_path.name} not found. Copy clinic_config.example.json to "
            f"{config_path.name} and fill in your practice's real details."
        )
    try:
        data = json.loads(config_path.read_text())
    except json.JSONDecodeError as e:
        raise ClinicConfigError(f"{config_path.name} is not valid JSON: {e}")

    missing = [k for k in REQUIRED_KEYS if not data.get(k)]
    if missing:
        raise ClinicConfigError(
            f"{config_path.name} is missing required field(s): {', '.join(missing)}"
        )
    return data
