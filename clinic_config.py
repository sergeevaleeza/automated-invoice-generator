#!/usr/bin/env python3
"""
Loads clinic identity (name, doctor, addresses, EIN/NPI, payment info, etc.)
from one of two sources, in order:

  (a) clinic_config.json in the project root, if present (local dev — kept
      out of source control, see .gitignore).
  (b) a [clinic_config] table in Streamlit secrets, if present (Streamlit
      Cloud deploys — clinic_config.json is gitignored so it doesn't exist
      there; see docs/DEPLOY.md for the exact Secrets format to paste in).

Either way, a real practice's business identity never ends up hardcoded in
this public repo. clinic_config.example.json is the committed template for
option (a); docs/DEPLOY.md has the equivalent for option (b).
"""
import json
from pathlib import Path
from typing import Optional, Tuple

DEFAULT_CONFIG_PATH = Path(__file__).parent / "clinic_config.json"
EXAMPLE_CONFIG_PATH = Path(__file__).parent / "clinic_config.example.json"
STREAMLIT_SECRETS_KEY = "clinic_config"

REQUIRED_KEYS = [
    "clinic_name", "doctor_name", "specialty", "ein", "npi",
    "office_address", "mailing_address", "email", "website",
    "pricing_page_url", "phone", "zelle_email", "check_payable_to",
    "provider_name_for_signature",
]

# Optional fields, not in REQUIRED_KEYS — a config from before these
# features existed keeps working unchanged (they just stay off/empty):
#   show_qr / qr_image_path /
#     qr_content                — see qr_code.py
#   default_icd10_codes         — list of ICD-10 strings pre-filled (still
#                                  editable in the UI) on the superbill
#   default_cpt_by_service_type — dict mapping a lowercased type_of_service
#                                  value to a CPT code, used when the
#                                  invoice workbook has no CPT column and
#                                  none is embedded in the description
#                                  (see invoice_models.extract_embedded_cpt_code)


class ClinicConfigError(Exception):
    """Raised when clinic_config.json and Streamlit secrets are both
    missing/incomplete. The message is meant to be shown directly to the
    user (e.g. via st.error) — it never contains config or secret values,
    only key names and setup guidance."""


def _attrdict_to_dict(value):
    """Recursively convert a Streamlit AttrDict (or any dict-like object)
    into a plain dict. st.secrets returns AttrDict-like objects, not plain
    dicts, so callers can't rely on isinstance(x, dict) or plain-dict
    methods without this. Duck-types on .items() rather than importing
    Streamlit's internal AttrDict class directly."""
    if hasattr(value, "items"):
        return {k: _attrdict_to_dict(v) for k, v in value.items()}
    return value


def _load_from_streamlit_secrets() -> Optional[dict]:
    """Return the [clinic_config] table from st.secrets as a plain dict, or
    None if secrets aren't usable at all (no secrets.toml, no Streamlit
    runtime context, e.g. plain pytest) or the section is simply absent.
    Never raises for the "not configured" case — that's a valid fallback
    path, not an error; only load_clinic_config() decides when to give up."""
    try:
        import streamlit as st
        if STREAMLIT_SECRETS_KEY not in st.secrets:
            return None
        return _attrdict_to_dict(st.secrets[STREAMLIT_SECRETS_KEY])
    except Exception:
        return None


def _validate(data: dict, source_name: str) -> None:
    # Only key names ever appear in this message — never values.
    missing = [k for k in REQUIRED_KEYS if not data.get(k)]
    if missing:
        raise ClinicConfigError(
            f"Clinic config from {source_name} is missing required field(s): {', '.join(missing)}"
        )


def _resolve_clinic_config(path: Optional[Path] = None) -> Tuple[dict, str]:
    """Resolve clinic config from clinic_config.json if present, else a
    [clinic_config] table in Streamlit secrets. Returns (config_dict,
    source_label) where source_label is "local file" or "Streamlit
    secrets". Raises ClinicConfigError if neither source has a complete,
    valid config."""
    config_path = path or DEFAULT_CONFIG_PATH

    if config_path.exists():
        try:
            data = json.loads(config_path.read_text())
        except json.JSONDecodeError as e:
            raise ClinicConfigError(f"{config_path.name} is not valid JSON: {e}")
        _validate(data, config_path.name)
        return data, "local file"

    secrets_data = _load_from_streamlit_secrets()
    if secrets_data is not None:
        _validate(secrets_data, "Streamlit secrets")
        return secrets_data, "Streamlit secrets"

    raise ClinicConfigError(
        f"{config_path.name} not found, and no [clinic_config] section found in "
        "Streamlit secrets. For local development, copy clinic_config.example.json "
        f"to {config_path.name} and fill in your practice's real details. For a "
        "Streamlit Cloud deployment, add a [clinic_config] table to the app's "
        "Secrets instead — see docs/DEPLOY.md for the exact format to paste in."
    )


def load_clinic_config(path: Optional[Path] = None) -> dict:
    """Load clinic identity: local clinic_config.json if present, else a
    [clinic_config] table in Streamlit secrets, else raise
    ClinicConfigError. Same dict shape regardless of source — callers don't
    need to know or care which one was used."""
    config, _source = _resolve_clinic_config(path)
    return config


def get_clinic_config_source(path: Optional[Path] = None) -> str:
    """Which source load_clinic_config() would use right now: "local file"
    or "Streamlit secrets". Raises the same ClinicConfigError as
    load_clinic_config() if neither is available. Used for the UI status
    line so it's clear at a glance which config is active."""
    _config, source = _resolve_clinic_config(path)
    return source
