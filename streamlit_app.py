"""Client Email & Form App.

A simple Streamlit app that uses a Google Sheet tab ("client data") as its
client database, shows the client list, and triggers emails (single or bulk).
Each email contains a link to a form page served by this same app; responses
are written back to a "ClientResponses" tab in the same workbook.

Sheet tabs used (created on demand):
    client data        -> sender-maintained: S.No, Name, Email Id, Company Name, ...
    FormConfig         -> one row per question (auto-managed by this app)
    ClientResponses    -> one row per form submission (auto-managed by this app)

Run:
    .venv/bin/streamlit run client_email_app.py --server.port 8501

Config:
    config.json  -> google_service_account_path
    secrets.json -> hub_workbook_id, smtp_email, smtp_app_password,
                    form_base_url, form_token_secret

SMTP: Generic SMTP relay (Mailer91 / any SMTP server). Configure in secrets.json:
      smtp_server, smtp_port, smtp_use_tls, smtp_username, smtp_password,
      smtp_from_email, smtp_from_name. If missing, the sidebar asks for them.
"""
from __future__ import annotations

import hashlib
import json
import os
import smtplib
import ssl
import uuid

import requests

try:
    import certifi
    _CA_BUNDLE = certifi.where()
except Exception:  # pragma: no cover
    _CA_BUNDLE = None
from datetime import datetime, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Any, Optional

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from google.oauth2 import service_account
from google.auth.transport.requests import Request

# ── Page config ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Client Email & Form",
    page_icon="📧",
    layout="wide",
)

st.markdown(
    """
    <style>
      .block-container { padding-top: 1.5rem !important; max-width: 1200px; }
      .ce-tag {
        display: inline-flex; align-items: center; gap: 6px;
        font-size: .72rem; font-weight: 600; letter-spacing: .04em;
        text-transform: uppercase; color: #64748b;
        padding: 3px 10px; border-radius: 20px; background: #f1f5f9;
      }
      .ce-note {
        background: #f0f9ff; border: 1px solid #bae6fd; border-radius: 8px;
        padding: 10px 14px; font-size: .82rem; color: #075985;
      }
      .ce-warn {
        background: #fffbeb; border: 1px solid #fde68a; border-radius: 8px;
        padding: 10px 14px; font-size: .82rem; color: #92400e;
      }
      /* Client portal — unified form + signing */
      .cp-wrap {
        max-width: 680px; margin: 0 auto; padding: 0 1rem;
      }
      .cp-hero {
        background: linear-gradient(135deg, #1e293b 0%, #334155 100%);
        border-radius: 16px; padding: 32px 36px; margin-bottom: 24px;
        color: #f1f5f9;
      }
      .cp-hero h1 {
        font-size: 1.6rem; font-weight: 700; margin: 0 0 6px 0; color: #fff;
      }
      .cp-hero p {
        font-size: .9rem; color: #cbd5e1; margin: 0; line-height: 1.5;
      }
      .cp-card {
        background: #fff; border: 1px solid #e2e8f0; border-radius: 12px;
        padding: 28px 32px; margin-bottom: 20px;
        box-shadow: 0 1px 3px rgba(0,0,0,.06);
      }
      .cp-step-row {
        display: flex; gap: 0; margin-bottom: 28px; align-items: center;
      }
      .cp-step {
        display: flex; align-items: center; gap: 10px; flex: 1;
      }
      .cp-step-num {
        width: 32px; height: 32px; border-radius: 50%; display: flex;
        align-items: center; justify-content: center; font-weight: 700;
        font-size: .85rem; flex-shrink: 0;
      }
      .cp-step-num.active { background: #2563eb; color: #fff; }
      .cp-step-num.done { background: #16a34a; color: #fff; }
      .cp-step-num.idle { background: #e2e8f0; color: #64748b; }
      .cp-step-label {
        font-size: .85rem; font-weight: 600; color: #334155;
      }
      .cp-step-label.idle { color: #94a3b8; }
      .cp-step-line {
        flex: 1; height: 2px; background: #e2e8f0; margin: 0 8px;
      }
      .cp-step-line.done { background: #16a34a; }
      .cp-field-label {
        font-size: .88rem; font-weight: 600; color: #1e293b;
        margin-bottom: 6px; display: block;
      }
      .cp-field-hint {
        font-size: .76rem; color: #64748b; margin-bottom: 4px;
      }
      .cp-success-banner {
        background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 10px;
        padding: 16px 20px; margin-bottom: 20px;
        display: flex; align-items: center; gap: 12px;
      }
      .cp-success-banner .icon {
        width: 36px; height: 36px; border-radius: 50%; background: #16a34a;
        color: #fff; display: flex; align-items: center; justify-content: center;
        font-size: 1.2rem; flex-shrink: 0;
      }
      .cp-success-banner .text { font-size: .9rem; color: #15803d; }
      .cp-success-banner .text b { font-weight: 700; }
      .cp-sign-box {
        background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 12px;
        padding: 24px; margin-top: 8px;
      }
      .cp-doc-info {
        background: #fffbeb; border: 1px solid #fde68a; border-radius: 8px;
        padding: 12px 16px; margin-bottom: 16px; font-size: .84rem; color: #92400e;
      }
      .cp-doc-info b { color: #78350f; }
    </style>
    """,
    unsafe_allow_html=True,
)


# ── Paths & config ──────────────────────────────────────────────────────────
ROOT = Path(__file__).resolve().parent
CONFIG_PATH = ROOT / "config.json"
SECRETS_PATH = ROOT / "secrets.json"

CLIENT_TAB = "client data"
FORMCONFIG_TAB = "FormConfig"
RESPONSES_TAB = "ClientResponses"

DEFAULT_QUESTIONS = [
    "What is your GSTIN?",
    "Which services are you interested in?",
    "Preferred contact number",
    "Any additional notes",
]


def _load_json(path: Path) -> dict:
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text())
    except (json.JSONDecodeError, OSError):
        return {}


config = _load_json(CONFIG_PATH)
secrets = _load_json(SECRETS_PATH)

# ── Hosted deployment (Streamlit Community Cloud) ───────────────────────────
# Locally, credentials come from the gitignored secrets.json / config.json.
# Those files are deliberately NOT in the repository, so on a hosted deploy
# they are absent and every value above would be blank. Fall back to the
# platform secrets manager, which supplies the same keys under [secrets] and
# [config] tables. Local files win when present, so nothing changes in dev.
try:
    if not secrets and "secrets" in st.secrets:
        secrets = dict(st.secrets["secrets"])
    if not config and "config" in st.secrets:
        config = dict(st.secrets["config"])
except Exception:
    # No secrets.toml configured — leave the dicts empty and let the
    # per-feature error messages below explain what is missing.
    pass

SERVICE_ACCOUNT_PATH: str = config.get("google_service_account_path", "")
WORKBOOK_ID: str = secrets.get("hub_workbook_id", "")
SMTP_SERVER: str = secrets.get("smtp_server", "") or ""
SMTP_PORT: int = int(secrets.get("smtp_port", 587) or 587)
SMTP_USE_TLS: bool = bool(secrets.get("smtp_use_tls", True))
SMTP_USERNAME: str = secrets.get("smtp_username", "") or ""
SMTP_PASSWORD: str = secrets.get("smtp_password", "") or ""
SMTP_FROM_EMAIL: str = secrets.get("smtp_from_email", "") or ""
SMTP_FROM_NAME: str = secrets.get("smtp_from_name", "") or "ERPDocs"
FORM_BASE_URL: str = (secrets.get("form_base_url", "") or "http://localhost:8501").rstrip("/")
FORM_TOKEN_SECRET: str = secrets.get("form_token_secret", "") or "change-me"
ERPDOCS_BASE_URL: str = (secrets.get("erpdocs_base_url", "") or "https://staging.erpdocs.com").rstrip("/")
ERPDOCS_API_KEY: str = secrets.get("erpdocs_api_key", "") or ""


# ── Google Sheets helpers (direct REST via requests — no httplib2) ──────────
_SHEETS_BASE = "https://sheets.googleapis.com/v4/spreadsheets"


_SHEETS_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


def _has_sheets_creds() -> bool:
    """Whether Sheets credentials are obtainable by EITHER supported route.

    Callers gating on credential availability must use this rather than
    testing the file path alone, otherwise a hosted deploy — which legitimately
    has no key file — looks unconfigured.
    """
    if SERVICE_ACCOUNT_PATH and os.path.exists(SERVICE_ACCOUNT_PATH):
        return True
    try:
        return bool(st.secrets.get("gcp_service_account"))
    except Exception:
        return False


@st.cache_resource(show_spinner=False)
def _sheets_creds():
    """Load service account credentials for Google Sheets.

    Two sources, in order:
      1. A local JSON key file (dev) — `google_service_account_path`.
      2. The inlined key from the platform secrets manager (hosted) — a
         [gcp_service_account] table. A hosted container has no key file, so
         this is the only way the Sheets API can be reached there.
    """
    if SERVICE_ACCOUNT_PATH and os.path.exists(SERVICE_ACCOUNT_PATH):
        return service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_PATH, scopes=_SHEETS_SCOPES,
        )

    try:
        sa_info = st.secrets.get("gcp_service_account")
    except Exception:
        sa_info = None
    if sa_info:
        return service_account.Credentials.from_service_account_info(
            dict(sa_info), scopes=_SHEETS_SCOPES,
        )

    raise FileNotFoundError(
        "No Google service account credentials found. Locally, set "
        "'google_service_account_path' in config.json. On a hosted deploy, add a "
        "[gcp_service_account] table to the app's secrets."
    )


def _sheets_get(path: str, params: dict = None) -> dict:
    """GET request to Sheets API with auth header."""
    creds = _sheets_creds()
    if not creds.valid:
        creds.refresh(Request())
    url = f"{_SHEETS_BASE}/{WORKBOOK_ID}{path}"
    r = requests.get(url, headers={"Authorization": f"Bearer {creds.token}"}, params=params, timeout=30)
    r.raise_for_status()
    return r.json()


def _sheets_post(path: str, body: dict, params: dict = None) -> dict:
    """POST request to Sheets API with auth header."""
    creds = _sheets_creds()
    if not creds.valid:
        creds.refresh(Request())
    url = f"{_SHEETS_BASE}/{WORKBOOK_ID}{path}"
    r = requests.post(url, headers={"Authorization": f"Bearer {creds.token}",
                                     "Content-Type": "application/json"},
                       json=body, params=params, timeout=30)
    r.raise_for_status()
    return r.json()


def _sheets_put(path: str, body: dict, params: dict = None) -> dict:
    """PUT request to Sheets API with auth header."""
    creds = _sheets_creds()
    if not creds.valid:
        creds.refresh(Request())
    url = f"{_SHEETS_BASE}/{WORKBOOK_ID}{path}"
    r = requests.put(url, headers={"Authorization": f"Bearer {creds.token}",
                                    "Content-Type": "application/json"},
                      json=body, params=params, timeout=30)
    r.raise_for_status()
    return r.json()


def _ensure_tab(tab: str, headers: list[str]) -> None:
    """Create a tab with the given headers if it does not already exist."""
    meta = _sheets_get("")
    titles = [s.get("properties", {}).get("title", "") for s in meta.get("sheets", [])]
    if tab in titles:
        return
    _sheets_post(":batchUpdate", {"requests": [{"addSheet": {"properties": {"title": tab}}}]})
    _sheets_put(f"/values/'{tab}'!A1", {"values": [headers]}, params={"valueInputOption": "RAW"})


def read_tab(tab: str, retries: int = 3) -> pd.DataFrame:
    """Read a tab as a DataFrame. Returns empty DF if tab missing/empty."""
    import time
    for attempt in range(retries):
        try:
            meta = _sheets_get("")
            titles = [s.get("properties", {}).get("title", "") for s in meta.get("sheets", [])]
            if tab not in titles:
                return pd.DataFrame()
            res = _sheets_get(f"/values/'{tab}'!A1:Z100000")
            values = res.get("values", [])
            if not values:
                return pd.DataFrame()
            headers = values[0]
            rows = values[1:]
            rows = [r + [""] * (len(headers) - len(r)) for r in rows]
            df = pd.DataFrame(rows, columns=headers)
            return df.astype(str).fillna("")
        except Exception:
            if attempt < retries - 1:
                time.sleep(1)
                continue
            raise


def append_row(tab: str, row: dict, headers: list[str]) -> None:
    """Append one row (dict) to a tab, ordered by `headers`."""
    values = [[row.get(h, "") for h in headers]]
    _sheets_post(f"/values/'{tab}'!A1:append",
                 {"values": values},
                 params={"valueInputOption": "RAW", "insertDataOption": "INSERT_ROWS"})


def save_questions(questions: list[str]) -> None:
    """Persist questions to FormConfig tab (one row per question)."""
    headers = ["order", "question"]
    _ensure_tab(FORMCONFIG_TAB, headers)
    # Clear existing data rows
    creds = _sheets_creds()
    if not creds.valid:
        creds.refresh(Request())
    url = f"{_SHEETS_BASE}/{WORKBOOK_ID}/values/'{FORMCONFIG_TAB}'!A2:Z100000:clear"
    requests.post(url, headers={"Authorization": f"Bearer {creds.token}"}, timeout=30)
    values = [[str(i + 1), q] for i, q in enumerate(questions) if q.strip()]
    if values:
        _sheets_put(f"/values/'{FORMCONFIG_TAB}'!A2", {"values": values},
                    params={"valueInputOption": "RAW"})


def load_questions() -> list[str]:
    try:
        df = read_tab(FORMCONFIG_TAB)
    except Exception:
        return list(DEFAULT_QUESTIONS)
    if df.empty or "question" not in df.columns:
        return list(DEFAULT_QUESTIONS)
    qs = df["question"].tolist()
    return [q for q in qs if q.strip()] or list(DEFAULT_QUESTIONS)


# ── ERPDocs API client ──────────────────────────────────────────────────────
class ERPDocs:
    """Thin REST client for ERPDocs v1 API. Mirrors the official docs."""

    def __init__(self, base_url: str, api_key: str):
        self.base = base_url.rstrip("/")
        self.key = api_key

    def _headers(self):
        return {"Authorization": f"Bearer {self.key}"}

    def _req(self, method: str, path: str, **kw):
        # Merge caller-supplied headers (e.g. Idempotency-Key) with auth headers
        extra_headers = kw.pop("headers", {})
        # Don't force Content-Type for multipart (requests sets boundary)
        has_files = "files" in kw
        merged = {**self._headers(), **extra_headers}
        if not has_files:
            merged.setdefault("Content-Type", "application/json")
        try:
            r = requests.request(
                method, f"{self.base}/api/v1{path}",
                headers=merged, timeout=30, **kw,
            )
        except requests.exceptions.ConnectionError:
            return None, {"error": {"code": "unreachable", "message": f"Cannot reach {self.base}"}}
        except requests.exceptions.Timeout:
            return None, {"error": {"code": "timeout", "message": "Request timed out"}}
        if r.status_code in (200, 201, 202):
            return r.json(), None
        try:
            return None, r.json()
        except Exception:
            return None, {"error": {"code": str(r.status_code), "message": r.text[:400]}}

    @staticmethod
    def rows(payload):
        if payload is None:
            return []
        if isinstance(payload, list):
            return payload
        if isinstance(payload, dict):
            for key in ("data", "templates", "users", "signpacks", "items"):
                val = payload.get(key)
                if isinstance(val, list):
                    return val
        return []

    def ping(self):
        return self._req("GET", "/ping")

    def templates(self):
        return self._req("GET", "/templates")

    def template(self, tid):
        return self._req("GET", f"/templates/{tid}")

    def users(self):
        return self._req("GET", "/users")

    def send_from_template(self, tid, payload):
        return self._req("POST", f"/templates/{tid}/send", json=payload)

    def bulk_send(self, payload, idempotency_key=None):
        headers = {}
        if idempotency_key:
            headers["Idempotency-Key"] = idempotency_key
        kw = {"json": payload}
        if headers:
            kw["headers"] = headers
        return self._req("POST", "/bulk-send", **kw)

    def bulk_status(self, batch_id):
        return self._req("GET", f"/bulk-send/{batch_id}")

    def bulk_rows(self, batch_id, limit=100, offset=0, status=None):
        params = {"limit": limit, "offset": offset}
        if status:
            params["status"] = status
        return self._req("GET", f"/bulk-send/{batch_id}/rows", params=params)

    def signpack(self, sid):
        return self._req("GET", f"/signpacks/{sid}")

    def signpacks(self, limit=50, offset=0, status=None):
        params = {"limit": limit, "offset": offset}
        if status:
            params["status"] = status
        return self._req("GET", "/signpacks", params=params)

    def download(self, sid):
        return self._req("GET", f"/signpacks/{sid}/download")

    def remind(self, sid, message="", recipient_ids=None):
        body = {}
        if message:
            body["message"] = message
        if recipient_ids:
            body["recipient_ids"] = recipient_ids
        return self._req("POST", f"/signpacks/{sid}/remind", json=body)

    def embed_token(self, sid, rid, payload):
        return self._req("POST", f"/signpacks/{sid}/recipients/{rid}/embed-token", json=payload)

    def void(self, sid, reason=""):
        body = {}
        if reason:
            body["reason"] = reason
        return self._req("POST", f"/signpacks/{sid}/void", json=body)

    def brands(self):
        return self._req("GET", "/brands")

    def brand(self, bid):
        return self._req("GET", f"/brands/{bid}")

    def create_brand(self, data: dict, logo_file=None):
        """Create a brand profile. Multipart form-data when logo is present."""
        if logo_file is not None:
            files = {"logo": (logo_file.name, logo_file, logo_file.type)}
            form = {k: str(v) for k, v in data.items() if v is not None}
            return self._req("POST", "/brands", data=form, files=files)
        return self._req("POST", "/brands", data={k: str(v) for k, v in data.items() if v is not None})

    def update_brand(self, bid, data: dict, logo_file=None):
        if logo_file is not None:
            files = {"logo": (logo_file.name, logo_file, logo_file.type)}
            form = {k: str(v) for k, v in data.items() if v is not None}
            return self._req("PATCH", f"/brands/{bid}", data=form, files=files)
        return self._req("PATCH", f"/brands/{bid}", data={k: str(v) for k, v in data.items() if v is not None})

    def delete_brand(self, bid):
        return self._req("DELETE", f"/brands/{bid}")

    def create_signpack_raw(self, file_bytes, filename, fields: dict):
        """POST /signpacks — raw-PDF send with anchor tags. Multipart form-data."""
        files = {"file": (filename, file_bytes, "application/pdf")}
        return self._req("POST", "/signpacks", data=fields, files=files)


# ── Tokens & form links ─────────────────────────────────────────────────────
def client_token(email: str) -> str:
    """Deterministic, non-trivially-guessable per-client token (no storage)."""
    raw = f"{email.lower().strip()}:{FORM_TOKEN_SECRET}".encode()
    return hashlib.sha256(raw).hexdigest()[:20]


def form_link(email: str) -> str:
    return f"{FORM_BASE_URL}/?view=form&token={client_token(email)}"


def resolve_token(token: str, clients: pd.DataFrame) -> Optional[pd.Series]:
    """Find the client row whose email hashes to this token."""
    email_col = _email_col(clients)
    if email_col is None:
        return None
    for _, row in clients.iterrows():
        if client_token(row[email_col]) == token:
            return row
    return None


# ── Column helpers (tolerant of header naming) ──────────────────────────────
def _col(df: pd.DataFrame, *candidates: str) -> Optional[str]:
    if df.empty:
        return None
    lower_map = {c.lower().strip(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in lower_map:
            return lower_map[cand.lower()]
    return None


def _email_col(df: pd.DataFrame) -> Optional[str]:
    return _col(df, "Email Id", "Email", "Email ID", "email id")


def _name_col(df: pd.DataFrame) -> Optional[str]:
    return _col(df, "Name", "Client Name")


def _company_col(df: pd.DataFrame) -> Optional[str]:
    return _col(df, "Company Name", "Company")


# ── Email sending ───────────────────────────────────────────────────────────
def send_email(to: str, subject: str, html_body: str, cc: str = "") -> None:
    if not SMTP_SERVER or not SMTP_USERNAME or not SMTP_PASSWORD:
        raise RuntimeError(
            "SMTP not configured. Set smtp_server, smtp_username, smtp_password "
            "in secrets.json."
        )
    from_header = f"{SMTP_FROM_NAME} <{SMTP_FROM_EMAIL or SMTP_USERNAME}>" if SMTP_FROM_NAME else (SMTP_FROM_EMAIL or SMTP_USERNAME)
    msg = MIMEMultipart("alternative")
    msg["From"] = from_header
    msg["To"] = to
    if cc:
        msg["Cc"] = cc
    msg["Subject"] = subject
    msg.attach(MIMEText(html_body, "html"))

    ctx = ssl.create_default_context(cafile=_CA_BUNDLE) if _CA_BUNDLE else ssl.create_default_context()
    if SMTP_USE_TLS:
        # STARTTLS: connect plain, upgrade, then auth
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=30) as s:
            s.ehlo()
            s.starttls(context=ctx)
            s.ehlo()
            s.login(SMTP_USERNAME, SMTP_PASSWORD)
            recipients = [x.strip() for x in to.split(",") if x.strip()]
            if cc:
                recipients += [x.strip() for x in cc.split(",") if x.strip()]
            s.sendmail(SMTP_FROM_EMAIL or SMTP_USERNAME, recipients, msg.as_string())
    else:
        # Implicit TLS (SMTPS) on the given port
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, context=ctx, timeout=30) as s:
            s.login(SMTP_USERNAME, SMTP_PASSWORD)
            recipients = [x.strip() for x in to.split(",") if x.strip()]
            if cc:
                recipients += [x.strip() for x in cc.split(",") if x.strip()]
            s.sendmail(SMTP_FROM_EMAIL or SMTP_USERNAME, recipients, msg.as_string())


def render_body(template: str, client: pd.Series) -> str:
    """Replace {{placeholders}} with client values. Unknown keys stay as-is."""
    name_c = _name_col(pd.DataFrame([client])) if not client.empty else None
    company_c = _company_col(pd.DataFrame([client])) if not client.empty else None
    email_c = _email_col(pd.DataFrame([client])) if not client.empty else None
    out = template
    repl = {
        "Name": client.get(name_c, "") if name_c else "",
        "Company": client.get(company_c, "") if company_c else "",
        "Email": client.get(email_c, "") if email_c else "",
        "FormLink": form_link(client[email_c]) if email_c else "",
    }
    for k, v in repl.items():
        out = out.replace("{{" + k + "}}", str(v))
    return out


# ── Session state ───────────────────────────────────────────────────────────
def _init_state():
    st.session_state.setdefault("last_send_log", [])
    st.session_state.setdefault("esign_active_token", None) # {recipient_id, sign_url, expires_at}
    st.session_state.setdefault("cp_form_submitted", False) # client portal: form done


_init_state()


# ── Signature settings persistence ──────────────────────────────────────────
SETTINGS_FILE = Path(__file__).parent / "signature_settings.json"

DEFAULT_SETTINGS = {
    # Embedded signing
    "embed_origin": "",
    "iframe_height": 820,
    "redirect_url": "",
    "ttl_seconds": 300,
    "suppress_completion_email": False,
    "signer_auth_method": "portal_link",
    # Send defaults
    "default_brand_id": "",
    "default_expires_days": 30,
    "default_signing_order": "parallel",
    "default_send_immediately": True,
    # Connection (overrides secrets.json when set in panel)
    "api_base_url": "",
    "api_key": "",
}


def load_settings() -> dict:
    """Load signature settings from disk, merged with defaults."""
    s = dict(DEFAULT_SETTINGS)
    if SETTINGS_FILE.exists():
        try:
            saved = json.loads(SETTINGS_FILE.read_text())
            s.update(saved)
        except Exception:
            pass
    return s


def save_settings(settings: dict) -> None:
    """Persist signature settings to disk."""
    SETTINGS_FILE.write_text(json.dumps(settings, indent=2))


# ── Apply signature settings overrides BEFORE any view runs ─────────────────
# signature_settings.json overrides secrets.json for API base URL and key.
# This must happen before the view router so that ?view=form (signer's page)
# uses the same credentials as the sender — otherwise the signer sees
# "Recipient not found" because the envelope lives on a different instance.
_sig_settings = load_settings()
if _sig_settings.get("api_base_url"):
    ERPDOCS_BASE_URL = _sig_settings["api_base_url"].rstrip("/")
if _sig_settings.get("api_key"):
    ERPDOCS_API_KEY = _sig_settings["api_key"]


# ── View router ─────────────────────────────────────────────────────────────
view = st.query_params.get("view", "sender")


# ═══════════════════════════════════════════════════════════════════════════
# SIGNATURE SETTINGS PANEL
# URL: ?view=settings
# Configuration only — 3 tabs: Connection, Send & Embed, Brands.
# Operational tools (raw-PDF send, void/remind/download) live at ?view=tools.
# ═══════════════════════════════════════════════════════════════════════════
if view == "settings":
    st.title("Signature Settings")
    st.caption("Configure ERPDocs e-sign API options. "
               "Settings persist to `signature_settings.json` and apply to all sends. "
               "Operational tools (send raw PDF, void, remind, download) are at "
               "[?view=tools](?view=tools).")

    settings = load_settings()

    # Resolve active API credentials (panel overrides > secrets.json)
    active_base = (settings.get("api_base_url") or ERPDOCS_BASE_URL).rstrip("/")
    active_key = settings.get("api_key") or ERPDOCS_API_KEY
    api = ERPDocs(active_base, active_key) if active_key else None

    # ── Tab layout — 3 clean categories ─────────────────────────────────────
    tab_conn, tab_send, tab_brands = st.tabs([
        "Connection",
        "Send & Embed",
        "Brands",
    ])

    # ── Tab 1: Connection ───────────────────────────────────────────────────
    with tab_conn:
        st.markdown("#### API Connection")
        st.caption("Overrides `secrets.json` when set. Leave blank to use secrets.json values. "
                   "Changes save automatically.")

        c1, c2 = st.columns(2)
        with c1:
            new_base = st.text_input(
                "ERPDocs Base URL",
                value=active_base,
                help="https://staging.erpdocs.com for staging, https://www.erpdocs.com for production.",
                key="set_base_url",
            )
        with c2:
            new_key = st.text_input(
                "API Key",
                value=active_key,
                type="password",
                help="LIVE key. Required scopes: template.read, signpack.read, signpack.write, "
                     "signpack.embed, brand.read, brand.write.",
                key="set_api_key",
            )

        # Auto-save connection
        settings["api_base_url"] = new_base.rstrip("/")
        settings["api_key"] = new_key
        save_settings(settings)

        if st.button("Test Connection", key="set_test_conn"):
            test_api = ERPDocs(new_base.rstrip("/"), new_key)
            data, err = test_api.ping()
            if err:
                st.error(err.get("error", {}).get("message", "Failed"))
            else:
                st.success(f"{data.get('org_name')} · {data.get('environment')}")
                scopes = data.get("scopes", [])
                st.write("**Granted scopes:**", ", ".join(scopes) if scopes else "*(none)*")
                needed = [
                    "template.read", "signpack.read", "signpack.write",
                    "signpack.embed", "brand.read", "brand.write",
                ]
                missing = [s for s in needed if s not in scopes]
                if missing:
                    st.warning(f"Missing scopes: {', '.join(missing)}")
                else:
                    st.info("All required scopes present.")
                if data.get("environment") == "sandbox":
                    st.warning("Sandbox key — writes return `sandbox_write_blocked`. "
                               "Use a LIVE key to send envelopes.")

    # ── Tab 2: Send & Embed ─────────────────────────────────────────────────
    with tab_send:
        # ── Section: Default Send Options ───────────────────────────────────
        st.markdown("#### Default Send Options")
        st.caption("Applied to all template sends and raw-PDF sends from this app.")

        # Brand selector
        brand_opts = ["(none — use org default)"]
        brand_map = {}
        if api:
            bdata, berr = api.brands()
            if not berr:
                for b in ERPDocs.rows(bdata):
                    label = f"{b.get('name', 'Unnamed')} (id {b.get('id')})"
                    brand_opts.append(label)
                    brand_map[label] = b.get("id")

        current_brand_idx = 0
        saved_bid = settings.get("default_brand_id")
        if saved_bid:
            for i, opt in enumerate(brand_opts):
                if str(brand_map.get(opt, "")) == str(saved_bid):
                    current_brand_idx = i
                    break

        chosen_brand = st.selectbox(
            "Default Brand Profile",
            options=brand_opts,
            index=current_brand_idx,
            help="Overrides the ERPDocs sender name, reply-to, logo and email header/footer. "
                 "Manage brands in the Brands tab.",
            key="set_default_brand",
        )
        settings["default_brand_id"] = brand_map.get(chosen_brand, "")

        c1, c2 = st.columns(2)
        with c1:
            settings["default_expires_days"] = st.number_input(
                "Default Expiry (days)",
                min_value=1, max_value=365,
                value=int(settings.get("default_expires_days", 30)),
                key="set_default_expiry",
            )
        with c2:
            settings["default_signing_order"] = st.selectbox(
                "Default Signing Order",
                options=["parallel", "sequential"],
                index=0 if settings.get("default_signing_order", "parallel") == "parallel" else 1,
                help="parallel = all sign at once. sequential = signers sign one after another. "
                     "Applies to raw-PDF sends only (template sends use the template's order).",
                key="set_default_order",
            )

        settings["default_send_immediately"] = st.checkbox(
            "Send immediately (uncheck to create Draft)",
            value=settings.get("default_send_immediately", True),
            help="When unchecked, envelopes are created as Drafts and must be sent manually "
                 "from the ERPDocs UI or via a separate call.",
            key="set_default_send",
        )

        st.divider()

        # ── Section: Embedded Signing ───────────────────────────────────────
        st.markdown("#### Embedded Signing")
        st.caption("Controls how the signing iframe is embedded for clients.")

        c1, c2 = st.columns(2)
        with c1:
            settings["embed_origin"] = st.text_input(
                "Embed Origin",
                value=settings.get("embed_origin", ""),
                help="Exact origin that will frame the signing page. Must be registered on the "
                     "API key in Admin → API & Webhooks → Embed Origins. "
                     "e.g. https://app.yourcompany.com (no path, no wildcard).",
                key="set_embed_origin",
            )
        with c2:
            settings["iframe_height"] = st.number_input(
                "Iframe Height (px)",
                min_value=400, max_value=1200,
                value=int(settings.get("iframe_height", 820)),
                step=20,
                help="ERPDocs recommends at least 540px so the action bar stays visible.",
                key="set_iframe_h",
            )

        settings["redirect_url"] = st.text_input(
            "Redirect URL (optional)",
            value=settings.get("redirect_url", ""),
            help="Where the signer goes after completing or declining. Must be on the same "
                 "origin as embed_origin. ERPDocs appends ?event=signing_complete or ?event=decline. "
                 "Leave blank: postMessage listener handles completion (works on Streamlit Cloud). "
                 "Set explicitly to test the redirect_url API: locally use "
                 "http://localhost:8501/?view=complete; on Streamlit Cloud this will redirect-loop.",
            key="set_redirect_url",
        )

        c3, c4 = st.columns(2)
        with c3:
            settings["ttl_seconds"] = st.number_input(
                "Token TTL (seconds)",
                min_value=30, max_value=900,
                value=int(settings.get("ttl_seconds", 300)),
                help="Embed token lifetime. Default 300 (5 min). Mint when about to display, "
                     "not when creating the envelope.",
                key="set_ttl",
            )
        with c4:
            settings["signer_auth_method"] = st.text_input(
                "Signer Auth Method",
                value=settings.get("signer_auth_method", "portal_link"),
                help="Recorded on the Certificate of Completion as the evidence of attribution. "
                     "e.g. 'password', 'portal_link', 'sso'. The identifier is the signer's email.",
                key="set_auth_method",
            )

        _send_completion = st.checkbox(
            "Send completion email to signer (signed-copy + completion notification)",
            value=not settings.get("suppress_completion_email", False),
            help="When checked, signers receive the completion notification and the final "
                 "signed-copy email from ERPDocs. Uncheck to suppress both — use this when "
                 "your application handles delivery of the signed document. "
                 "Applies to BOTH send-time and embed-token requests.",
            key="set_send_ce",
        )
        settings["suppress_completion_email"] = not _send_completion

        # Auto-save
        save_settings(settings)

        with st.expander("Email matrix for embedded recipients", expanded=False):
            st.markdown("""
| Email type | Sent to embedded signer? |
|---|---|
| Invitation / "please sign" | **No** — suppressed |
| Reminder (automatic and manual) | **No** — suppressed |
| Your-turn notification | **No** — suppressed |
| CC notification | **No** — suppressed |
| Completion notification | **Yes** (unless suppressed above) |
| Final signed copy | **Yes** (unless suppressed above) |
""")

    # ── Tab 3: Brands ───────────────────────────────────────────────────────
    with tab_brands:
        st.markdown("#### Brand Profiles")
        st.caption("Manage `OrgBrandProfile` records. Each profile overrides the default ERPDocs "
                   "branding on emails — sender name, reply-to, logo, colors, header/footer HTML. "
                   "Requires `brand.read` and `brand.write` scopes.")

        if not api:
            st.warning("API key required. Set it in the Connection tab.")
        else:
            # List existing brands
            bdata, berr = api.brands()
            if berr:
                e = berr.get("error", {})
                st.error(f"`{e.get('code')}` — {e.get('message')}")
                if e.get("code") == "insufficient_scope":
                    st.info("Add `brand.read` scope to your API key in ERPDocs Admin → API & Webhooks.")
            else:
                brands = ERPDocs.rows(bdata)
                if brands:
                    st.write(f"**{len(brands)} brand profile(s):**")
                    for b in brands:
                        with st.expander(f"{b.get('name', 'Unnamed')} (id {b.get('id')})"):
                            for k in ("name", "primary_color", "email_sender_name",
                                       "reply_to_email", "logo_path"):
                                v = b.get(k)
                                if v:
                                    st.write(f"**{k}:** `{v}`")

                            # Delete button
                            bid = b.get("id")
                            if st.button(f"Delete brand #{bid}", key=f"del_brand_{bid}"):
                                _, derr = api.delete_brand(bid)
                                if derr:
                                    st.error(derr.get("error", {}).get("message", "Failed"))
                                else:
                                    st.success(f"Brand #{bid} deleted.")
                                    st.rerun()
                else:
                    st.info("No brand profiles yet. Create one below.")

            st.divider()

            # Create / edit brand form
            st.markdown("##### Create New Brand")
            with st.form("create_brand_form"):
                cb_name = st.text_input("Name *", key="cb_name")
                cb_color = st.text_input(
                    "Primary Color",
                    value="#0891b2",
                    help="Hex color code, e.g. #0891b2 (teal).",
                    key="cb_color",
                )
                cb_sender = st.text_input(
                    "Email Sender Name",
                    help="Display name on outbound emails. Falls back to 'ERPDocs'.",
                    key="cb_sender",
                )
                cb_reply = st.text_input(
                    "Reply-To Email",
                    help="Reply-to address on outbound emails.",
                    key="cb_reply",
                )
                cb_header = st.text_area(
                    "Email Header HTML (optional)",
                    height=80,
                    key="cb_header",
                )
                cb_footer = st.text_area(
                    "Email Footer HTML (optional)",
                    height=80,
                    key="cb_footer",
                )
                cb_logo = st.file_uploader(
                    "Logo (PNG/JPG, optional)",
                    type=["png", "jpg", "jpeg"],
                    key="cb_logo",
                )

                if st.form_submit_button("Create Brand", type="primary"):
                    if not cb_name.strip():
                        st.error("Name is required.")
                    else:
                        data = {
                            "name": cb_name.strip(),
                            "primary_color": cb_color.strip() or None,
                            "email_sender_name": cb_sender.strip() or None,
                            "reply_to_email": cb_reply.strip() or None,
                            "email_header_html": cb_header.strip() or None,
                            "email_footer_html": cb_footer.strip() or None,
                        }
                        _, cerr = api.create_brand(data, logo_file=cb_logo)
                        if cerr:
                            e = cerr.get("error", {})
                            st.error(f"`{e.get('code')}` — {e.get('message')}")
                        else:
                            st.success(f"Brand '{cb_name}' created.")
                            st.rerun()

    st.divider()
    st.markdown(
        f'<div class="ce-note">Settings are stored in '
        f'<code>{SETTINGS_FILE.name}</code>. '
        f'<a href="?view=tools">Operational tools →</a> · '
        f'<a href="?view=sender">← Back to sender</a></div>',
        unsafe_allow_html=True,
    )
    st.stop()


# ═══════════════════════════════════════════════════════════════════════════
# ERPDOCS TOOLS PANEL
# URL: ?view=tools
# Operational tools: raw-PDF anchor send, void, remind, download, envelopes.
# Separated from settings to keep configuration clean.
# ═══════════════════════════════════════════════════════════════════════════
if view == "tools":
    st.title("E-Sign Tools")
    st.caption("Send raw PDFs with anchor tags, void/remind/download envelopes. "
               "Configuration is at [?view=settings](?view=settings).")

    settings = load_settings()
    active_base = (settings.get("api_base_url") or ERPDOCS_BASE_URL).rstrip("/")
    active_key = settings.get("api_key") or ERPDOCS_API_KEY
    api = ERPDocs(active_base, active_key) if active_key else None

    tab_raw, tab_envelopes = st.tabs([
        "Raw-PDF Send",
        "Envelope Actions",
    ])

    # ── Tab 1: Raw-PDF / Anchors ────────────────────────────────────────────
    with tab_raw:
        st.markdown("#### Raw-PDF Send with Anchor Tags")
        st.caption("Upload a PDF, DOCX or image containing anchor tags. The API detects the tags, "
                   "redacts them, places real fields at the same positions, and sends the envelope. "
                   "No template required. Requires `signpack.write` scope.")

        # Anchor tag cheat sheet
        with st.expander("Anchor Tag Reference", expanded=False):
            st.markdown("""
| Tag | Field Type | Example |
|-----|------------|---------|
| `{{sig:Role}}` or `{{signature:Role}}` | Signature | `{{sig:Signer 1}}` |
| `{{name:Role}}` | Signer name | `{{name:Signer 1}}` |
| `{{date:Role}}` | Date signed | `{{date:Signer 1}}` |
| `{{text:Role:Label}}` | Free text | `{{text:Signer 1:Address}}` |
| `{{checkbox:Role:Label}}` | Checkbox | `{{checkbox:Signer 1:I agree}}` |
| `{{email:Role}}` | Email address | `{{email:Signer 1}}` |
| `{{number:Role:Label}}` | Number | `{{number:Signer 1:Amount}}` |
| `{{attachment:Role:Label}}` | File attachment | `{{attachment:Signer 1:Upload ID}}` |
| `{{location:Role}}` | Location | `{{location:Signer 1}}` |
| `{{tag:Value}}` or `{{tag:Label:Value}}` | Read-only tag | `{{tag:Invoice: 12345}}` |

The role in the anchor (e.g. `Signer 1`) must match a key in the recipients JSON below.
A recipient with at least one actionable field becomes a **Signer**; a recipient with
none becomes **CC**.
""")

        if not api:
            st.warning("API key required. Set it in [Settings → Connection](?view=settings).")
        else:
            uploaded = st.file_uploader(
                "Upload document (PDF, DOCX, or image)",
                type=["pdf", "docx", "doc", "png", "jpg", "jpeg", "bmp", "tiff", "webp"],
                key="raw_upload",
            )

            raw_title = st.text_input(
                "Title (optional)",
                help="Defaults to the filename.",
                key="raw_title",
            )

            # Sender
            users_data, _ = api.users()
            user_emails = [u.get("email") for u in ERPDocs.rows(users_data) if u.get("email")]
            if user_emails:
                raw_sender = st.selectbox("Sender (org member)", user_emails, key="raw_sender")
            else:
                raw_sender = st.text_input("Sender email", key="raw_sender_manual")

            # Recipients JSON
            st.caption("**Recipients** — JSON object. Each key is a role name matching an anchor tag.")
            default_recips = '{"Signer 1": {"name": "Ravi", "email": "ravi@example.com", "embedded": true}}'
            raw_recips = st.text_area(
                "Recipients JSON",
                value=default_recips,
                height=120,
                key="raw_recips",
            )

            c1, c2, c3 = st.columns(3)
            with c1:
                raw_order = st.selectbox(
                    "Signing Order",
                    options=["parallel", "sequential"],
                    key="raw_order",
                )
            with c2:
                raw_expiry = st.number_input(
                    "Expires in (days)",
                    min_value=1, max_value=365, value=30,
                    key="raw_expiry",
                )
            with c3:
                raw_send = st.checkbox("Send immediately", value=True, key="raw_send")

            # Brand selector for raw-PDF
            raw_brand_opts = ["(none)"]
            raw_brand_map = {}
            bdata, _ = api.brands() if api else (None, {"error": {}})
            if bdata:
                for b in ERPDocs.rows(bdata):
                    label = f"{b.get('name', 'Unnamed')} (id {b.get('id')})"
                    raw_brand_opts.append(label)
                    raw_brand_map[label] = b.get("id")
            raw_chosen_brand = st.selectbox("Brand profile", options=raw_brand_opts, key="raw_brand")
            raw_brand_id = raw_brand_map.get(raw_chosen_brand, "")

            if st.button("Send Raw PDF", type="primary", key="raw_send_btn"):
                if not uploaded:
                    st.error("Upload a document first.")
                elif not raw_sender:
                    st.error("Sender email is required.")
                else:
                    try:
                        recips = json.loads(raw_recips)
                    except json.JSONDecodeError as e:
                        st.error(f"Invalid recipients JSON: {e}")
                        st.stop()

                    fields = {
                        "sender_email": raw_sender,
                        "recipients": json.dumps(recips),
                        "signing_order": raw_order,
                        "expires_in_days": str(int(raw_expiry)),
                        "send": str(raw_send).lower(),
                    }
                    if raw_title.strip():
                        fields["title"] = raw_title.strip()
                    if raw_brand_id:
                        fields["brand_id"] = str(raw_brand_id)

                    with st.spinner("Processing anchors and sending…"):
                        data, rerr = api.create_signpack_raw(
                            uploaded.read(), uploaded.name, fields,
                        )
                    if rerr:
                        e = rerr.get("error", {})
                        st.error(f"`{e.get('code')}` — {e.get('message')}")
                    else:
                        st.success(f"Envelope #{data.get('signpack_id')} created — "
                                   f"status: {data.get('status')}")
                        st.write(f"**Title:** {data.get('title')}")
                        st.write(f"**Recipients:** {data.get('recipient_count')}")
                        st.write(f"**Fields detected:** {data.get('field_count')}")
                        unrec = data.get("anchor_unrecognised") or []
                        if unrec:
                            st.warning(f"Unrecognised anchors: {unrec}")
                        st.write("**Recipients:**")
                        for r in data.get("recipients", []):
                            st.write(f"- {r.get('role_key')}: {r.get('name')} <{r.get('email')}> "
                                     f"({r.get('role')}, embedded={r.get('embedded')})")

    # ── Tab 2: Envelope Actions ─────────────────────────────────────────────
    with tab_envelopes:
        st.markdown("#### Envelope Actions")
        st.caption("Void, remind, and download signed envelopes. Requires `signpack.write` "
                   "for void/remind and `signpack.read` for download.")

        if not api:
            st.warning("API key required. Set it in [Settings → Connection](?view=settings).")
        else:
            # Void section
            st.markdown("##### Void Envelope")
            st.caption("Permanently cancel an in-flight envelope. Recipients can no longer access "
                       "the document. Permitted statuses: Sent, In Progress, Correction Requested, Recalled.")
            with st.form("void_form"):
                void_sid = st.text_input("SignPack ID", key="void_sid")
                void_reason = st.text_input("Reason (optional)", key="void_reason")
                if st.form_submit_button("Void Envelope", type="primary"):
                    if not void_sid.strip():
                        st.error("SignPack ID is required.")
                    else:
                        _, verr = api.void(void_sid.strip(), void_reason.strip())
                        if verr:
                            e = verr.get("error", {})
                            st.error(f"`{e.get('code')}` — {e.get('message')}")
                        else:
                            st.success(f"Envelope #{void_sid} voided — status: Cancelled")

            st.divider()

            # Remind section
            st.markdown("##### Send Reminder")
            st.caption("Queue a reminder email to pending signers. Only applies to envelopes "
                       "in Sent, In Progress, or Correction Requested status. "
                       "Embedded signers are not emailed — they are driven by your application.")
            with st.form("remind_form"):
                rem_sid = st.text_input("SignPack ID", key="rem_sid")
                rem_msg = st.text_area("Message (optional)", height=80, key="rem_msg")
                if st.form_submit_button("Send Reminder"):
                    if not rem_sid.strip():
                        st.error("SignPack ID is required.")
                    else:
                        _, rerr = api.remind(rem_sid.strip(), rem_msg.strip())
                        if rerr:
                            e = rerr.get("error", {})
                            st.error(f"`{e.get('code')}` — {e.get('message')}")
                        else:
                            st.success(f"Reminder queued for envelope #{rem_sid}")

            st.divider()

            # Download section
            st.markdown("##### Download Signed PDF")
            st.caption("Fetch the completed, sealed PDF for a completed envelope.")
            with st.form("dl_form"):
                dl_sid = st.text_input("SignPack ID", key="dl_sid")
                if st.form_submit_button("Download"):
                    if not dl_sid.strip():
                        st.error("SignPack ID is required.")
                    else:
                        dl, dlerr = api.download(dl_sid.strip())
                        if dlerr:
                            e = dlerr.get("error", {})
                            st.error(f"`{e.get('code')}` — {e.get('message')}")
                        else:
                            docs = dl.get("documents", [])
                            if docs:
                                for doc in docs:
                                    st.write(f"**[{doc.get('filename')}]({doc.get('url')})**")
                                    st.caption(f"SHA-256: `{doc.get('sha256', '')[:30]}…`")
                            else:
                                st.info("No documents available yet.")

            st.divider()

            # Recent envelopes with void buttons
            st.markdown("##### Recent Envelopes")
            if st.button("Fetch Last 20 Envelopes", key="set_fetch_envs"):
                sp_data, sp_err = api.signpacks(limit=20)
                if sp_err:
                    st.error(sp_err.get("error", {}).get("message", "Failed"))
                else:
                    envs = ERPDocs.rows(sp_data)
                    if envs:
                        summary = []
                        for env in envs:
                            recipients = env.get("recipients", [])
                            signer_status = "—"
                            for rec in recipients:
                                if rec.get("role", "").lower() in ("signer", "reviewer"):
                                    signer_status = rec.get("status", "—")
                                    break
                            summary.append({
                                "ID": env.get("id"),
                                "Title": env.get("title", ""),
                                "Status": env.get("status", "—"),
                                "Signer": signer_status,
                                "Created": env.get("created_at", "—"),
                                "Brand": env.get("brand_id", "—"),
                            })
                        st.dataframe(pd.DataFrame(summary), width='stretch', hide_index=True)

                        # Quick void for in-flight envelopes
                        voidable = [e for e in envs if e.get("status") in
                                    ("Sent", "In Progress", "Correction Requested", "Recalled")]
                        if voidable:
                            st.caption(f"{len(voidable)} voidable envelope(s).")
                            for env in voidable[:5]:
                                eid = env.get("id")
                                with st.expander(f"Void #{eid} — {env.get('title', '')}"):
                                    vreason = st.text_input(
                                        "Reason", key=f"qvoid_r_{eid}",
                                    )
                                    if st.button(f"Void #{eid}", key=f"qvoid_b_{eid}"):
                                        _, verr = api.void(eid, vreason)
                                        if verr:
                                            st.error(verr.get("error", {}).get("message", "Failed"))
                                        else:
                                            st.success(f"Envelope #{eid} voided.")
                                            st.rerun()
                    else:
                        st.info("No envelopes found.")

    st.divider()
    st.markdown(
        f'<div class="ce-note">'
        f'<a href="?view=settings">← Back to Settings</a> · '
        f'<a href="?view=sender">← Back to sender</a></div>',
        unsafe_allow_html=True,
    )
    st.stop()


# ═══════════════════════════════════════════════════════════════════════════
# CLIENT PORTAL VIEW (one link: form → sign, unified)
# URL: ?view=form&token=...&sid=...&rid=...
# ═══════════════════════════════════════════════════════════════════════════
if view == "form":
    token = st.query_params.get("token", "")
    sid = st.query_params.get("sid", "")
    rid = st.query_params.get("rid", "")

    if not token:
        st.error("Missing client token in the link.")
        st.stop()

    # ── Resolve client identity ─────────────────────────────────────────────
    try:
        clients = read_tab(CLIENT_TAB)
    except Exception as e:
        st.error(f"Could not load client data: {e}")
        st.stop()

    if clients.empty:
        st.warning("No client data available.")
        st.stop()

    client = resolve_token(token, clients)
    if client is None:
        st.error("Invalid or expired link. Please request a new email.")
        st.stop()

    name_c = _name_col(clients)
    company_c = _company_col(clients)
    email_c = _email_col(clients)
    client_name = str(client.get(name_c, "")) if name_c else ""
    client_company = str(client.get(company_c, "")) if company_c else ""
    client_email = str(client.get(email_c, "")) if email_c else ""

    # ── Check if there's a signing envelope attached ────────────────────────
    has_signing = bool(sid and rid and ERPDOCS_API_KEY)
    envelope_info = None
    recipient_info = None
    already_signed = False

    if has_signing:
        api = ERPDocs(ERPDOCS_BASE_URL, ERPDOCS_API_KEY)
        sp, sp_err = api.signpack(sid)
        if not sp_err:
            envelope_info = sp
            for rec in sp.get("recipients", []):
                if str(rec.get("id")) == str(rid):
                    recipient_info = rec
                    if rec.get("status") == "Signed":
                        already_signed = True
                    break

    # ── Determine step state ────────────────────────────────────────────────
    form_done = st.session_state.cp_form_submitted
    # If already signed, both steps are done
    if already_signed:
        form_done = True

    # ── Step indicators ─────────────────────────────────────────────────────
    step1_cls = "done" if form_done else "active"
    step1_lbl = "" if form_done else ""
    step2_cls = "done" if already_signed else ("active" if form_done and has_signing else "idle")
    step2_lbl_cls = "" if (form_done and has_signing or already_signed) else "idle"
    line_cls = "done" if form_done else ""

    st.markdown(f"""
    <div class="cp-wrap">
      <div class="cp-hero">
        <h1>{'Welcome back, ' + client_name if already_signed else 'Hello ' + client_name}</h1>
        <p>{('From ' + client_company) if client_company else ''}{(' · ' + client_email) if client_email else ''}</p>
      </div>
      <div class="cp-card">
        <div class="cp-step-row">
          <div class="cp-step">
            <div class="cp-step-num {step1_cls}">{('&#10003;') if form_done else '1'}</div>
            <div class="cp-step-label {'idle' if False else ''}">Your Details</div>
          </div>
          <div class="cp-step-line {line_cls}"></div>
          <div class="cp-step">
            <div class="cp-step-num {step2_cls}">{('&#10003;') if already_signed else '2'}</div>
            <div class="cp-step-label {step2_lbl_cls}">Sign Document</div>
          </div>
        </div>
    """, unsafe_allow_html=True)

    # ── Already signed state ────────────────────────────────────────────────
    if already_signed:
        st.markdown("""
        <div class="cp-success-banner">
          <div class="icon">&#10003;</div>
          <div class="text"><b>All done!</b> You've completed the form and signed the document.
          We'll be in touch shortly.</div>
        </div>
        """, unsafe_allow_html=True)
        if envelope_info:
            st.markdown(f"<div class='cp-doc-info'>Signed document: <b>{envelope_info.get('title', '')}</b></div>",
                        unsafe_allow_html=True)
        st.markdown("</div></div>", unsafe_allow_html=True)
        st.stop()

    # ── STEP 1: Form ────────────────────────────────────────────────────────
    if not form_done:
        questions = load_questions()
        st.markdown("<p style='font-size:.9rem;color:#475569;margin:0 0 20px 0;'>"
                    "Please fill in the information below. After submitting, you'll proceed to sign the document.</p>",
                    unsafe_allow_html=True)

        with st.form("client_form"):
            answers: dict[str, str] = {}
            for i, q in enumerate(questions):
                st.markdown(f"<div class='cp-field-label'>{q}</div>", unsafe_allow_html=True)
                if i == len(questions) - 1:
                    st.markdown("<div class='cp-field-hint'>Optional — add any context you'd like us to know.</div>",
                                unsafe_allow_html=True)
                answers[q] = st.text_area(
                    "Answer", key=f"q_{q}", label_visibility="collapsed",
                    height=80 if i < len(questions) - 1 else 100,
                )
                if i < len(questions) - 1:
                    st.markdown("<div style='height:16px;'></div>", unsafe_allow_html=True)

            submit_label = "Continue to signing →" if has_signing else "Submit"
            submitted = st.form_submit_button(submit_label, type="primary",
                                               width='stretch')

        if submitted:
            missing = [q for q, a in answers.items() if not a.strip() and q != questions[-1]]
            if missing:
                st.error("Please answer: " + "; ".join(missing))
            else:
                resp_headers = [
                    "response_id", "submitted_at", "client_token",
                    "client_name", "client_email", "company",
                    "signpack_id", "recipient_id",
                ] + questions
                row = {
                    "response_id": uuid.uuid4().hex[:12],
                    "submitted_at": datetime.now().isoformat(timespec="seconds"),
                    "client_token": token,
                    "client_name": client_name,
                    "client_email": client_email,
                    "company": client_company,
                    "signpack_id": sid or "",
                    "recipient_id": rid or "",
                }
                row.update(answers)
                try:
                    _ensure_tab(RESPONSES_TAB, resp_headers)
                    append_row(RESPONSES_TAB, row, resp_headers)
                    st.session_state.cp_form_submitted = True
                    st.rerun()
                except Exception as e:
                    st.error(f"Could not save responses: {e}")

        st.markdown("</div></div>", unsafe_allow_html=True)
        st.stop()

    # ── STEP 2: Sign ────────────────────────────────────────────────────────
    if not has_signing:
        st.success("Form submitted. Thank you!")
        st.markdown("</div></div>", unsafe_allow_html=True)
        st.stop()

    if recipient_info is None:
        st.error(f"Recipient #{rid} not found on envelope #{sid}.")
        st.markdown("</div></div>", unsafe_allow_html=True)
        st.stop()

    rstatus = recipient_info.get("status", "")
    if rstatus == "Declined":
        st.warning("You have declined this document.")
        st.markdown("</div></div>", unsafe_allow_html=True)
        st.stop()

    if rstatus == "Signed":
        st.success("Document signed. Thank you!")
        st.markdown("</div></div>", unsafe_allow_html=True)
        st.stop()

    # ── Check for completion event from ERPDocs postMessage ─────────────
    # ERPDocs sends postMessage to the parent frame on signing complete/decline.
    # Our listener (below, near the iframe) navigates the parent window here
    # with ?event=signing_complete or ?event=decline appended.
    completion_event = st.query_params.get("event", "")
    if completion_event in ("signing_complete", "decline"):
        if completion_event == "signing_complete":
            _title = "Document signed"
            _msg = ("Thank you \u2014 your signature has been recorded. A signed copy "
                    "and completion notification will be emailed to you shortly. "
                    "You may now close this window.")
            _icon_bg, _icon_color, _icon_char = "#16a34a", "#fff", "\u2713"
        else:
            _title = "Signing declined"
            _msg = ("You\u2019ve declined to sign this document. The sender has been "
                    "notified. If this was a mistake, please contact the sender "
                    "to request a new signing link.")
            _icon_bg, _icon_color, _icon_char = "#dc2626", "#fff", "\u2715"
        st.markdown(
            f"""
            <style>
              .complete-wrap {{
                max-width: 560px; margin: 2rem auto 0; padding: 0 1rem;
              }}
              .complete-card {{
                background: #fff; border: 1px solid #e2e8f0; border-radius: 14px;
                padding: 40px 36px; text-align: center;
                box-shadow: 0 1px 3px rgba(0,0,0,.06);
              }}
              .complete-icon {{
                width: 64px; height: 64px; border-radius: 50%;
                background: {_icon_bg}; color: {_icon_color};
                display: inline-flex; align-items: center; justify-content: center;
                font-size: 1.8rem; font-weight: 700; margin-bottom: 20px;
              }}
              .complete-card h1 {{
                font-size: 1.5rem; font-weight: 700; color: #0f172a;
                margin: 0 0 10px 0;
              }}
              .complete-card p {{
                font-size: .92rem; color: #475569; line-height: 1.6;
                margin: 0 0 8px 0;
              }}
              .complete-footer {{
                margin-top: 22px; font-size: .8rem; color: #94a3b8;
              }}
            </style>
            <div class="complete-wrap">
              <div class="complete-card">
                <div class="complete-icon">{_icon_char}</div>
                <h1>{_title}</h1>
                <p>{_msg}</p>
                <div class="complete-footer">ERPDocs \u00b7 Secure e-signature</div>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown("</div></div>", unsafe_allow_html=True)
        st.stop()

    # Mint embed token — auto-mint on page load, no extra button click
    _ss = load_settings()
    origin = _ss.get("embed_origin") or FORM_BASE_URL
    tok = st.session_state.esign_active_token

    # Auto-clear expired tokens
    if tok and str(tok.get("recipient_id")) == str(rid):
        try:
            exp = datetime.fromisoformat(tok["expires_at"].replace("Z", "+00:00"))
            if datetime.now(timezone.utc) > exp:
                st.session_state.esign_active_token = None
                tok = None
        except Exception:
            pass

    has_valid_token = tok and str(tok.get("recipient_id")) == str(rid)

    if not has_valid_token:
        # Mint token immediately — use settings for ttl, suppress, auth method.
        # NO redirect_url is sent: Streamlit Community Cloud cannot be loaded
        # inside an iframe via navigation (share.streamlit.io redirect-loops).
        # Instead, we listen for ERPDocs postMessage events (complete/decline)
        # in the parent frame and navigate the parent window ourselves.
        payload = {
            "origin": origin,
            "ttl_seconds": int(_ss.get("ttl_seconds", 300)),
            "signer_authenticated": {
                "method": _ss.get("signer_auth_method", "portal_link"),
                "identifier": recipient_info.get("email"),
                "authenticated_at": datetime.now(timezone.utc).isoformat(),
            },
        }
        if _ss.get("redirect_url"):
            payload["redirect_url"] = _ss["redirect_url"]
        if _ss.get("suppress_completion_email"):
            payload["suppress_completion_email"] = True
        with st.spinner("Preparing signing session…"):
            tok_data, tok_err = api.embed_token(sid, rid, payload)
        if tok_err:
            e = tok_err.get("error", {})
            st.error(f"`{e.get('code')}` — {e.get('message')}")
            if e.get("code") == "origin_not_allowed":
                st.info(f"Origin **{origin}** is not registered on the API key. "
                        "The sender must add it in ERPDocs Admin → API & Webhooks → Embed Origins.")
        else:
            st.session_state.esign_active_token = {
                "recipient_id": rid,
                "sign_url": tok_data["sign_url"],
                "expires_at": tok_data["expires_at"],
            }
            st.rerun()
    else:
        # Fullscreen signing
        st.markdown(
            "<style>.block-container { max-width: 100% !important; padding-top: 1rem !important; }</style>",
            unsafe_allow_html=True,
        )
        st.iframe(tok["sign_url"], height=int(_ss.get("iframe_height", 820)))

        # postMessage listener: ERPDocs posts 'complete' or 'decline' events
        # to the parent frame (this Streamlit app) when the signer finishes.
        # We catch them and navigate the PARENT window (not the iframe) to
        # add ?event=signing_complete|decline, which triggers the completion
        # card above. This avoids iframe navigation to Streamlit Cloud (which
        # redirect-loops on share.streamlit.io).
        # postMessage listener: ERPDocs posts 'complete' or 'decline' events
        # to the parent frame (this Streamlit app) when the signer finishes.
        # We use st.components.v1.html (not st.html) because it creates a
        # persistent iframe whose script reliably executes and stays alive.
        # The listener watches window.parent (the Streamlit app window) for
        # ERPDocs messages, then navigates the parent to add ?event=... which
        # triggers the completion card above.
        components.html("""
        <script>
        (function() {
            window.parent.addEventListener('message', function(event) {
                var d = event.data || {};
                if (d.type === 'complete') {
                    var url = new URL(window.parent.location.href);
                    url.searchParams.set('event', 'signing_complete');
                    window.parent.location.href = url.toString();
                } else if (d.type === 'decline') {
                    var url = new URL(window.parent.location.href);
                    url.searchParams.set('event', 'decline');
                    window.parent.location.href = url.toString();
                } else if (d.type === 'redirect') {
                    var ev = d.event || 'signing_complete';
                    var url = new URL(window.parent.location.href);
                    url.searchParams.set('event', ev);
                    window.parent.location.href = url.toString();
                }
            });
        })();
        </script>
        """, height=0, width=0)

        if st.button("🔄 I signed — refresh status", width='stretch'):
            st.session_state.esign_active_token = None
            st.session_state.cp_form_submitted = False
            st.rerun()

    st.markdown("</div></div>", unsafe_allow_html=True)
    st.stop()


# ═══════════════════════════════════════════════════════════════════════════
# SIGNING COMPLETION VIEW (redirect_url API target)
# URL: ?view=complete&event=signing_complete|decline
# Used when testing the ERPDocs redirect_url API. ERPDocs navigates the
# signing iframe here with ?event= appended. Works locally (localhost);
# on Streamlit Community Cloud this will redirect-loop — use postMessage
# there instead (the form view's completion_event check handles that).
# ═══════════════════════════════════════════════════════════════════════════
if view == "complete":
    event = st.query_params.get("event", "")

    if event == "decline":
        title = "Signing declined"
        message = ("You\u2019ve declined to sign this document. The sender has been "
                   "notified. If this was a mistake, please contact the sender "
                   "to request a new signing link.")
        icon_bg, icon_color, icon_char = "#dc2626", "#fff", "\u2715"
    elif event == "signing_complete":
        title = "Document signed"
        message = ("Thank you \u2014 your signature has been recorded. A signed copy "
                   "and completion notification will be emailed to you shortly. "
                   "You may now close this window.")
        icon_bg, icon_color, icon_char = "#16a34a", "#fff", "\u2713"
    else:
        title = "Signing session ended"
        message = ("Your signing session has ended. If you expected a "
                   "confirmation, please check your email or contact the sender.")
        icon_bg, icon_color, icon_char = "#0891b2", "#fff", "\u2713"

    st.markdown(
        f"""
        <style>
          .complete-wrap {{
            max-width: 560px; margin: 3rem auto 0; padding: 0 1rem;
          }}
          .complete-card {{
            background: #fff; border: 1px solid #e2e8f0; border-radius: 14px;
            padding: 40px 36px; text-align: center;
            box-shadow: 0 1px 3px rgba(0,0,0,.06);
          }}
          .complete-icon {{
            width: 64px; height: 64px; border-radius: 50%;
            background: {icon_bg}; color: {icon_color};
            display: inline-flex; align-items: center; justify-content: center;
            font-size: 1.8rem; font-weight: 700; margin-bottom: 20px;
          }}
          .complete-card h1 {{
            font-size: 1.5rem; font-weight: 700; color: #0f172a;
            margin: 0 0 10px 0;
          }}
          .complete-card p {{
            font-size: .92rem; color: #475569; line-height: 1.6;
            margin: 0 0 8px 0;
          }}
          .complete-footer {{
            margin-top: 22px; font-size: .8rem; color: #94a3b8;
          }}
          .complete-debug {{
            margin-top: 20px; padding: 10px 14px; background: #f1f5f9;
            border-radius: 8px; font-size: .76rem; color: #64748b;
            text-align: left; font-family: monospace; word-break: break-all;
          }}
        </style>
        <div class="complete-wrap">
          <div class="complete-card">
            <div class="complete-icon">{icon_char}</div>
            <h1>{title}</h1>
            <p>{message}</p>
            <div class="complete-footer">ERPDocs \u00b7 Secure e-signature</div>
            <div class="complete-debug">event = {event or "(none)"}<br>view = complete<br>url = {st.query_params.get("view", "")}</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.stop()


# ═══════════════════════════════════════════════════════════════════════════
# SENDER VIEW (default) — one unified flow
# ═══════════════════════════════════════════════════════════════════════════
st.title("Client Outreach")
st.caption("Send clients a form to fill. Optionally attach an e-sign agreement. "
           "One email, one link — client fills the form, then signs if attached.")

# ── Sidebar ─────────────────────────────────────────────────────────────────
st.sidebar.markdown('<div class="ce-tag">Connection</div>', unsafe_allow_html=True)

if not WORKBOOK_ID:
    st.sidebar.error("hub_workbook_id missing in secrets.json")
if not _has_sheets_creds():
    st.sidebar.error("Google service account not configured")
    st.info(
        "No Google credentials found. Either set `google_service_account_path` "
        "in `config.json` (local), or add a `[gcp_service_account]` table to the "
        "app's secrets (hosted). Then reload."
    )
    st.stop()

st.sidebar.caption(f"Workbook: `{WORKBOOK_ID[:12]}…`")

st.sidebar.divider()
st.sidebar.markdown('<div class="ce-tag">Email (SMTP)</div>', unsafe_allow_html=True)
smtp_ok = bool(SMTP_SERVER and SMTP_USERNAME and SMTP_PASSWORD)
if not smtp_ok:
    st.sidebar.warning("SMTP not configured. Enter below to send this session.")
    SMTP_SERVER = st.sidebar.text_input("SMTP server", value=SMTP_SERVER, key="sb_smtp_host")
    SMTP_PORT = st.sidebar.number_input("Port", value=SMTP_PORT, key="sb_smtp_port")
    SMTP_USERNAME = st.sidebar.text_input("Username", value=SMTP_USERNAME, key="sb_smtp_user")
    SMTP_PASSWORD = st.sidebar.text_input("Password", type="password", key="sb_smtp_pass")
    SMTP_FROM_EMAIL = st.sidebar.text_input("From email", value=SMTP_FROM_EMAIL, key="sb_smtp_from")
else:
    st.sidebar.caption(f"Server: `{SMTP_SERVER}:{SMTP_PORT}`")
    st.sidebar.caption(f"From: {SMTP_FROM_NAME} <{SMTP_FROM_EMAIL or SMTP_USERNAME}>")

st.sidebar.divider()
st.sidebar.markdown('<div class="ce-tag">App URL</div>', unsafe_allow_html=True)
FORM_BASE_URL = st.sidebar.text_input(
    "Base URL for client links", value=FORM_BASE_URL,
    help="Where this app is reachable. Used to build the link in each email. "
         "Local testing: http://localhost:8501",
).rstrip("/")

st.sidebar.divider()
st.sidebar.markdown('<div class="ce-tag">ERPDocs E-Sign</div>', unsafe_allow_html=True)

# Settings override already applied above (before view router).
# Sidebar allows quick inline override of base URL / key for the sender view.
ERPDOCS_BASE_URL = st.sidebar.text_input(
    "ERPDocs base URL", value=ERPDOCS_BASE_URL,
    help="https://staging.erpdocs.com for staging, https://www.erpdocs.com for production.",
).rstrip("/")
if not ERPDOCS_API_KEY:
    ERPDOCS_API_KEY = st.sidebar.text_input(
        "API key", type="password", key="sb_erpdocs_key",
        help="LIVE key with scopes: template.read, signpack.write, signpack.read, "
             "signpack.embed, brand.read, brand.write",
    )
else:
    st.sidebar.caption("API key: ✓ set")

erpdocs_api = ERPDocs(ERPDOCS_BASE_URL, ERPDOCS_API_KEY) if ERPDOCS_API_KEY else None
if erpdocs_api:
    if st.sidebar.button("Test ERPDocs connection", key="test_erpdocs"):
        data, err = erpdocs_api.ping()
        if err:
            st.sidebar.error(err.get("error", {}).get("message", "Failed"))
        else:
            st.sidebar.success(f"{data.get('org_name')} · {data.get('environment')}")
            needed = ["template.read", "signpack.write", "signpack.read",
                      "signpack.embed", "brand.read", "brand.write"]
            scopes = data.get("scopes", [])
            missing = [s for s in needed if s not in scopes]
            if missing:
                st.sidebar.error("Missing scopes: " + ", ".join(missing))

st.sidebar.markdown(
    '<a href="?view=settings" target="_self" '
    'style="display:inline-block;padding:6px 12px;background:#0891b2;color:#fff;'
    'border-radius:6px;text-decoration:none;font-size:.85rem;font-weight:600;margin-bottom:6px;">'
    '⚙️ Settings</a>',
    unsafe_allow_html=True,
)
st.sidebar.markdown(
    '<a href="?view=tools" target="_self" '
    'style="display:inline-block;padding:6px 12px;background:#475569;color:#fff;'
    'border-radius:6px;text-decoration:none;font-size:.85rem;font-weight:600;">'
    '🛠️ E-Sign Tools</a>',
    unsafe_allow_html=True,
)

st.sidebar.divider()
if st.sidebar.button("Reload client data", key="reload"):
    st.cache_resource.clear()
    st.rerun()

# ── Load client data ────────────────────────────────────────────────────────
try:
    clients = read_tab(CLIENT_TAB)
except Exception as e:
    st.error(f"Could not read '{CLIENT_TAB}' tab: {e}")
    st.info(f"Create a tab named exactly **{CLIENT_TAB}** in the workbook with headers "
            "like: S.No, Name, Email Id, Company Name, ...")
    st.stop()

if clients.empty:
    st.warning(
        f"The **{CLIENT_TAB}** tab is empty or missing. Add a header row "
        "(S.No, Name, Email Id, Company Name, ...) and at least one client row, "
        f"then click **Reload client data**.\n\n"
        f"Workbook: https://docs.google.com/spreadsheets/d/{WORKBOOK_ID}/edit"
    )
    st.stop()

email_c = _email_col(clients)
if email_c is None:
    st.error(
        f"No email column found in '{CLIENT_TAB}'. Expected a header like "
        "'Email Id', 'Email', or 'Email ID'. Found: {list(clients.columns)}"
    )
    st.stop()

name_c = _name_col(clients)
company_c = _company_col(clients)

# ── 1. Client list ──────────────────────────────────────────────────────────
st.markdown(f"**{len(clients)} client(s) loaded** from the '{CLIENT_TAB}' tab.")
display_cols = [c for c in clients.columns if c.lower() not in ("form link", "token")]
st.dataframe(clients[display_cols], width='stretch', hide_index=True)

# ── 2. Form questions ───────────────────────────────────────────────────────
st.divider()
st.markdown("### Step 1 — Form questions")
st.caption("These questions appear on the client's form page. One per line.")
with st.expander("Edit form questions", expanded=False):
    questions = load_questions()
    q_text = st.text_area("One question per line", value="\n".join(questions), height=120, key="q_edit")
    if st.button("Save questions", key="save_q"):
        qs = [q.strip() for q in q_text.splitlines() if q.strip()]
        try:
            save_questions(qs)
            st.success(f"Saved {len(qs)} question(s).")
        except Exception as e:
            st.error(f"Could not save: {e}")

# ── 3. Email composer ───────────────────────────────────────────────────────
st.divider()
st.markdown("### Step 2 — Email content")
st.caption("Placeholders: {{Name}} {{Company}} {{Email}} {{FormLink}} — replaced per client.")

default_subject = "Action required: Please complete your form"
default_body = (
    "<p>Hi {{Name}},</p>"
    "<p>We need you to complete a short form. "
    "If an agreement is attached, you'll be able to sign it after filling the form.</p>"
    "<p><a href=\"{{FormLink}}\">Click here to start →</a></p>"
    "<p>If the link doesn't work, copy this into your browser:<br>"
    "<small>{{FormLink}}</small></p>"
    "<p>Thanks,<br>Accounts Team</p>"
)

c_subj, c_cc = st.columns([3, 1])
with c_subj:
    subject_tmpl = st.text_input("Subject", value=default_subject, key="subj")
with c_cc:
    cc = st.text_input("CC (optional, comma-separated)", key="cc")
body_tmpl = st.text_area("Body (HTML)", value=default_body, height=150, key="body")

# ── 4. E-sign toggle ────────────────────────────────────────────────────────
st.divider()
st.markdown("### Step 3 — E-sign agreement (optional)")

# Show active signature settings summary
if _sig_settings.get("default_brand_id") or _sig_settings.get("suppress_completion_email"):
    parts = []
    if _sig_settings.get("default_brand_id"):
        parts.append(f"Brand ID: `{_sig_settings['default_brand_id']}`")
    if _sig_settings.get("suppress_completion_email"):
        parts.append("Completion email: **suppressed**")
    else:
        parts.append("Completion email: **sent**")
    parts.append(f"Expiry: `{_sig_settings.get('default_expires_days', 30)} days`")
    if not _sig_settings.get("default_send_immediately", True):
        parts.append("Mode: **Draft** (not sent immediately)")
    st.info("**Active signature settings:** " + " · ".join(parts) +
            f" · [⚙️ Change](?view=settings)")

enable_esign = st.checkbox(
    "Attach an e-sign agreement to the form",
    value=False, key="enable_esign",
    help="When enabled, each client will see the signing page after submitting the form. "
         "Requires an ERPDocs API key (sidebar).",
)

# E-sign config state
tpl_chosen = None
tpl_detail = None
tpl_roles = []
tpl_prefillable = []
prefill_map: dict[str, str] = {}
sender_email = ""
expires_days = int(_sig_settings.get("default_expires_days", 30))

if enable_esign:
    if not erpdocs_api:
        st.error("ERPDocs API key required. Add it in the sidebar to enable e-signing.")
        enable_esign = False
    else:
        # Load templates
        tpl_data, tpl_err = erpdocs_api.templates()
        if tpl_err:
            st.error(f"Could not load templates: {tpl_err.get('error', {}).get('message')}")
            enable_esign = False
        else:
            templates = ERPDocs.rows(tpl_data)
            if not templates:
                st.warning("No templates found. Create one in the ERPDocs UI first.")
                enable_esign = False
            else:
                tpl_labels = {f"{t.get('name')} (id {t.get('id')})": t for t in templates}
                chosen_label = st.selectbox("Agreement template", list(tpl_labels.keys()), key="esign_tpl")
                tpl_chosen = tpl_labels[chosen_label]

                # Read template contract
                tpl_detail, d_err = erpdocs_api.template(tpl_chosen["id"])
                if d_err:
                    st.error(f"Could not read template: {d_err.get('error', {}).get('message')}")
                else:
                    tpl_roles = tpl_detail.get("roles", [])
                    tpl_prefillable = tpl_detail.get("prefillable_fields", [])

                    # Show roles + fields
                    c_tpl1, c_tpl2 = st.columns(2)
                    with c_tpl1:
                        st.caption("**Template roles**")
                        for r in tpl_roles:
                            st.markdown(f"- `{r.get('name')}` ({r.get('type', 'Signer')})")
                    with c_tpl2:
                        st.caption("**Prefillable fields**")
                        if tpl_prefillable:
                            for f in tpl_prefillable:
                                st.markdown(f"- `{f}`")
                        else:
                            st.markdown("- *(none)*")

                    # Sender email
                    users_data, _ = erpdocs_api.users()
                    user_emails = [u.get("email") for u in ERPDocs.rows(users_data) if u.get("email")]
                    if user_emails:
                        sender_email = st.selectbox("Sender (org member)", user_emails, key="esign_sender")
                    else:
                        sender_email = st.text_input("Sender email", key="esign_sender_manual")

                    # Prefill mapping
                    if tpl_prefillable:
                        st.caption("**Map sheet columns → template fields**")
                        sheet_cols = list(clients.columns)
                        for label in tpl_prefillable:
                            default_col = ""
                            for col in sheet_cols:
                                if col.lower().strip() == label.lower().strip():
                                    default_col = col
                                    break
                            opts = ["(none)"] + sheet_cols
                            mapped = st.selectbox(
                                f"{label}", options=opts,
                                index=0 if not default_col else opts.index(default_col),
                                key=f"pf_map_{label}",
                            )
                            if mapped != "(none)":
                                prefill_map[label] = mapped

                    expires_days = st.number_input(
                        "Agreement expires in (days)", min_value=1, max_value=365,
                        value=int(_sig_settings.get("default_expires_days", 30)),
                        key="es_expiry",
                    )

# ── 5. Select clients ───────────────────────────────────────────────────────
st.divider()
st.markdown("### Step 4 — Select clients")
sel_all = st.checkbox("Select all clients", value=False, key="sel_all")
selected_idx: list[int] = []
if sel_all:
    selected_idx = list(range(len(clients)))
else:
    chosen = st.multiselect(
        "Or pick specific clients",
        options=list(range(len(clients))),
        format_func=lambda i: f"{clients.iloc[i].get(name_c or '', '') or '(no name)'} — {clients.iloc[i][email_c]}",
        key="sel_multi",
    )
    selected_idx = chosen

# ── 6. Send ─────────────────────────────────────────────────────────────────
st.divider()
st.markdown("### Step 5 — Send")

if not (SMTP_SERVER and SMTP_USERNAME and SMTP_PASSWORD):
    st.warning("SMTP credentials required in the sidebar to send emails.")
elif not selected_idx:
    st.info("Select at least one client above.")
else:
    send_label = f"Send to {len(selected_idx)} client(s)"
    if enable_esign and tpl_chosen:
        send_label += " (form + e-sign)"
    else:
        send_label += " (form only)"

    if st.button(send_label, type="primary", key="send_all"):
        log = st.session_state.last_send_log
        ok, err = 0, 0

        for idx in selected_idx:
            client = clients.iloc[idx]
            to_email = str(client[email_c]).strip()
            client_name = str(client.get(name_c, "")) if name_c else ""
            client_company = str(client.get(company_c, "")) if company_c else ""

            if not to_email:
                err += 1
                log.append({"client": "(no email)", "status": "error", "msg": "blank email"})
                continue

            # Determine the link for this client
            sid = ""
            rid = ""

            if enable_esign and tpl_chosen and sender_email:
                # Create envelope for this client
                pf = {}
                for label, col in prefill_map.items():
                    val = str(client.get(col, "")).strip()
                    if val:
                        pf[label] = val

                recips = {}
                for r in tpl_roles:
                    rname = r.get("name")
                    rtype = r.get("type", "Signer")
                    if rtype.lower() in ("signer", "reviewer"):
                        recips[rname] = {
                            "name": client_name,
                            "email": to_email,
                            "embedded": True,
                        }
                        if _sig_settings.get("suppress_completion_email"):
                            recips[rname]["suppress_completion_email"] = True
                    elif rtype.lower() == "cc":
                        recips[rname] = {"email": to_email}

                payload = {
                    "sender_email": sender_email,
                    "title": f"Agreement — {client_company or client_name or to_email}",
                    "recipients": recips,
                    "prefill": pf,
                    "expires_in_days": int(expires_days),
                }
                if _sig_settings.get("default_brand_id"):
                    payload["brand_id"] = int(_sig_settings["default_brand_id"])
                if not _sig_settings.get("default_send_immediately", True):
                    payload["send"] = False

                # Log the actual payload for debugging
                log.append({"client": to_email, "status": "info",
                            "msg": f"payload: brand_id={payload.get('brand_id', 'none')}, "
                                   f"suppress_ce={_sig_settings.get('suppress_completion_email')}, "
                                   f"expires={payload.get('expires_in_days')}, "
                                   f"send={payload.get('send', True)}"})

                data, env_err = erpdocs_api.send_from_template(tpl_chosen["id"], payload)
                if env_err:
                    e = env_err.get("error", {})
                    err += 1
                    log.append({"client": to_email, "status": "error",
                                "msg": f"envelope: {e.get('code')} — {e.get('message')}"})
                    continue

                sid = str(data.get("id", ""))
                # Get recipient ID
                sp_detail, _ = erpdocs_api.signpack(sid)
                if sp_detail:
                    for rec in sp_detail.get("recipients", []):
                        if rec.get("role", "").lower() in ("signer", "reviewer"):
                            rid = str(rec.get("id", ""))
                            break

                # Check for unmatched prefill
                unmatched = data.get("prefill_unmatched") or []
                if unmatched:
                    st.warning(f"{to_email}: prefill labels not matched: {', '.join(unmatched)}")

            # Build the link (form + sign if envelope was created)
            token = client_token(to_email)
            if sid and rid:
                link = f"{FORM_BASE_URL}/?view=form&token={token}&sid={sid}&rid={rid}"
            else:
                link = f"{FORM_BASE_URL}/?view=form&token={token}"

            # Send email
            try:
                # Inject the actual link (with sid/rid if e-sign) before rendering
                body_with_link = body_tmpl.replace("{{FormLink}}", link)
                subj_with_link = subject_tmpl.replace("{{FormLink}}", link)
                html = render_body(body_with_link, client)
                subj = render_body(subj_with_link, client)
                send_email(to_email, subj, html, cc=cc)
                ok += 1
                log.append({"client": to_email, "status": "sent",
                            "msg": f"envelope #{sid}" if sid else "form only"})
            except Exception as e:
                err += 1
                log.append({"client": to_email, "status": "error", "msg": str(e)})

        st.session_state.last_send_log = log
        if err == 0:
            st.success(f"Sent {ok} email(s).")
        else:
            st.warning(f"Sent {ok}, {err} failed. See log below.")
        st.rerun()

# ── Send log ────────────────────────────────────────────────────────────────
log = st.session_state.last_send_log
if log:
    st.divider()
    st.markdown("**Send log**")
    log_df = pd.DataFrame(log[-50:])
    st.dataframe(log_df, width='stretch', hide_index=True)
    if st.button("Clear log", key="clear_log"):
        st.session_state.last_send_log = []
        st.rerun()

# ── Responses ───────────────────────────────────────────────────────────────
st.divider()
st.markdown("### Form responses")
try:
    resp = read_tab(RESPONSES_TAB)
except Exception:
    resp = pd.DataFrame()
if resp.empty:
    st.caption("No responses yet.")
else:
    st.dataframe(resp, width='stretch', hide_index=True)

# ── Envelope tracking ───────────────────────────────────────────────────────
if erpdocs_api:
    st.divider()
    st.markdown("### Envelope tracking")
    with st.expander("View recent envelopes", expanded=False):
        if st.button("Fetch last 20 envelopes", key="fetch_envs"):
            sp_data, sp_err = erpdocs_api.signpacks(limit=20)
            if sp_err:
                st.error(f"Could not load: {sp_err.get('error', {}).get('message')}")
            else:
                envs = ERPDocs.rows(sp_data)
                if envs:
                    # Show a clean summary
                    summary = []
                    for env in envs:
                        recipients = env.get("recipients", [])
                        signer_status = "—"
                        for rec in recipients:
                            if rec.get("role", "").lower() in ("signer", "reviewer"):
                                signer_status = rec.get("status", "—")
                                break
                        summary.append({
                            "ID": env.get("id"),
                            "Title": env.get("title", ""),
                            "Status": env.get("status", "—"),
                            "Signer": signer_status,
                            "Created": env.get("created_at", "—"),
                        })
                    st.dataframe(pd.DataFrame(summary), width='stretch', hide_index=True)

                    # Download button for completed ones
                    completed = [e for e in envs if e.get("status") == "Completed"]
                    if completed:
                        st.caption(f"{len(completed)} completed envelope(s).")
                        for env in completed:
                            eid = env.get("id")
                            if st.button(f"Download signed PDF — #{eid}", key=f"dl_{eid}"):
                                dl, dl_err = erpdocs_api.download(eid)
                                if dl_err:
                                    st.error(f"Download failed: {dl_err.get('error', {}).get('message')}")
                                else:
                                    docs = dl.get("documents", [])
                                    if docs:
                                        doc = docs[0]
                                        st.markdown(f"**[{doc.get('filename')}]({doc.get('url')})**")
                                        st.caption(f"SHA-256: `{doc.get('sha256', '')[:30]}…`")
                else:
                    st.caption("No envelopes found.")
