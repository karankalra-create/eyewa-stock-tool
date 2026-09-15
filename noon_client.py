"""Noon Partners API client — service-account JWT auth.

Auth flow (per Noon's Partners API docs, Authentication > Authenticating Your
Requests): build a short-lived JWT signed with the service account's private
key, POST it to the login endpoint to get back a session cookie, then use
that cookie (via a requests.Session) on every subsequent call. Every request
must carry a User-Agent header identifying the app.

Credentials are never hardcoded or committed. They come from Streamlit's
secrets store: a local `.streamlit/secrets.toml` (git-ignored — see
`.streamlit/secrets.toml.example` for the shape) when running locally, or the
app's own Secrets panel on Streamlit Community Cloud when deployed.
"""

from __future__ import annotations

import time
import uuid

import jwt
import requests
import streamlit as st

BASE_URL = "https://noon-api-gateway.noon.partners"
USER_AGENT = "EyewaStockTool/1.0"

REQUIRED_FIELDS = ("key_id", "private_key", "channel_identifier", "project_code")


class NoonAuthError(RuntimeError):
    """Raised when Noon credentials are missing/invalid or a Noon call fails."""


def _credentials() -> dict:
    try:
        creds = st.secrets["noon"]
    except Exception as exc:  # noqa: BLE001 — st.secrets raises various things when absent
        raise NoonAuthError(
            "Noon credentials aren't configured. Add a [noon] section with "
            "key_id, private_key, channel_identifier and project_code to "
            ".streamlit/secrets.toml (local) or the app's Secrets panel "
            "(Streamlit Community Cloud)."
        ) from exc
    missing = [f for f in REQUIRED_FIELDS if f not in creds]
    if missing:
        raise NoonAuthError(f"Noon credentials are missing: {', '.join(missing)}")
    return dict(creds)


def _create_jwt(creds: dict) -> str:
    """Signed RS256 JWT for the login request, per Noon's docs."""
    return jwt.encode(
        {
            "sub": creds["key_id"],
            "iat": int(time.time()),
            "jti": str(uuid.uuid4()),
        },
        creds["private_key"],
        algorithm="RS256",
    )


def get_authenticated_session() -> requests.Session:
    """Log in with the service-account key; returns a session carrying the auth cookie."""
    creds = _credentials()
    session = requests.Session()
    session.headers.update({"User-Agent": USER_AGENT})

    response = session.post(
        f"{BASE_URL}/identity/public/v1/api/login",
        json={
            "token": _create_jwt(creds),
            "default_project_code": creds["project_code"],
        },
        timeout=30,
    )
    if response.status_code != 200:
        raise NoonAuthError(
            f"Noon login failed with HTTP {response.status_code}: {response.text}"
        )
    return session


def whoami(session: requests.Session | None = None) -> dict:
    """Confirm the session is authenticated. Returns Noon's identity payload."""
    session = session or get_authenticated_session()
    response = session.get(f"{BASE_URL}/identity/v1/whoami", timeout=30)
    if response.status_code != 200:
        raise NoonAuthError(
            f"Whoami failed with HTTP {response.status_code}: {response.text}"
        )
    return response.json()
