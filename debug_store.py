"""Supabase-backed debug run store.

Persists extraction run records (including optional source file bytes)
into the new_book_form_generator schema on Supabase.

Table used:
  - IS_PROD=true  → new_book_form_generator.extraction_runs
  - IS_PROD=false → new_book_form_generator.dev_extraction_runs
"""
from __future__ import annotations

import streamlit as st
from supabase import create_client, Client


_SCHEMA = "new_book_form_generator"


def _table() -> str:
    is_prod = str(st.secrets.get("IS_PROD", "false")).lower() == "true"
    return "extraction_runs" if is_prod else "dev_extraction_runs"


@st.cache_resource
def _get_client() -> Client:
    url = st.secrets["connections"]["supabase"]["SUPABASE_URL"]
    key = st.secrets["connections"]["supabase"]["SUPABASE_KEY"]
    return create_client(url, key)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def add_run(
    *,
    file_name: str,
    model: str,
    raw_prompt: str,
    raw_response: str,
    elapsed_seconds: float,
    success: bool,
    error_message: str = "",
    user_email: str = "",
    source_file: bytes | None = None,
) -> str | None:
    """Insert a run record into Supabase and return its UUID (or None on error)."""
    try:
        row: dict = {
            "file_name": file_name,
            "model": model,
            "elapsed_seconds": round(elapsed_seconds, 2),
            "success": success,
            "error_message": error_message or None,
            "raw_prompt": raw_prompt or None,
            "raw_response": raw_response or None,
            "user_email": user_email or None,
        }
        if source_file is not None:
            # PostgREST encodes BYTEA as a hex-escaped string: \x<hex>
            row["source_file"] = "\\x" + source_file.hex()

        client = _get_client()
        resp = client.schema(_SCHEMA).table(_table()).insert(row).execute()
        if resp.data:
            return resp.data[0]["id"]
    except Exception as exc:  # noqa: BLE001
        print(f"[debug_store] add_run failed: {exc}")
    return None


def get_runs(limit: int = 50) -> list[dict]:
    """Return the most recent runs (newest first)."""
    try:
        client = _get_client()
        resp = (
            client
            .schema(_SCHEMA)
            .table(_table())
            .select("id, created_at, user_email, file_name, model, elapsed_seconds, success, error_message")
            .order("created_at", desc=True)
            .limit(limit)
            .execute()
        )
        return resp.data or []
    except Exception as exc:  # noqa: BLE001
        print(f"[debug_store] get_runs failed: {exc}")
        return []


def get_run_detail(run_id: str) -> dict | None:
    """Return a single run including raw_prompt, raw_response, and source_file."""
    try:
        client = _get_client()
        resp = (
            client
            .schema(_SCHEMA)
            .table(_table())
            .select("*")
            .eq("id", run_id)
            .limit(1)
            .execute()
        )
        if resp.data:
            row = resp.data[0]
            # Decode hex-escaped BYTEA back to bytes
            sf = row.get("source_file")
            if sf:
                try:
                    if isinstance(sf, str) and sf.startswith("\\x"):
                        row["source_file"] = bytes.fromhex(sf[2:])
                    elif isinstance(sf, (bytes, bytearray)):
                        row["source_file"] = bytes(sf)
                    else:
                        row["source_file"] = None
                except Exception:
                    row["source_file"] = None
            return row
    except Exception as exc:  # noqa: BLE001
        print(f"[debug_store] get_run_detail failed: {exc}")
    return None
