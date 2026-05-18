"""In-memory debug run store.

Uses st.cache_resource so the list persists across page reruns and
page switches for as long as the Streamlit server is running.
Entries are prepended (newest first) and are lost on server restart.
"""
from __future__ import annotations

import uuid
from datetime import datetime

import streamlit as st


@st.cache_resource
def _get_store() -> list[dict]:
    return []


def add_run(
    *,
    file_name: str,
    model: str,
    raw_prompt: str,
    raw_response: str,
    elapsed_seconds: float,
    success: bool,
    error_message: str = "",
) -> str:
    """Append a run record and return its ID."""
    store = _get_store()
    run_id = str(uuid.uuid4())
    store.insert(0, {
        "id": run_id,
        "timestamp": datetime.now().isoformat(),
        "file_name": file_name,
        "model": model,
        "elapsed_seconds": round(elapsed_seconds, 2),
        "success": success,
        "raw_prompt": raw_prompt,
        "raw_response": raw_response,
        "error_message": error_message,
    })
    return run_id


def get_runs() -> list[dict]:
    """Return a snapshot of all stored runs (newest first)."""
    return list(_get_store())
