from __future__ import annotations

import json
from datetime import datetime

import streamlit as st

from auth import enforce_workspace_auth
from debug_store import get_runs, get_run_detail


# ---------------------------------------------------------------------------
# Formatting helpers
# ---------------------------------------------------------------------------

def _fmt_timestamp(iso_str: str) -> str:
    try:
        return datetime.fromisoformat(iso_str).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return iso_str


def _fmt_elapsed(seconds: float) -> str:
    if seconds < 60:
        return f"{seconds:.1f} 秒"
    m, s = divmod(int(seconds), 60)
    return f"{m} 分 {s:02d} 秒"


def _status_icon(success: bool) -> str:
    return "✅" if success else "❌"


# ---------------------------------------------------------------------------
# Detail view
# ---------------------------------------------------------------------------

def _show_run_detail(run: dict) -> None:
    st.subheader(
        f"{_status_icon(run['success'])} {run['file_name']}"
    )

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("狀態", "成功" if run["success"] else "失敗")
    c2.metric("模型", run["model"])
    c3.metric("耗時", _fmt_elapsed(run["elapsed_seconds"]))
    c4.metric("時間", _fmt_timestamp(run.get("created_at", "")))

    if run.get("user_email"):
        st.caption(f"使用者：{run['user_email']}")

    if run.get("error_message"):
        st.error(f"錯誤訊息：{run['error_message']}")

    # Source file download
    if run.get("source_file"):
        st.download_button(
            label="⬇️ 下載原始 Word 檔案",
            data=run["source_file"],
            file_name=run["file_name"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    with st.expander("📤 原始 Prompt", expanded=False):
        st.text(run.get("raw_prompt") or "（無資料）")

    with st.expander("📥 原始回應", expanded=False):
        raw = run.get("raw_response") or ""
        if raw:
            try:
                st.json(json.loads(raw))
            except Exception:
                st.text(raw)
        else:
            st.caption("（無回應資料）")


# ---------------------------------------------------------------------------
# Page
# ---------------------------------------------------------------------------

def main() -> None:
    enforce_workspace_auth()

    st.title("🔍 除錯紀錄")
    st.caption(
        "查看所有擷取操作的詳細紀錄，包含使用模型、耗時、原始 Prompt 與 LLM 回應。"
    )

    with st.sidebar:
        st.header("⚙️ 設定")
        st.caption(f"已登入：{getattr(st.user, 'email', '未知帳號')}")
        if st.button("登出", key="sidebar_logout"):
            st.logout()
            st.stop()

    if "selected_run_id" not in st.session_state:
        st.session_state.selected_run_id = None

    runs = get_runs()

    if not runs:
        st.info("目前沒有操作紀錄。請先在主頁執行擷取後再查看。")
        return

    # ---- Detail view ----
    if st.session_state.selected_run_id:
        if st.button("⬅️ 返回列表"):
            st.session_state.selected_run_id = None
            st.rerun()
        run_detail = get_run_detail(st.session_state.selected_run_id)
        if run_detail:
            _show_run_detail(run_detail)
        else:
            st.warning("找不到該紀錄。")
            st.session_state.selected_run_id = None
        return

    # ---- List view ----
    st.subheader(f"最近 {len(runs)} 筆紀錄")

    # Table header
    hcols = st.columns([0.5, 1.8, 2.5, 2.0, 1.8, 1.2, 1])
    hcols[0].markdown("**狀態**")
    hcols[1].markdown("**時間**")
    hcols[2].markdown("**檔案**")
    hcols[3].markdown("**使用者**")
    hcols[4].markdown("**模型**")
    hcols[5].markdown("**耗時**")
    hcols[6].markdown("")
    st.markdown("<hr style='margin:4px 0;'>", unsafe_allow_html=True)

    for run in runs:
        cols = st.columns([0.5, 1.8, 2.5, 2.0, 1.8, 1.2, 1])
        cols[0].write(_status_icon(run["success"]))
        cols[1].caption(_fmt_timestamp(run.get("created_at", "")))
        cols[2].write(run["file_name"])
        cols[3].caption(run.get("user_email") or "—")
        cols[4].caption(run["model"])
        cols[5].caption(_fmt_elapsed(run["elapsed_seconds"]))
        if cols[6].button("查看", key=f"view_{run['id']}"):
            st.session_state.selected_run_id = run["id"]
            st.rerun()
        st.markdown(
            "<hr style='margin:2px 0;border:none;border-top:1px solid #eee;'>",
            unsafe_allow_html=True,
        )


if __name__ == "__main__":
    main()
