from __future__ import annotations

import streamlit as st

from app_config import ALLOWED_EMAIL_DOMAIN


def enforce_workspace_auth() -> None:
    """Verify Google OAuth login and enforce @red-publish.com domain."""
    if str(st.secrets.get("DISABLE_LOGIN", "false")).lower() == "true":
        return

    user_email = (getattr(st.user, "email", "") or "").strip().lower()
    if not user_email:
        st.login()
        st.stop()

    if not user_email.endswith(ALLOWED_EMAIL_DOMAIN):
        st.error("❌ 未授權：僅允許 @red-publish.com 帳號存取此應用程式。")
        st.caption(f"目前登入帳號：{user_email or '未知'}")
        if st.button("登出", key="unauthorized_logout", type="primary"):
            st.logout()
        st.stop()
