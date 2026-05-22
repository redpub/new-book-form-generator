from __future__ import annotations

import streamlit as st

from app_config import ALLOWED_EMAIL_DOMAIN


def enforce_workspace_auth() -> None:
    """Verify Google OAuth login and enforce @red-publish.com domain."""
    if str(st.secrets.get("DISABLE_LOGIN", "false")).lower() == "true":
        return

    auth_settings = st.secrets.get("auth", {})
    if hasattr(auth_settings, "items"):
        auth_settings = dict(auth_settings.items())
    elif not isinstance(auth_settings, dict):
        auth_settings = {}

    missing: list[str] = []
    if not auth_settings.get("cookie_secret"):
        missing.append("auth.cookie_secret")
    if not auth_settings.get("redirect_uri"):
        missing.append("auth.redirect_uri")

    if missing:
        st.error("尚未完成驗證設定，請在 .streamlit/secrets.toml 補齊必要欄位。")
        st.code("\n".join(missing))
        st.stop()

    user_email = (getattr(st.user, "email", "") or "").strip().lower()
    if not user_email:
        try:
            st.login()
        except Exception as exc:
            st.error("驗證設定無效，請檢查 .streamlit/secrets.toml。")
            st.code(str(exc))
        st.stop()

    if not user_email.endswith(ALLOWED_EMAIL_DOMAIN):
        st.error("未授權：僅允許 @red-publish.com 帳號存取此應用程式。")
        st.caption(f"目前登入帳號：{user_email}")
        if st.button("登出", key="unauthorized_logout", type="primary"):
            st.logout()
        st.stop()
