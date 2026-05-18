import streamlit as st

st.set_page_config(page_title="書籍表單產生器", page_icon="📘", layout="wide")

home = st.Page("pages/home.py", title="主頁", icon="📝", default=True)
debug = st.Page("pages/debug_log.py", title="除錯紀錄", icon="🔍")

pg = st.navigation([home, debug])
pg.run()
