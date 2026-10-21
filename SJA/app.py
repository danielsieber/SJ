"""
SJA – Ski Jumping Analysis
Thin entry point. All logic lives in src/.
Run with: python -m streamlit run app.py
"""
import streamlit as st

from src.config import init_paths, load_data_lists
from src.session_state import init_session_state
from src.ui.styles import inject_css
from src.ui.scripts import inject_js
from src.ui.header import render_header
from src.ui.sidebar import render_sidebar
from src.views.editing_stage import render_editing_stage
from src.views.new_weekend import render_new_weekend


# ── PAGE CONFIG ──
st.set_page_config(page_title="SJA - Ski Jumping Analysis", layout="wide", initial_sidebar_state="collapsed")

# ── PATHS & DATA ──
path_basic, path_Data_FIS, path_Weekend, path_Data_PDF, _ = init_paths()
list_Data_FIS, _, _ = load_data_lists(path_basic, path_Data_FIS, path_Weekend)

# ── SESSION STATE ──
init_session_state()
# Refresh the configured destination even if an editing session was already
# open when the application code was reloaded.
st.session_state.path_Data_PDF = path_Data_PDF

# ── CSS & JS INJECTION ──
inject_css()
inject_js()

# ── HEADER ──
render_header()

# ── SIDEBAR + WORKSPACE ──
col_workspace = render_sidebar(path_basic)

# ── WORKSPACE CONTENT ──
with col_workspace:
    st.markdown('<div class="workspace-marker"></div>', unsafe_allow_html=True)

    if st.session_state.get("show_editing_stage", False):
        render_editing_stage()
    else:
        render_new_weekend(
            list_Data_FIS,
            path_basic,
            path_Data_FIS,
            path_Weekend,
            path_Data_PDF,
        )
