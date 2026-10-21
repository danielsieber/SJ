"""
Session state initialization for the Streamlit app.
"""
import streamlit as st


def init_session_state():
    """Initializes all session state keys with their default values."""
    if "list_weekend_codex" not in st.session_state:
        st.session_state.list_weekend_codex = []

