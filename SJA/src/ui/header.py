"""
Header rendering for the SJA Streamlit app.
"""
import streamlit as st

from src.utils.parsers import parse_weekend_filename


def render_header():
    """Renders the application header for the New Weekend workflow."""
    gender_class = "gender-A"
    header_center_text = st.session_state.get("filename_weekend", "")
    if header_center_text:
        parsed_info = parse_weekend_filename(header_center_text)
        if parsed_info:
            gender = parsed_info["weekend_gender"]
            if gender == "M":
                gender_class = "gender-M"
            elif gender == "W":
                gender_class = "gender-W"
            elif gender == "A":
                gender_class = "gender-A"

    # Verwende st.columns zur Aufteilung des Headers in 3 Teile (Links, Mitte, Rechts)
    col_left, col_center, col_right = st.columns([3.5, 5, 3.5], vertical_alignment="center")

    with col_left:
        st.markdown(
            '<div class="header-left"><b>SJA</b> - Ski Jumping <b>Analysis</b></div>',
            unsafe_allow_html=True
        )

    with col_center:
        if header_center_text:
            st.markdown(
                f'<div class="header-center"><span class="{gender_class}"><b>{header_center_text}</b></span></div>',
                unsafe_allow_html=True
            )

    with col_right:
        st.empty()

    # Durchgehende Trennlinie unter dem Header
    st.markdown('<hr class="header-divider">', unsafe_allow_html=True)
