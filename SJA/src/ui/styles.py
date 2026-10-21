"""
All CSS injection for the SJA Streamlit app.
"""
import streamlit as st


def inject_css():
    """Injects all custom CSS styles into the Streamlit app."""
    st.markdown("""
<style>
/* ── Schriftart & Hintergrund ────────────────────────────────────────── */

html, body, .stApp, .stMarkdown, p, span, label, select, button {
    font-family: 'Century Gothic', 'CenturyGothic', sans-serif !important;
    color: #FFFFFF !important;
}

.stApp {
    background-color: #262626 !important;
}

/* ── HIDE STREAMLIT ELEMENTS ────────────────────────────────────────── */
header { visibility: hidden !important; }
footer { visibility: hidden !important; }
#MainMenu { visibility: hidden !important; }
.stDeployButton { display: none !important; }
div[data-testid="stDecoration"] { display: none !important; }

/* Remove default padding from block container to make layout fill screen */
div.block-container {
    padding-top: 0rem !important;
    padding-bottom: 0rem !important;
    padding-left: 0rem !important;
    padding-right: 0rem !important;
    max-width: 100% !important;
}

/* ── ALIGN CAPTIONS LEFT ────────────────────────────────────────────── */
div[data-testid="stCaptionContainer"], [data-testid="stCaptionContainer"] p {
    text-align: left !important;
    justify-content: flex-start !important;
}

/* ── HEADER LAYOUT ───────────────────────────────────────────────────── */
.header-container {
    display: flex;
    justify-content: space-between;
    align-items: center;
    background-color: #262626;
    padding: 10px 20px;
    height: 60px;
    border-bottom: 1px solid #FFFFFF !important;
}

/* Ensure the page header has a constant height in all views */
div[data-testid="stHorizontalBlock"]:has(.header-left) {
    min-height: 30px !important;
}
div[data-testid="stHorizontalBlock"]:has(.header-left) > div {
    display: flex !important;
    align-items: center !important;
    min-height: 40px !important;
}

.header-left, .header-left * {
    font-family: 'Century Gothic', 'CenturyGothic', sans-serif !important;
}

.header-left {
    font-size: 1.15rem;
    letter-spacing: 1px;
    margin-left: 10px;
}

.header-center {
    font-size: 1.25rem;
    font-weight: bold !important;
    text-align: center;
}
.header-center * {
    font-weight: bold !important;
}

.header-right {
    display: flex;
    align-items: center;
    justify-content: flex-end;
    gap: 15px;
}

.header-divider {
    margin: 0 !important;
    border: 0;
    height: 1px;
    background-color: #FFFFFF !important;
}

/* Gender-spezifische Farben */
.gender-M { color: #87CEEB !important; } /* Hellblau */
.gender-W { color: #FFB6C1 !important; } /* Hellrosa */
.gender-A { color: #EE82EE !important; } /* Hellviolett */

/* ── SIDEBAR CONTAINER & TAB STYLING ────────────────────────────────── */
/* ── BUTTON STYLING (GLOBAL & WORKSPACE) ────────────────────────────────── */
button {
    background-color: #262626 !important;
    color: #FFFFFF !important;
    border: none !important;
    border-radius: 4px !important;
    padding: 6px 16px !important;
    transition: background-color 0.2s, border-color 0.2s !important;
}
button:hover {
    background-color: #4A4A4A !important;
    border: none !important;
}
button p {
    color: #FFFFFF !important;
}

div[data-testid*="olumn"]:has(.sidebar-marker) {
    background-color: #262626 !important;
    border-right: 1px solid #FFFFFF !important;
    padding: 15px 10px 0px 10px !important;
    min-height: calc(100vh - 61px) !important;
    flex: 0 0 auto !important;
    width: max-content !important;
    max-width: none !important;
    min-width: 140px !important;
    
    /* Make the column container use flexbox */
    display: flex !important;
    flex-direction: column !important;
}

/* Make the vertical block container of the sidebar stretch to fill height and use flexbox */
div[data-testid*="olumn"]:has(.sidebar-marker) div[data-testid="stVerticalBlock"] {
    flex-grow: 1 !important;
    display: flex !important;
    flex-direction: column !important;
    width: max-content !important;
    min-width: 100% !important;
    gap: 10px !important;
}

/* Ensure intermediate button wrappers also expand */
div[data-testid*="olumn"]:has(.sidebar-marker) div[data-testid="stVerticalBlock"] > div,
div[data-testid*="olumn"]:has(.sidebar-marker) div.stButton {
    width: max-content !important;
    min-width: 100% !important;
}

/* Push the sidebar spacer and subsequent children (like Settings) to the bottom */
div[data-testid*="olumn"]:has(.sidebar-marker) div[data-testid="stVerticalBlock"] div:has(.sidebar-spacer) {
    margin-top: auto !important;
}

/* Custom styling for workspace_marker container to hide it */
div[data-testid="element-container"]:has(.workspace-marker) {
    display: none !important;
}

/* Workspace-Spalte (zweite Spalte unterhalb der Trennlinie) füllt den Rest aus */
div[data-testid*="olumn"]:has(.workspace-marker) {
    flex: 1 1 0% !important;
    width: 100% !important;
    min-width: 0 !important;
    overflow-x: hidden !important;
    padding: 10px 30px 30px 30px !important;
}

/* Button-Styling in der Sidebar */
div[data-testid*="olumn"]:has(.sidebar-marker) button {
    text-align: left !important;
    justify-content: flex-start !important;
    font-size: 0.78rem !important;
    font-family: 'Century Gothic', 'CenturyGothic', sans-serif !important;
    border: none !important;
    border-radius: 4px !important;
    width: 100% !important;
    min-width: max-content !important;
    padding: 8px 12px !important;
    transition: background-color 0.2s, color 0.2s !important;
    white-space: nowrap !important;
    height: 39px !important;
    min-height: 39px !important;
    max-height: 39px !important;
    
    /* Default inactive styles */
    background-color: #262626 !important;
    color: #CCCCCC !important;
}

/* Text-Tag im Button linksbündig ausrichten und Umbruch verhindern */
div[data-testid*="olumn"]:has(.sidebar-marker) button p {
    text-align: left !important;
    justify-content: flex-start !important;
    width: 100% !important;
    margin: 0 !important;
    display: flex !important;
    white-space: nowrap !important;
    color: #CCCCCC !important;
}

div[data-testid*="olumn"]:has(.sidebar-marker) button * {
    white-space: nowrap !important;
}


/* Hover-Effekte für alle Buttons in der Sidebar */
div[data-testid*="olumn"]:has(.sidebar-marker) button:hover {
    background-color: #4A4A4A !important;
    color: #FFFFFF !important;
    border: none !important;
}

div[data-testid*="olumn"]:has(.sidebar-marker) button:hover p {
    color: #FFFFFF !important;
}

/* Spezifische Styles für den aktiven Tab (Primary Button) */
div[data-testid*="olumn"]:has(.sidebar-marker) button[data-testid*="primary"],
div[data-testid*="olumn"]:has(.sidebar-marker) button[class*="primary"] {
    background-color: #4A4A4A !important;
    border: none !important;
    color: #FFFFFF !important;
}

div[data-testid*="olumn"]:has(.sidebar-marker) button[data-testid*="primary"] p,
div[data-testid*="olumn"]:has(.sidebar-marker) button[class*="primary"] p {
    color: #FFFFFF !important;
    font-weight: bold !important;
}


/* ── WORKSPACE LAYOUT ────────────────────────────────────────────────── */
.workspace-container {
    padding: 30px;
    min-height: calc(100vh - 61px);
    background-color: #262626;
}

/* ── SELECTBOX STYLING (In der Header-Zeile) ─────────────────────────── */
div[data-testid="stSelectbox"] > div > div {
    background-color: #262626 !important;
    border-color: #444444 !important;
    color: #FFFFFF !important;
    font-size: 0.8rem !important;
    border-radius: 4px !important;
}
div[data-testid="stSelectbox"] svg {
    color: #FFFFFF !important;
}

/* Custom styling for the raw HTML tables and dataframes */
div[data-testid="stDataFrame"] {
    border: 1px solid #333333;
    border-radius: 6px;
}

/* ── Compact mode for data_editor grids ──────────────────────────── */
div[data-testid="stDataFrame"] [data-testid="glideDataEditor"] {
    font-size: 12px !important;
}

/* Smaller header row */
div[data-testid="stDataFrame"] .dvn-scroller .header-row {
    height: 26px !important;
    line-height: 26px !important;
}



/* ── Toast Styling ────────────────────────────────────────────────── */
div[data-testid="stToast"] {
    background-color: #262626 !important;
    border: 1px solid #FFFFFF !important;
}

div[data-testid="stToast"] p,
div[data-testid="stToast"] span,
div[data-testid="stToast"] div {
    color: #FFFFFF !important;
    font-size: 0.78rem !important;
}

div[data-testid="stToast"] button svg {
    fill: #FFFFFF !important;
    color: #FFFFFF !important;
}
</style>
""", unsafe_allow_html=True)
