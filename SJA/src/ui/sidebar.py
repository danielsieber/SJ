"""Sidebar rendering for the New Weekend workflow."""

import os

import streamlit as st


def render_sidebar(path_basic):
    """Render the permanent New Weekend sidebar and return the workspace."""
    col_sidebar, col_workspace = st.columns([0.5, 9.5], gap="small")

    with col_sidebar:
        st.markdown('<div class="sidebar-marker"></div>', unsafe_allow_html=True)

        if st.button(
            "New Weekend",
            key="tab_New Weekend",
            type="primary",
            use_container_width=True,
        ) and st.session_state.get("show_editing_stage", False):
            from src.views.editing_stage import _cleanup_editing_state

            _cleanup_editing_state()
            st.rerun()

        if st.button(
            "Listen.xlsm",
            key="tab_Listen.xlsm",
            type="secondary",
            use_container_width=True,
        ):
            _open_listen_workbook(path_basic / "Listen.xlsm")

        st.markdown('<div style="height: 50px;"></div>', unsafe_allow_html=True)

    return col_workspace


def _open_listen_workbook(listen_path):
    """Open Listen.xlsm and bring its Excel window to the foreground."""
    if not listen_path.exists():
        return

    try:
        os.startfile(listen_path)

        import subprocess
        import time

        for _ in range(3):
            time.sleep(0.3)
            result = subprocess.run(
                [
                    "powershell",
                    "-Command",
                    "$ws = New-Object -ComObject wscript.shell; $ws.SendKeys('%'); $ws.AppActivate('Listen.xlsm') -or $ws.AppActivate('Excel'); $ws.SendKeys('%')",
                ],
                capture_output=True,
                text=True,
                creationflags=0x08000000,
            )
            if "True" in result.stdout:
                break
    except Exception:
        pass
