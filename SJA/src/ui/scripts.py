"""
JavaScript injection for chart fullscreen toggle, multiselect hide, and debug overlay.
"""
import streamlit as st


def inject_js():
    """Injects all custom JavaScript into the Streamlit app."""
    st.components.v1.html(
        """
    <script>
    const parentDoc = window.parent.document;
    
    // Cleanup: remove leftover debug overlay from previous sessions
    const oldOverlay = parentDoc.getElementById('debug-dropdown-overlay');
    if (oldOverlay) { oldOverlay.remove(); }
    
    // CSS-Regel für maximierte Diagramme injizieren
    if (!parentDoc.getElementById('fullscreen-chart-style')) {
        const style = parentDoc.createElement('style');
        style.id = 'fullscreen-chart-style';
        style.innerHTML = `
            .fullscreen-chart-active {
                position: fixed !important;
                top: 0 !important;
                left: 0 !important;
                width: 100vw !important;
                height: 100vh !important;
                z-index: 999999 !important;
                background-color: #1E1E1E !important;
                padding: 30px !important;
                box-sizing: border-box !important;
            }
        `;
        parentDoc.head.appendChild(style);
    }
    
    // Event-Listener im Parent-Dokument für Doppelklicks auf Plotly-Charts registrieren
    if (!parentDoc.fullscreenListenerAdded) {
        parentDoc.addEventListener('dblclick', function(e) {
            let target = e.target;
            while (target && target !== parentDoc.body) {
                // Prüfe, ob es sich um ein Plotly-Chart oder ein stPlotlyChart-Container handelt
                if (target.classList && (target.classList.contains('js-plotly-plot') || target.getAttribute('data-testid') === 'stPlotlyChart')) {
                    target.classList.toggle('fullscreen-chart-active');
                    // Fenster-Größenänderung triggern, damit sich Plotly anpasst
                    window.parent.dispatchEvent(new Event('resize'));
                    break;
                }
                target = target.parentNode;
            }
        }, true); // capturing phase = true
        parentDoc.fullscreenListenerAdded = true;
    }

    // Hide "Select all" option in multiselect dropdowns
    if (!parentDoc.selectAllObserverAdded) {
        const hideSelectAll = () => {
            const options = parentDoc.querySelectorAll('li[role="option"]');
            options.forEach(option => {
                if (option.textContent && option.textContent.trim() === "Select all" && option.style.display !== "none") {
                    option.style.display = "none";
                    option.style.position = "absolute";
                    option.style.height = "0";
                    option.style.width = "0";
                    option.style.margin = "0";
                    option.style.padding = "0";
                    option.style.border = "none";
                    option.style.visibility = "hidden";
                }
            });
        };
        hideSelectAll();
        const observer = new MutationObserver((mutations) => {
            hideSelectAll();
        });
        observer.observe(parentDoc.body, { childList: true, subtree: true });
        parentDoc.selectAllObserverAdded = true;
    }

    </script>
    """,
        height=0,
        width=0
    )
