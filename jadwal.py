# jadwal.py
import streamlit as st
import pandas as pd

# Modular imports
from app.config import Config
from app.core.scheduler import Scheduler
from app.core.excel_writer import ExcelWriter
from app.core.analyzer import ErrorAnalyzer
from app.ui.sidebar import render_sidebar
from app.ui.tab_upload import render_upload_tab
from app.ui.tab_analyzer import render_analyzer_tab
from app.ui.tab_visualization import render_visualization_tab
from app.ui.tab_settings import render_settings_tab


def main():
    st.set_page_config(
        page_title="🏥 Pengisi Jadwal Poli Modular",
        layout="wide"
    )

    st.title("🏥 Pengisi Jadwal Poli — Struktur Modular")

    # Init config in session
    if "config" not in st.session_state:
        st.session_state.config = Config()

    config = st.session_state.config

    # Sidebar = UI global
    render_sidebar(config)

    # Instantiate core services once
    scheduler = Scheduler(config)
    writer = ExcelWriter(config)
    analyzer = ErrorAnalyzer()

    # Tabs
    tab1, tab2, tab3, tab4 = st.tabs([
        "📤 Upload & Proses",
        "🔍 Error Analyzer",
        "📊 Visualisasi",
        "⚙️ Pengaturan"
    ])

    # Tab: Upload & Proses
    with tab1:
        render_upload_tab(
            scheduler=scheduler,
            writer=writer,
            analyzer=analyzer,
            config=config
        )

    # Tab: Error Analyzer
    with tab2:
        render_analyzer_tab(
            analyzer=analyzer,
            config=config
        )

    # Tab: Visualisasi jadwal
    with tab3:
        render_visualization_tab(
            config=config
        )

    # Tab: Settings
    with tab4:
        render_settings_tab(
            config=config
        )


if __name__ == "__main__":
    main()
