# app/ui/sidebar.py

import streamlit as st

def render_sidebar(config):
    st.title("🏥 Pengisi Jadwal Poli")
    st.markdown("---")

    config.enable_sabtu = st.checkbox(
        "Aktifkan Jadwal Hari Sabtu",
        value=config.enable_sabtu
    )

    st.markdown("### Pengaturan Auto-fix (placeholder)")
