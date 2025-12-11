import streamlit as st
import pandas as pd

from app.config import Config
from app.core.scheduler import Scheduler
from app.core.excel_writer import ExcelWriter
from app.ui.sidebar import render_sidebar

def main():
    st.set_page_config(page_title="Jadwal Poli Modular", layout="wide")

    config = Config()

    # Sidebar
    render_sidebar(config)

    st.title("🏥 Pengisi Jadwal Poli Modular")

    uploaded = st.file_uploader("Upload File Excel (.xlsx)", type=["xlsx"])

    if uploaded:
        excel = pd.ExcelFile(uploaded)

        df_reg = excel.parse("Reguler")
        df_pol = excel.parse("Poleks")

        sched = Scheduler(config)

        df_r = sched.process_one(df_reg, "Reguler")
        df_e = sched.process_one(df_pol, "Poleks")

        df_all = pd.concat([df_r, df_e]).reset_index(drop=True)

        st.subheader("📋 Hasil Jadwal")
        st.dataframe(df_all, use_container_width=True)

        slot_str = [c for c in df_all.columns if ":" in c]

        writer = ExcelWriter(config)
        buf = writer.write(uploaded, df_all, slot_str)

        st.download_button(
            "📥 Download Jadwal",
            data=buf,
            file_name="jadwal_modular.xlsx"
        )

if __name__ == "__main__":
    main()
