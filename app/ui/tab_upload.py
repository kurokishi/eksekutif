import streamlit as st
import pandas as pd
from app.core.validator import Validator

def render_upload_tab(scheduler, writer, analyzer, config):

    st.subheader("📤 Upload Jadwal")
    st.info("Silakan upload file Excel berformat Reguler & Poleks.")

    # ================= TEMPLATE DOWNLOAD =================
    st.subheader("📄 Download Template Excel")
    if st.button("📥 Download Template Jadwal"):
        template_buf = writer.generate_template(config.slot_times)
        st.download_button(
            label="Klik untuk Download Template",
            data=template_buf,
            file_name="template_jadwal.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.write("---")

    # ================== FILE UPLOADER =====================
    uploaded = st.file_uploader("Upload file Excel (.xlsx)", type=["xlsx"])

    if not uploaded:
        return

    ok, err = Validator.validate(uploaded)
    if not ok:
        st.error(f"❌ File tidak valid: {err}")
        return

    xl = pd.ExcelFile(uploaded)
    st.success(f"File valid. Sheets: {xl.sheet_names}")

    if "Reguler" in xl.sheet_names and st.checkbox("Preview sheet Reguler"):
        st.dataframe(pd.read_excel(uploaded, sheet_name="Reguler", nrows=10))

    # ================== PROSES ============================
    if st.button("🚀 Proses Jadwal"):

        df_reg = xl.parse("Reguler") if "Reguler" in xl.sheet_names else pd.DataFrame()
        df_pol = xl.parse("Poleks") if "Poleks" in xl.sheet_names else pd.DataFrame()

        df_r = scheduler.process_schedule(df_reg, "Reguler")
        df_e = scheduler.process_schedule(df_pol, "Poleks")

        df_all = pd.concat([df_r, df_e], ignore_index=True)

        st.session_state["processed_data"] = df_all
        st.session_state["time_slots"] = config.slot_times

        st.success("Jadwal berhasil diproses!")
        st.dataframe(df_all)

        # SAVE
        buf = writer.write(uploaded, df_all, config.slot_times)
        st.download_button(
            "📥 Download Jadwal Hasil",
            data=buf,
            file_name="jadwal_hasil.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
