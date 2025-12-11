# app/ui/tab_upload.py
import streamlit as st
import pandas as pd
from app.core.validator import Validator


def render_upload_tab(scheduler, writer, analyzer, config):
    st.subheader("📤 Upload & Proses Jadwal")

    uploaded = st.file_uploader("Upload file Excel (.xlsx)", type=["xlsx"])

    if not uploaded:
        st.info("Silakan upload file Excel yang memiliki sheet 'Reguler' dan 'Poleks'.")
        return

    # === FIX: gunakan Validator.validate() ===
    ok, err = Validator.validate(uploaded)
    if not ok:
        st.error(f"❌ File tidak valid: {err}")
        return

    # Preview sheet Reguler
    try:
        xl = pd.ExcelFile(uploaded)
        st.success(f"File valid. Sheet ditemukan: {', '.join(xl.sheet_names)}")

        if "Reguler" in xl.sheet_names and st.checkbox("Tampilkan preview 'Reguler' (10 baris)"):
            st.dataframe(
                pd.read_excel(uploaded, sheet_name="Reguler", nrows=10),
                use_container_width=True,
            )

    except Exception as e:
        st.error(f"Gagal membaca Excel: {e}")
        return

    # Tombol proses
    if st.button("🚀 Proses Jadwal"):
        with st.spinner("Memproses jadwal..."):

            try:
                df_reg = xl.parse("Reguler") if "Reguler" in xl.sheet_names else pd.DataFrame()
                df_pol = xl.parse("Poleks") if "Poleks" in xl.sheet_names else pd.DataFrame()
            except Exception as e:
                st.error(f"Gagal membaca sheet: {e}")
                return

            # === Proses jadwal ===
            df_r = scheduler.process_schedule(df_reg, "Reguler") if not df_reg.empty else pd.DataFrame()
            df_e = scheduler.process_schedule(df_pol, "Poleks") if not df_pol.empty else pd.DataFrame()

            df_all = pd.concat([df_r, df_e], ignore_index=True) \
                if not df_r.empty or not df_e.empty \
                else pd.DataFrame()

            if df_all.empty:
                st.warning("Hasil kosong. Tidak ada jadwal yang dapat diproses.")
                return

            # Simpan
            st.session_state["processed_data"] = df_all
            st.session_state["time_slots"] = [c for c in df_all.columns if ":" in c]

            st.success("✅ Jadwal berhasil diproses!")
            st.dataframe(df_all, use_container_width=True)

            # Export Excel
            buf = writer.write(uploaded, df_all, st.session_state["time_slots"])
            st.download_button(
                "📥 Download Jadwal Hasil",
                data=buf,
                file_name="jadwal_modular.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
