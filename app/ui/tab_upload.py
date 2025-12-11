# app/ui/tab_upload.py
import streamlit as st
import pandas as pd
from app.core.validator import Validator

def render_upload_tab(scheduler, writer, analyzer, config):
    st.subheader("📤 Upload & Proses")

    uploaded = st.file_uploader("Upload file Excel (.xlsx)", type=['xlsx'])
    if not uploaded:
        st.info("Upload file Excel yang memiliki sheet 'Reguler' dan 'Poleks'.")
        return

    ok, err = Validator.validate_file(uploaded)
    if not ok:
        st.error(f"File tidak valid: {err}")
        return

    # preview
    try:
        xls = pd.ExcelFile(uploaded)
        st.success(f"{len(xls.sheet_names)} sheet ditemukan: {', '.join(xls.sheet_names)}")
        if st.checkbox("Tampilkan preview sheet 'Reguler' (10 baris)"):
            try:
                st.dataframe(pd.read_excel(uploaded, sheet_name='Reguler', nrows=10), use_container_width=True)
            except Exception as e:
                st.warning(f"Gagal preview: {e}")
    except Exception as e:
        st.error(f"Gagal membaca file: {e}")
        return

    if st.button("🚀 Proses Jadwal"):
        with st.spinner("Memproses..."):
            try:
                df_reg = pd.read_excel(uploaded, sheet_name='Reguler')
            except Exception:
                df_reg = pd.DataFrame()
            try:
                df_pol = pd.read_excel(uploaded, sheet_name='Poleks')
            except Exception:
                df_pol = pd.DataFrame()

            df_r = scheduler.process_schedule(df_reg, 'Reguler') if not df_reg.empty else pd.DataFrame()
            df_e = scheduler.process_schedule(df_pol, 'Poleks') if not df_pol.empty else pd.DataFrame()

            df_all = pd.concat([df_r, df_e], ignore_index=True) if not df_r.empty or not df_e.empty else pd.DataFrame()

            if df_all.empty:
                st.warning("Tidak ada data jadwal yang dihasilkan.")
                return

            # simpan ke session
            st.session_state['processed_data'] = df_all
            st.session_state['time_slots'] = [c for c in df_all.columns if ':' in c]

            st.success("✅ Selesai memproses.")
            st.dataframe(df_all, use_container_width=True)

            # create excel
            buf = writer.write(uploaded, df_all, st.session_state['time_slots'])
            st.download_button(
                "📥 Download Hasil (Excel)",
                data=buf,
                file_name="jadwal_hasil_modular.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
