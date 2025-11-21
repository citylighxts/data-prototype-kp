import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from datetime import datetime, time, timedelta
import os

# Konfigurasi Halaman
st.set_page_config(page_title="SLA Analytics Dashboard", layout="wide")

def run():
    st.title("📊 SLA Analytics & Handling Dashboard")
    st.markdown("""
    **Fitur:** Auto-Fill Data Kosong, Pembersihan Spasi, **SLA Achievement %**, dan Tabel Kompak.
    """)
    st.markdown("---")
    
    # --- DEFINISI LOKASI REGIONAL 3 ---
    regional_3_locations = [
        "P. Lembar", "Regional 3", "P. Batulicin", "R. Jawa", "Terminal Celukan Bawang",
        "Sub Regional BBN", "P. Tg. Emas", "P. Bumiharjo", "Tanjung Perak", "R. Bali Nusra",
        "P. Badas", "TANJUNGPERAK", "TANJUNGEMAS/KEUANGAN", "TANJUNGEMAS", "P. Tg. Intan",
        "BANJARMASIN/TPK", "KOTABARU/MEKARPUTIH", "P. Waingapu", "R. Kalimantan", "Terminal Nilam",
        "Terminal Kumai", "P. Kalimas", "P. Tg. Wangi", "P. Gresik", "P. Kotabaru",
        "BANJARMASIN/KOMERSIAL", "TANJUNGWANGI/TEKNIK", "Sub Regional Kalimantan", "GRESIK/TERMINAL",
        "Terminal Kota Baru", "P. Sampit", "BANJARMASIN/TMP", "P. Bagendang", "BANJARMASIN/PDS",
        "TENAU/KALABAHI", "P. Bima", "P. Tenau Kupang", "Terminal Lembar", "P. Tegal",
        "Terminal Trisakti", "BENOA/OPKOM", "P. Benoa", "BANJARMASIN/TEKNIK", "BANJARMASIN/PBJ",
        "TANJUNGINTAN", "KOTABARU", "TENAU", "Sub Regional Jawa Timur", "KUMAI/OPKOM",
        "Terminal Batulicin", "Terminal Gresik", "KUMAI/KEUPER", "LEMBAR/KEUPER", "P. Kalabahi",
        "BIMA/BADAS", "Terminal Jamrud", "TENAU/WAINGAPU", "Terminal Benoa", "P. Tg. Tembaga",
        "BIMA/PDS", "BENOA/SUK", "P. Clk. Bawang", "KUMAI/BUMIHARJO", "P. Pulang Pisau",
        "Terminal Labuan Bajo", "P. Maumere", "BENOA/KEUANGAN", "BENOA/PKWT", "Terminal Kalimas",
        "BANJARMASIN/KEUANGAN", "BENOA/PEMAGANG", "GRESIK/KEUANGAN", "Terminal Petikemas Banjarmasin",
        "CELUKANBAWANG", "P. Ende-Ippi", "SAMPIT/BAGENDANG", "Terminal Bima", "KOTABARU/KEPANDUAN",
        "Terminal Sampit", "Terminal Kupang", "BENOA/TEKNIK", "Terminal Maumere", "PROBOLINGGO/PLS",
        "SAMPIT/PKWT", "P. Labuan Bajo", "P. Kalianget", "Banjarmasin", "Terminal Waingapu", "MAUMERE/ENDE"
    ]

    st.sidebar.header("📂 Upload File")
    uploaded_req = st.sidebar.file_uploader("1. File Request Item (.xlsx)", type=["xlsx"])
    
    # --- LOGIKA OTOMATIS BACA FILE MAPPING ---
    default_sla_path = os.path.join(os.path.dirname(__file__), "data_sc_req_mapping.xlsx")

    if os.path.exists(default_sla_path):
        uploaded_sla = default_sla_path
    else:
        st.sidebar.warning("⚠️ File mapping default tidak ditemukan.")
        uploaded_sla = st.sidebar.file_uploader("2. Upload File Mapping SLA (.xlsx)", type=["xlsx"])

    # --- FUNGSI BANTUAN ---
    def parse_sla_duration(val):
        if pd.isna(val) or val == "": return pd.Timedelta(0)
        if isinstance(val, time): return timedelta(hours=val.hour, minutes=val.minute, seconds=val.second)
        if isinstance(val, (int, float)): return timedelta(days=val) 
        if isinstance(val, str):
            try: return timedelta(hours=pd.to_datetime(val, format='%H:%M:%S').hour, minutes=pd.to_datetime(val, format='%H:%M:%S').minute)
            except: pass
        return pd.Timedelta(0)

    def timedelta_to_excel_float(td):
        if pd.isna(td) or td == pd.Timedelta(0): return None
        return td.total_seconds() / 86400.0

    def clean_string_col(series):
        return series.fillna('').astype(str).str.strip().replace('nan', '')
    
    def remove_all_spaces(series):
        return series.astype(str).str.replace(" ", "", regex=False)

    def find_col(df, keywords):
        for col in df.columns:
            for k in keywords:
                if k.lower() in col.lower().replace(" ", ""):
                    return col
        return None

    if uploaded_req is not None:
        try:
            df_req = pd.read_excel(uploaded_req)
            
            # --- 1. DETEKSI KOLOM ---
            col_loc = find_col(df_req, ["Lokasi", "Location"]) or "Lokasi Pelapor"
            col_judul = find_col(df_req, ["Judul", "Short"]) or "Judul Permasalahan"
            col_bc = find_col(df_req, ["Business", "Criticality"]) or "Businesscriticality"
            col_sev = find_col(df_req, ["Severity"]) or "Severity"
            col_dibuat = find_col(df_req, ["Dibuat", "Created"]) or "Tiket Dibuat"
            col_ditutup = find_col(df_req, ["Ditutup", "Closed"]) or "Tiket Ditutup"
            col_target_asli = find_col(df_req, ["TargetSelesai", "Due"]) or "Target Selesai"
            col_contact = find_col(df_req, ["Contact", "Type"]) or "Contact type"
            col_item = find_col(df_req, ["Item"]) or "Item"
            col_service = find_col(df_req, ["Service", "Offering"]) or "Service offering"

            # Bersihkan String Dasar
            if col_loc in df_req.columns: df_req[col_loc] = clean_string_col(df_req[col_loc])
            if col_judul in df_req.columns: df_req[col_judul] = clean_string_col(df_req[col_judul])
            
            # --- AUTO FILL DEFAULT ---
            if col_bc in df_req.columns:
                df_req[col_bc] = clean_string_col(df_req[col_bc])
                df_req[col_bc] = df_req[col_bc].replace(['', 'nan', 'None'], '3-Medium')
            
            if col_sev in df_req.columns:
                df_req[col_sev] = clean_string_col(df_req[col_sev])
                df_req[col_sev] = df_req[col_sev].replace(['', 'nan', 'None'], '3-Low')

            # --- 2. FILTER REGIONAL ---
            if col_loc in df_req.columns:
                filter_option = st.radio("Pilih Data:", ("All Data", "Regional 3 Only"), horizontal=True)
                if filter_option == "Regional 3 Only":
                    df_main = df_req[df_req[col_loc].isin(regional_3_locations)].copy()
                else:
                    df_main = df_req.copy()
                
                df_main['Data Reg3'] = df_main[col_loc].apply(lambda x: 'Regional 3' if x in regional_3_locations else 'Non-Reg3')
            else:
                st.error("Kolom Lokasi Pelapor tidak ditemukan.")
                st.stop()

            # --- 3. MAPPING SLA ---
            if uploaded_sla is not None:
                try:
                    map_item = pd.read_excel(uploaded_sla, sheet_name='Map_Item')
                    map_sev = pd.read_excel(uploaded_sla, sheet_name='Map_Severity')
                    map_dur = pd.read_excel(uploaded_sla, sheet_name='Map_Durasi')

                    # Cleaning Mapping
                    for m in [map_item, map_sev, map_dur]:
                        m.columns = m.columns.str.strip()

                    map_item['Judul Permasalahan'] = clean_string_col(map_item['Judul Permasalahan'])
                    sev_map_col = find_col(map_sev, ["BusinessCritical", "Severity"])
                    map_sev['Clean_Key_Map'] = remove_all_spaces(map_sev[sev_map_col])
                    map_dur['ID SLA'] = clean_string_col(map_dur['ID SLA'])

                    # --- LOGIKA GABUNG ---
                    df_main['Key_Clean_Req'] = df_main.apply(
                        lambda x: (x[col_bc].replace(" ", "") + x[col_sev].replace(" ", "")), axis=1
                    )
                    df_main['Businesscriticality-Severity'] = df_main[col_bc] + df_main[col_sev]

                    # Merge
                    df_merged = pd.merge(df_main, map_item[['Judul Permasalahan', 'ID']], left_on=col_judul, right_on='Judul Permasalahan', how='left')
                    df_merged.rename(columns={'ID': 'ID_Item'}, inplace=True)

                    df_merged = pd.merge(df_merged, map_sev[['Clean_Key_Map', 'ID']], left_on='Key_Clean_Req', right_on='Clean_Key_Map', how='left')
                    df_merged.rename(columns={'ID': 'ID_Sev'}, inplace=True)

                    # ID SLA Final
                    df_merged['ID_Item_Str'] = df_merged['ID_Item'].fillna('').astype(str).str.replace(r'\.0$', '', regex=True)
                    df_merged['ID_Sev_Str'] = df_merged['ID_Sev'].fillna('').astype(str).replace('nan', '')
                    
                    df_merged['ID SLA Final'] = df_merged.apply(
                        lambda x: x['ID_Item_Str'] + x['ID_Sev_Str'] if x['ID_Item_Str'] != "" and x['ID_Sev_Str'] != "" else None, axis=1
                    )

                    # Get Durasi
                    df_final = pd.merge(df_merged, map_dur[['ID SLA', 'SLA']], left_on='ID SLA Final', right_on='ID SLA', how='left')
                    df_final.rename(columns={'SLA': 'Target SLA Raw'}, inplace=True)

                    # --- 4. HITUNG & KONVERSI ---
                    df_final[col_dibuat] = pd.to_datetime(df_final[col_dibuat], errors='coerce')
                    if col_ditutup in df_final.columns:
                        df_final[col_ditutup] = pd.to_datetime(df_final[col_ditutup], errors='coerce')

                    df_final['SLA_Timedelta'] = df_final['Target SLA Raw'].apply(parse_sla_duration)
                    df_final['Target SLA'] = df_final['SLA_Timedelta'].apply(timedelta_to_excel_float)

                    df_final['Target Selesai Hitung'] = df_final.apply(
                        lambda row: row[col_dibuat] + row['SLA_Timedelta'] if pd.notnull(row['SLA_Timedelta']) and row['SLA_Timedelta'] != pd.Timedelta(0) else pd.NaT, axis=1
                    )

                    # Hitung Status SLA
                    def hitung_status(row):
                        if pd.isna(row['Target Selesai Hitung']): return ""
                        if col_ditutup not in row or pd.isna(row[col_ditutup]): return "WP"
                        return 1 if row[col_ditutup] <= row['Target Selesai Hitung'] else 0

                    df_final['SLA'] = df_final.apply(hitung_status, axis=1)

                    # --- PREPARE DATA FOR DOWNLOAD ---
                    if col_target_asli in df_final.columns:
                        df_final.rename(columns={col_target_asli: "Target Selesai (Due Date Asli)"}, inplace=True)
                    df_final.rename(columns={'Target Selesai Hitung': 'Target Selesai'}, inplace=True)

                    desired_columns = [
                        "No. Tiket", col_dibuat, "Disetujui", "Status", "Item", "Permintaan", 
                        "Requested for", "Target Selesai (Due Date Asli)", "Tahapan", 
                        "Dibuka Oleh", "Jumlah", "Name", "PIC", "Comments and Work notes", 
                        "Deskripsi Permasalahan", col_judul, "Komentar Tambahan", 
                        "Root Cause and Solution", "Service offering", col_loc, 
                        "Deskripsi Permasalahan", col_ditutup, col_bc, 
                        "Contact type", col_sev, "Data Reg3", 
                        "Businesscriticality-Severity", "Target SLA", "Target Selesai", "SLA"
                    ]
                    
                    seen = set()
                    final_cols = [x for x in desired_columns if not (x in seen or seen.add(x))]
                    available_cols = [c for c in final_cols if c in df_final.columns]
                    df_display = df_final[available_cols]

                    # ==========================================
                    # 📊 VISUALIZATION DASHBOARD SECTION
                    # ==========================================
                    
                    st.success("✅ Data berhasil diproses!")
                    
                    tab1, tab2 = st.tabs(["📊 Dashboard Visualisasi", "📄 Data Preview"])

                    with tab1:
                        # --- 1. SLA RECAP (ANGKA & PERCENTAGE) ---
                        st.subheader("1. SLA Performance Recap")
                        
                        # Hitung Count
                        sla_counts = df_final['SLA'].value_counts()
                        on_time = sla_counts.get(1, 0)
                        late = sla_counts.get(0, 0)
                        
                        # Hitung Total Tiket (Untuk Card Paling Kanan)
                        total_tickets = len(df_final)
                        
                        # Hitung Total Calculated (Untuk Persentase)
                        total_calculated = on_time + late
                        
                        # --- RUMUS EXCEL: =AE2064/AH2064 ---
                        if total_calculated > 0:
                            achievement_rate = (on_time / total_calculated) * 100
                        else:
                            achievement_rate = 0
                        
                        # Metric Cards
                        c1, c2, c3, c4 = st.columns(4)
                        c1.metric("🏆 Achievement Rate", f"{achievement_rate:.1f}%", "SLA Performance")
                        c2.metric("On Time (Achieved)", f"{on_time} Tiket", "Sesuai Target")
                        c3.metric("Late (Breached)", f"{late} Tiket", "-Terlambat", delta_color="inverse")
                        # UBAH DISINI: Menampilkan Total Tiket (Semua Data)
                        c4.metric("Total Tiket", f"{total_tickets} Tiket", "Total Data")

                        # Donut Chart SLA
                        if total_calculated > 0:
                            df_sla_chart = pd.DataFrame({
                                'Status': ['On Time', 'Late'],
                                'Jumlah': [on_time, late]
                            })
                            fig_sla = px.pie(df_sla_chart, values='Jumlah', names='Status', hole=0.4, 
                                             color='Status', color_discrete_map={'On Time':'#00CC96', 'Late':'#EF553B'},
                                             title="SLA Compliance Rate")
                            st.plotly_chart(fig_sla, use_container_width=True)
                        else:
                            st.info("Belum ada data SLA yang terhitung.")

                        st.markdown("---")

                        # --- 2. LAYOUT GRID BAWAH ---
                        col_left, col_right = st.columns(2)

                        with col_left:
                            # TOP 5 BUSINESS CRITICALITY - SEVERITY
                            st.subheader("2. Top 5 Business Criticality - Severity")
                            if 'Businesscriticality-Severity' in df_final.columns:
                                top_bc = df_final['Businesscriticality-Severity'].value_counts().head(5).reset_index()
                                top_bc.columns = ['Category', 'Count']
                                fig_bc = px.bar(top_bc, x='Category', y='Count', text='Count', color='Count',
                                                title="Most Frequent Severity Combinations")
                                st.plotly_chart(fig_bc, use_container_width=True)
                            
                            # TOP 5 REQUESTED ITEMS
                            st.subheader("3. Top 5 Most Requested Items")
                            if col_item in df_final.columns:
                                top_items = df_final[col_item].value_counts().head(5).reset_index()
                                top_items.columns = ['Item', 'Count']
                                fig_items = px.bar(top_items, x='Count', y='Item', orientation='h', text='Count',
                                                   title="Top 5 Items (Angka)", color='Count')
                                fig_items.update_layout(yaxis={'categoryorder':'total ascending'})
                                st.plotly_chart(fig_items, use_container_width=True)

                        with col_right:
                            # TOP 4 CONTACT TYPE
                            st.subheader("4. Top 4 Contact Type Analysis")
                            if col_contact in df_final.columns:
                                top_contact = df_final[col_contact].value_counts().head(4).reset_index()
                                top_contact.columns = ['Type', 'Count']
                                fig_contact = px.pie(top_contact, values='Count', names='Type', hole=0.4,
                                                     title="Channel Pelaporan Terbanyak")
                                st.plotly_chart(fig_contact, use_container_width=True)

                            # --- DISINI DULUNYA ADA SERVICE OFFERING (SUDAH DIHAPUS) ---

                    with tab2:
                        # --- DATA PREVIEW TABLE (SEMPIT/KOMPAK) ---
                        st.subheader("📄 Data Preview (Excel Format)")
                        # use_container_width=False agar tabel tidak melebar
                        st.dataframe(df_display, use_container_width=False)

                    # --- DOWNLOAD BUTTON ---
                    st.markdown("### 📥 Download Report")
                    csv = df_display.to_csv(index=False, sep=';', decimal=',').encode('utf-8')
                    st.download_button("Download Hasil (.csv)", csv, "SLA_Dashboard_Report.csv", "text/csv")

                except Exception as e:
                    st.error(f"Error Proses: {e}")
            else:
                # Tampilkan tabel awal juga dengan mode sempit
                st.dataframe(df_main, use_container_width=False)

        except Exception as e:
            st.error(f"Gagal Baca File: {e}")

if __name__ == "__main__":
    run()