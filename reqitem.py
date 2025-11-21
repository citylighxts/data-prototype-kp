import streamlit as st
import pandas as pd
import plotly.express as px
import io
import re

# --- 1. FUNGSI BANTUAN (PENTING) ---
def normalize_text(text):
    """Membersihkan teks: lowercase, hapus spasi berlebih, ubah ke string"""
    if pd.isna(text):
        return ""
    text = str(text).strip().lower()
    return text

def clean_id_str(val):
    """
    CRITICAL FIX: Mengubah float 105.0 menjadi string '105'.
    Tanpa ini, mapping sering gagal (105.0 != 105).
    """
    if pd.isna(val): 
        return ""
    try:
        # Coba ubah ke float dulu, lalu int untuk hilangkan desimal, lalu string
        return str(int(float(val)))
    except:
        # Jika gagal (misal ada huruf), ambil string aslinya
        return str(val).strip()

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

def run():
    # --- 2. HEADER & UPLOAD ---
    st.markdown(
        """
        <h1 style="display: flex; align-items: center; gap: 10px;">
            <img src="https://cdnjs.cloudflare.com/ajax/libs/twemoji/14.0.2/svg/1f4ca.svg" 
                 width="40" height="40">
            Request Item (SLA Calculator)
        </h1>
        """,
        unsafe_allow_html=True
    )

    st.write("Upload File Data & Mapping SLA.")

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📂 1. Data Tiket")
        file_data = st.file_uploader("Upload Excel Data", type=["xlsx", "xls"], key="data_uploader")

    with col2:
        st.subheader("📂 2. Mapping SLA")
        file_mapping = st.file_uploader("Upload Excel Mapping", type=["xlsx", "xls"], key="map_uploader")

    if not file_data or not file_mapping:
        st.warning("⚠️ Harap upload kedua file untuk memproses.")
        return

    # --- 3. PROSES MAPPING (FILE 2) ---
    try:
        xls_map = pd.ExcelFile(file_mapping)
        if 'Mapping SLA' not in xls_map.sheet_names:
            st.error("❌ Sheet 'Mapping SLA' tidak ditemukan dalam file Excel Mapping!")
            return
        
        df_map = pd.read_excel(xls_map, sheet_name='Mapping SLA')

        # --- MAPPING ITEM (Kolom L & M / Index 11 & 12) ---
        valid_items = df_map.iloc[:, [11, 12]].dropna()
        problem_map = dict(zip(
            valid_items.iloc[:, 0].apply(normalize_text), 
            valid_items.iloc[:, 1].apply(clean_id_str)  # Pakai clean_id_str
        ))

        # --- MAPPING SEVERITY (Kolom O & P / Index 14 & 15) ---
        valid_sev = df_map.iloc[:, [14, 15]].dropna()
        severity_map = dict(zip(
            valid_sev.iloc[:, 0].apply(normalize_text), 
            valid_sev.iloc[:, 1].apply(clean_id_str)    # Pakai clean_id_str
        ))

        # --- MAPPING DURASI (Kolom C & I / Index 2 & 8) ---
        valid_sla = df_map.iloc[:, [2, 8]].dropna()
        sla_days_map = dict(zip(
            valid_sla.iloc[:, 0].apply(clean_id_str),   # Kunci utama mapping (ID Gabungan)
            valid_sla.iloc[:, 1]                        # Nilai Durasi (Hari)
        ))
        
        st.success(f"✅ Mapping Loaded: {len(problem_map)} Items, {len(sla_days_map)} Rules Durasi.")

    except Exception as e:
        st.error(f"Error membaca file Mapping: {e}")
        st.info("Pastikan format kolom file Mapping tidak berubah posisi.")
        return
    
    # --- 4. PROSES DATA TIKET (FILE 1) ---
    try:
        df = pd.read_excel(file_data)
    except Exception as e:
        st.error(f"Error membaca file Data Tiket: {e}")
        return
    
    # --- 5. FILTER REGIONAL ---
    reg3_raw = "P. Lembar,Regional 3,P. Batulicin,R. Jawa,Terminal Celukan Bawang,Sub Regional BBN,P. Tg. Emas,P. Bumiharjo,Tanjung Perak,R. Bali Nusra,P. Badas,TANJUNGPERAK,TANJUNGEMAS/KEUANGAN,TANJUNGEMAS,P. Tg. Intan, BANJARMASIN/TPK, KOTABARU/MEKARPUTIH,P. Waingapu, R. Kalimantan,Terminal Nilam, Terminal Kumai,P. Kalimas, P. Tg. Wangi,P. Gresik, P. Kotabaru,BANJARMASIN/KOMERSIAL,TANJUNGWANGI/TEKNIK,Sub Regional Kalimantan,GRESIK/TERMINAL,Terminal Kota Baru,P. Sampit,BANJARMASIN/TMP,P. Bagendang,BANJARMASIN/PDS,TENAU/KALABAHI,P. Bima,P. Tenau Kupang,Terminal Lembar,P. Tegal,Terminal Trisakti,BENOA/OPKOM,P. Benoa,BANJARMASIN/TEKNIK,BANJARMASIN/PBJ,TANJUNGINTAN,KOTABARU,TENAU,Sub Regional Jawa Timur,KUMAI/OPKOM,Terminal Batulicin,Terminal Gresik, KUMAI/KEUPER,LEMBAR/KEUPER,P. Kalabahi,BIMA/BADAS,Terminal Jamrud,TENAU/WAINGAPU,Terminal Benoa,P. Tg. Tembaga,BIMA/PDS,BENOA/SUK,P. Clk. Bawang,KUMAI/BUMIHARJO,P. Pulang Pisau,Terminal Labuan Bajo,P. Maumere,BENOA/KEUANGAN,BENOA/PKWT,Terminal Kalimas,BANJARMASIN/KEUANGAN,BENOA/PEMAGANG,GRESIK/KEUANGAN,Terminal Petikemas Banjarmasin,CELUKANBAWANG,P. Ende-Ippi,SAMPIT/BAGENDANG,Terminal Bima,KOTABARU/KEPANDUAN,Terminal Sampit,Terminal Kupang, BENOA/TEKNIK, Terminal Maumere, PROBOLINGGO/PLS, SAMPIT/PKWT, P. Labuan Bajo, P. Kalianget, Banjarmasin, Terminal Waingapu, MAUMERE/ENDE"
    list_reg3 = [x.strip() for x in reg3_raw.split(',') if x.strip()]
    
    loc_col = next((c for c in ['Lokasi Pelapor', 'Lokasi', 'Location'] if c in df.columns), None)
    
    if loc_col:
        st.markdown("### 🔍 Filter")
        opt = st.radio("Pilih Data:", ["All", "Regional 3 Only"], horizontal=True)
        if opt == "Regional 3 Only":
            df = df[df[loc_col].astype(str).str.strip().isin(list_reg3)].copy()
            df['Data Reg3'] = df[loc_col]
            st.info(f"Menampilkan {len(df)} data Regional 3.")
        else:
            df['Data Reg3'] = df[loc_col]
    else:
        df['Data Reg3'] = ""
    
    # --- 6. KALKULASI SLA (CORE LOGIC) ---
    # --- 6. KALKULASI SLA (REVISI LOGIKA MAPPING) ---
    req_cols = {
        'Tiket Dibuat': ['Tiket Dibuat', 'Created', 'created'],
        'Tiket Ditutup': ['Tiket Ditutup', 'Closed', 'closed'],
        'Judul Permasalahan': ['Judul Permasalahan', 'Item', 'item'],
        'Severity': ['Severity', 'severity'] # Kita fokus ke Severity saja
    }
    
    col_map = {}
    for key, candidates in req_cols.items():
        found = next((c for c in candidates if c in df.columns), None)
        if not found:
            st.error(f"❌ Kolom '{key}' tidak ditemukan di Excel Data Tiket!")
            return
        col_map[key] = found

    # A. Mapping ID Angka (Dari Judul Permasalahan) - Sudah Benar
    # Pakai clean_id_str agar 105.0 menjadi "105"
    df['ID_Angka'] = df[col_map['Judul Permasalahan']].apply(normalize_text).map(problem_map)
    
    # B. Mapping ID Huruf (FIX: DARI SEVERITY LANGSUNG)
    # Kita ambil kolom Severity, bersihkan, lalu map.
    def get_clean_severity(val):
        s = str(val).strip()
        if s.lower() == 'nan': return ""
        return s

    df['Clean_Severity'] = df[col_map['Severity']].apply(get_clean_severity)
    
    # Coba mapping langsung dari Severity
    df['ID_Huruf'] = df['Clean_Severity'].apply(normalize_text).map(severity_map)
    
    # NOTE: Jika Mapping masih gagal, sistem akan mencoba mencari dengan format "BusinessCrit - Severity"
    # Tapi berdasarkan gambar Anda, sepertinya Severity saja sudah cukup.

    # C. Gabung ID & Map Durasi
    # Pastikan ID Angka dan Huruf bersih sebelum digabung
    df['SLA_Code'] = df['ID_Angka'].apply(clean_id_str) + df['ID_Huruf'].apply(clean_id_str)
    
    # Replace kode yang kosong/rusak dengan NA
    df['SLA_Code'] = df['SLA_Code'].replace(['', 'nan', 'NaN', 'None'], pd.NA)
    
    # Mapping ke Durasi (Hari)
    df['Target SLA'] = df['SLA_Code'].map(sla_days_map) 

    # D. Hitung Tanggal Target
    df[col_map['Tiket Dibuat']] = pd.to_datetime(df[col_map['Tiket Dibuat']])
    df['Target SLA'] = pd.to_numeric(df['Target SLA'], errors='coerce')
    df['Target Selesai'] = df[col_map['Tiket Dibuat']] + pd.to_timedelta(df['Target SLA'], unit='D')

    # E. Status SLA Calculator
    df[col_map['Tiket Ditutup']] = pd.to_datetime(df[col_map['Tiket Ditutup']], errors='coerce')
    
    def calc_sla(row):
        closed = row[col_map['Tiket Ditutup']]
        target = row['Target Selesai']
        
        if pd.isna(closed): return "WP"       # Belum ditutup
        if pd.isna(target): return "Unknown"  # Gagal Mapping
        return 1 if closed <= target else 0   # 1=Achieved, 0=Late

    df['SLA'] = df.apply(calc_sla, axis=1)

    # --- 7. VISUALISASI (Gaya Kamu) ---
    st.markdown("---")
    st.subheader("✨ SLA Recap")

    sla_tercapai = int((df['SLA'] == 1).sum())
    sla_tidak_tercapai = int((df['SLA'] == 0).sum())
    sla_open = int((df['SLA'] == "WP").sum())
    sla_unknown = int((df['SLA'] == "Unknown").sum())
    total_semua = len(df)
    
    if sla_unknown > 0:
        st.warning(f"⚠️ PERINGATAN: Ada {sla_unknown} tiket berstatus 'Unknown' (Mapping Gagal).")
        with st.expander("🔍 Klik untuk Debugging (Lihat Data Gagal)"):
            # PERBAIKAN DI SINI: Hapus 'Businesscriticality-Severity', ganti 'Clean_Severity'
            cols_debug = [
                col_map['Judul Permasalahan'], 
                'Clean_Severity',  
                'ID_Angka', 
                'ID_Huruf', 
                'SLA_Code'
            ]
            # Pastikan kolom ada sebelum ditampilkan
            cols_debug = [c for c in cols_debug if c in df.columns]
            
            debug_df = df[df['SLA'] == 'Unknown'][cols_debug].head(20)
            st.dataframe(debug_df)
            st.write("Tips: Jika 'ID_Huruf' NaN, berarti Severity tidak cocok dengan Mapping.")

    # Metric Cards
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Tickets", total_semua)
    c2.metric("🏆 Achieved", sla_tercapai)
    c3.metric("🚨 Not Achieved", sla_tidak_tercapai)
    c4.metric("⏳ Open / WP", sla_open)

    # Bar Chart Recap
    rekap_data = pd.DataFrame({
        'Kategori': ['SLA Achieved', 'SLA Not Achieved', 'Open (WP)', 'Unknown'],
        'Jumlah': [sla_tercapai, sla_tidak_tercapai, sla_open, sla_unknown]
    })

    fig2 = px.bar(rekap_data, x='Kategori', y='Jumlah', text='Jumlah', title='SLA Status Distribution',
                  color='Kategori', 
                  color_discrete_map={'SLA Achieved':'#00CC96', 'SLA Not Achieved':'#EF553B', 'Open (WP)':'#636EFA', 'Unknown':'#AB63FA'})
    fig2.update_traces(textposition='outside')
    st.plotly_chart(fig2, use_container_width=True)

    # Donut Chart (Closed Tickets Only)
    if sla_tercapai + sla_tidak_tercapai > 0:
        donut_data = pd.DataFrame({
            "SLA": ["On Time", "Late"],
            "Jumlah": [sla_tercapai, sla_tidak_tercapai]
        })

        fig_donut = px.pie(donut_data, names="SLA", values="Jumlah", hole=0.4, title="On Time Percentage (Closed Tickets Only)",
                           color_discrete_sequence=['#00CC96', '#EF553B'])
        st.plotly_chart(fig_donut, use_container_width=True)

    # === Analisis tambahan ===
    st.subheader("✨ Additional Analysis")

    # Top 5 Kombinasi Severity
    if 'Clean_Severity' in df.columns:
        top5 = df['Clean_Severity'].value_counts().reset_index()
        top5.columns = ['Severity Level', 'Jumlah'] # Ubah nama kolom agar rapi
        top5.index = top5.index + 1
        top5 = top5.head(5)
        
        st.markdown("**Top Severity Frequency:**")
        st.markdown(top5.to_html(index=True, justify='left').replace('<td>', '<td style="text-align:left;">'), unsafe_allow_html=True)
    # === Contact Type & Item ===
    contact_col = next((c for c in ['Contact Type', 'ContactType', 'Contact type'] if c in df.columns), None)

    if contact_col:
        contact_summary = df[contact_col].value_counts(dropna=False).reset_index()
        contact_summary.columns = ['Tipe Kontak', 'Jumlah']
        contact_summary.index = contact_summary.index + 1
        top3_contact = contact_summary.head(3)

        st.subheader("✨ Contact Type Analysis")
        st.markdown("**Top 3 contact type analysis**")
        st.markdown(top3_contact.to_html(index=True, justify='left').replace('<td>', '<td style="text-align:left;">'), unsafe_allow_html=True)

        fig_contact = px.pie(top3_contact, names='Tipe Kontak', values='Jumlah', hole=0.4, title='Top 3 Contact Type')
        st.plotly_chart(fig_contact, use_container_width=True)
        
        # Top 5 Items
        if col_map['Judul Permasalahan']:
            top5_item = df[col_map['Judul Permasalahan']].value_counts(dropna=False).reset_index()
            top5_item.columns = ['Item', 'Jumlah']
            top5_item.index = top5_item.index + 1
            top5_item = top5_item.head(5)

            st.subheader("✨ Top 5 Most Requested Items")
            st.markdown(top5_item.to_html(index=True, justify='left').replace('<td>', '<td style="text-align:left;">'), unsafe_allow_html=True)

            fig_item = px.bar(
                top5_item, x='Item', y='Jumlah', text='Jumlah',
                title='Top 5 Most Requested Items',
            )
            fig_item.update_traces(textposition='outside')
            st.plotly_chart(fig_item, use_container_width=True)

        # === Analisis Service Offering ===
        service_col = next((c for c in ['Service Offering', 'ServiceOffering'] if c in df.columns), None)
        if service_col:
            st.subheader("✨ Service Offering Analysis")
            service_summary = (
                df[service_col].dropna().astype(str)
                .replace(['', 'None', 'nan', 'NaN'], pd.NA).dropna()
                .value_counts().reset_index()
            )
            service_summary.columns = ['Service Offering', 'Jumlah']
            service_summary.index = service_summary.index + 1
            top3_service = service_summary.head(3)

            st.markdown("**Top 3 Service Offerings:**")
            st.markdown(top3_service.to_html(index=True, justify='left').replace('<td>', '<td style="text-align:left;">'), unsafe_allow_html=True)

            fig_service = px.bar(top3_service, x='Service Offering', y='Jumlah', text='Jumlah', title='Top 3 Service Offering')
            fig_service.update_traces(textposition='outside')
            fig_service.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig_service, use_container_width=True)

    # --- 8. OUTPUT DATA ---
    st.subheader("🔥 Calculation Result")
    
    def get_status_label(val):
        if val == 1: return "Achieved"
        if val == 0: return "Not Achieved"
        if val == "WP": return "Open/WP"
        return "Unknown"

    df["Status SLA"] = df['SLA'].apply(get_status_label)

    # Menentukan kolom yang ditampilkan
    tiket_col = next((c for c in ['No. Tiket', 'Ticket No', 'No Ticket'] if c in df.columns), None)
    
    target_cols = [
        tiket_col, 
        col_map['Judul Permasalahan'],
        'Clean_Severity',
        'Target SLA',       # Ini nilai hari (misal 0.02)
        'Target Selesai',   # Ini tanggal
        'SLA',              # Ini 1/0
        'Status SLA',       # Ini Label Text
        'Data Reg3'
    ]
    
    show_cols = [c for c in target_cols if c and c in df.columns]
    
    # Tampilkan Data Table
    st.dataframe(df[show_cols].head(50))

    # Download Button
    excel_bytes = to_excel(df)
    st.download_button(
        "Download Result XLSX",
        data=excel_bytes,
        file_name="reqitem_hasil.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    run()