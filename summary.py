import streamlit as st
import pandas as pd
import plotly.express as px
import io
import re

# --- Fungsi Helper (diambil dari file lain) ---

def find_column(df_cols, possible_names):
    """Helper function to find the first matching column name in a list."""
    for col in possible_names:
        if col in df_cols:
            return col
    return None

def normalize_label(s: str) -> str:
    """Membersihkan dan menormalkan label BC/Severity."""
    if pd.isna(s):
        return ""
    s = str(s).strip()
    s = re.sub(r'\s+', ' ', s)
    s = re.sub(r'\s*-\s*', ' - ', s)
    s = re.sub(r'(?<=\d)(?=[A-Za-z])', ' ', s)
    s = re.sub(r'(?<=\D)(?=\d)', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def format_hari_jam_menit(total_hours_decimal):
    """Mengubah jam desimal menjadi format 'X hari Y jam Z menit'."""
    if pd.isna(total_hours_decimal):
        return "N/A"
    if total_hours_decimal <= 0:
        # Jika tidak breach (negatif atau 0), tampilkan sebagai 0
        return "0 hari 0 jam 0 menit"
    
    total_hours_decimal = float(total_hours_decimal)
    total_minutes = round(total_hours_decimal * 60)
    
    days = total_minutes // (24 * 60)
    minutes_remaining = total_minutes % (24 * 60)
    hours = minutes_remaining // 60
    minutes = minutes_remaining % 60
        
    return f"{days} hari {hours} jam {minutes} menit"

def map_to_hours(label: str, sla_mapping_hours: dict):
    """Mencocokkan label yang dinormalisasi ke jam SLA."""
    if not label:
        return None
    norm_label = normalize_label(label)
    for key in sla_mapping_hours.keys():
        if norm_label == normalize_label(key):
            return sla_mapping_hours[key]
    return None

def process_sla_dataframe(df, type_name: str, sla_mapping_hours: dict):
    """
    Fungsi inti untuk menghitung SLA & Time Breach dari DataFrame mentah.
    Mengembalikan DataFrame yang telah diproses dengan kolom kalkulasi.
    """
    
    # 1. Tentukan nama kolom yang mungkin
    possible_bc_cols = ['Businesscriticality', 'Business criticality', 'Business Criticality', 'BusinessCriticality']
    possible_sev_cols = ['Severity', 'severity', 'SEVERITY']
    possible_created_cols = ['Tiket Dibuat', 'Tiket dibuat', 'Created', 'Created Date', 'CreatedAt']
    possible_resolved_cols = ['Resolved', 'Tiket Ditutup', 'Closed', 'Closed At', 'Tiket ditutup']

    # 2. Cari kolom yang sebenarnya
    bc_col = find_column(df.columns, possible_bc_cols)
    sev_col = find_column(df.columns, possible_sev_cols)
    date_created_col = find_column(df.columns, possible_created_cols)
    date_resolved_col = find_column(df.columns, possible_resolved_cols)
    
    df_calc = df.copy() # Hindari mengubah df asli

    # 3. Validasi
    if not all([bc_col, sev_col, date_created_col]):
        st.warning(f"**[{type_name}]**: Kolom penting (BC, Severity, Created) tidak ditemukan. Tidak dapat menghitung SLA.")
        return df_calc # Kembalikan df asli
    
    if not date_resolved_col:
        st.warning(f"**[{type_name}]**: Kolom 'Resolved'/'Tiket Ditutup' tidak ditemukan. SLA dan Time Breach hanya akan dihitung untuk tiket yang sudah ada datanya (jika kolom ada).")
        # Buat kolom dummy jika tidak ada, agar proses di bawah tidak error
        date_resolved_col = 'Resolved' # Nama placeholder
        if date_resolved_col not in df_calc.columns:
            df_calc[date_resolved_col] = pd.NaT

    # 4. Lakukan perhitungan SLA
    try:
        df_calc[date_created_col] = pd.to_datetime(df_calc[date_created_col], errors='coerce')
        df_calc[date_resolved_col] = pd.to_datetime(df_calc[date_resolved_col], errors='coerce')

        df_calc['_bc_raw'] = df_calc[bc_col].astype(str).fillna('').str.strip()
        df_calc['_sev_raw'] = df_calc[sev_col].astype(str).fillna('').str.strip()
        
        df_calc['Businesscriticality-Severity'] = (df_calc['_bc_raw'] + " - " + df_calc['_sev_raw']).apply(normalize_label)
        
        df_calc['Target SLA (jam)'] = df_calc['Businesscriticality-Severity'].apply(lambda x: map_to_hours(x, sla_mapping_hours))

        df_calc['Target Selesai'] = df_calc.apply(
            lambda r: (r[date_created_col] + pd.to_timedelta(r['Target SLA (jam)'], unit='h'))
            if pd.notna(r[date_created_col]) and pd.notna(r['Target SLA (jam)']) else pd.NaT,
            axis=1
        )

        df_calc['SLA'] = df_calc.apply(
            lambda r: 1 if (pd.notna(r['Target Selesai']) and pd.notna(r[date_resolved_col]) and r[date_resolved_col] <= r['Target Selesai'])
            else (0 if (pd.notna(r['Target Selesai']) and pd.notna(r[date_resolved_col])) else pd.NA),
            axis=1
        )
        
        # 5. Hitung Time Breach
        def calculate_time_breach(row):
            sla_val = row.get('SLA')
            if pd.isna(sla_val): return pd.NA # Open ticket

            if pd.notna(row[date_resolved_col]) and pd.notna(row[date_created_col]) and pd.notna(row['Target SLA (jam)']):
                resolution_duration = row[date_resolved_col] - row[date_created_col]
                sla_timedelta = pd.to_timedelta(row['Target SLA (jam)'], unit='h')
                # Breach = Waktu Resolusi - Alokasi SLA. Positif jika breach.
                breach = resolution_duration - sla_timedelta
                return breach.total_seconds() / 3600 # dalam jam
            return pd.NA

        df_calc['Time Breach'] = df_calc.apply(calculate_time_breach, axis=1)

        return df_calc

    except Exception as e:
        st.error(f"Error saat menghitung SLA untuk {type_name}: {e}")
        return df # Kembalikan df asli jika gagal

def get_sla_summary(df_processed):
    """Mengambil ringkasan SLA dari DataFrame yang sudah diproses."""
    if 'SLA' not in df_processed.columns:
        return {'percent': 0.0, 'achieved': 0, 'not_achieved': 0, 'total_closed': 0}

    sla_achieved = int((df_processed['SLA'] == 1).sum())
    sla_not_achieved = int((df_processed['SLA'] == 0).sum())
    total_closed = sla_achieved + sla_not_achieved

    if total_closed == 0:
        sla_percent = 0.0
    else:
        sla_percent = (sla_achieved / total_closed) * 100

    return {
        'percent': sla_percent,
        'achieved': sla_achieved,
        'not_achieved': sla_not_achieved,
        'total_closed': total_closed
    }

def display_occurrence_table(df_slice, service_col, group_by_col, static_type=None):
    """
    Membuat tabel HTML untuk Top 3 Occurrence.
    Mengembalikan string HTML dari tabel.
    """
    
    # Validasi kolom
    if not service_col in df_slice.columns:
        return "<p>Error: Kolom 'Service Offering' tidak ditemukan.</p>"
    if not static_type and not group_by_col in df_slice.columns:
        return f"<p>Error: Kolom Kategori ('{group_by_col}') tidak ditemukan.</p>"

    # Lakukan agregasi
    try:
        df_agg = df_slice.copy()
        if static_type:
            df_agg['Type'] = static_type
            group_cols = ['Type', service_col]
        else:
            df_agg['Type'] = df_agg[group_by_col].fillna('N/A')
            group_cols = ['Type', service_col]
        
        agg = df_agg.groupby(group_cols).size().reset_index(name='Number of Case')
        agg = agg.sort_values(by=['Type', 'Number of Case'], ascending=[True, False])
        
        # Ambil Top 3 *per Tipe*
        top3_df = agg.groupby('Type').head(3).reset_index(drop=True)

    except Exception as e:
        return f"<p>Error saat agregasi data: {e}</p>"

    # Buat Tabel HTML
    html_table = '<table class="manual-sla-table"><thead><tr>'
    html_table += "<th>Type</th>"
    html_table += "<th>Top 3 Occurrence</th>"
    html_table += "<th>Number of Case</th>"
    html_table += "</tr></thead><tbody>"

    if top3_df.empty:
        html_table += "<tr><td colspan='3' style='text-align:center;'>Tidak ada data untuk ditampilkan.</td></tr>"
    else:
        # Loop berdasarkan 'Type' yang unik untuk membuat rowspan
        for type_name in top3_df['Type'].unique():
            group = top3_df[top3_df['Type'] == type_name]
            rowspan = len(group)
            
            for i, (_, row) in enumerate(group.iterrows()):
                html_table += "<tr>"
                if i == 0:
                    # Baris pertama dari grup, tambahkan sel Tipe dengan rowspan
                    html_table += f"<td rowspan='{rowspan}'>{row['Type']}</td>"
                # Sel untuk Service Offering dan Number of Case
                html_table += f"<td>{row[service_col]}</td>"
                html_table += f"<td>{row['Number of Case']}</td>"
                html_table += "</tr>"
    
    html_table += "</tbody></table>"
    return html_table


# --- Fungsi Utama Streamlit ---

def run():
    st.markdown(
        """
        <h1 style="display: flex; align-items: center; gap: 10px;">
            <img src="https://cdnjs.cloudflare.com/ajax/libs/twemoji/14.0.2/svg/1f4dd.svg" 
                 width="40" height="40">
            General Summary
        </h1>
        """,
        unsafe_allow_html=True
    )
    st.write("Upload 4 file (Incident & Request untuk Agustus & September) untuk melihat ringkasan gabungan.")

    # === CSS Tabel (didefinisikan sekali) ===
    css_tabel = """
    <style>
        .manual-sla-table {
            width: 100%;
            border-collapse: collapse;
            text-align: left;
        }
        .manual-sla-table th, .manual-sla-table td {
            padding: 8px 10px;
            border-bottom: 1px solid #EEEEEE;
            vertical-align: top;
        }
        .manual-sla-table th {
            background-color: #F0F2F6;
            text-align: left;
        }
        .manual-sla-table tr:hover {
            background-color: #F5F5F5;
        }
    </style>
    """
    
    # === Mapping SLA (Global) ===
    sla_mapping_hours = {
        '1 - Critical - 1 - High': 4.0,
        '1 - Critical - 2 - Medium': 6.0,
        '1 - Critical - 3 - Low': 8.0,
        '2 - High - 1 - High': 6.0,
        '2 - High - 2 - Medium': 8.0,
        '2 - High - 3 - Low': 12.0,
        '3 - Medium - 1 - High': 8.0,
        '3 - Medium - 2 - Medium': 12.0,
        '3 - Medium - 3 - Low': 16.0,
        '4 - Low - 1 - High': 16.0,
        '4 - Low - 2 - Medium': 24.0,
        '4 - Low - 3 - Low': 48.0
    }

    # --- 1. File Uploaders (2x2 Grid) ---
    col1, col2 = st.columns(2)
    with col1:
        uploaded_incident_aug = st.file_uploader(
            "Upload Incident File (August)", 
            type=["xlsx", "xls"], 
            key="summary_inc_aug"
        )
        uploaded_request_aug = st.file_uploader(
            "Upload Request Item File (August)", 
            type=["xlsx", "xls"], 
            key="summary_req_aug"
        )
    with col2:
        uploaded_incident_sept = st.file_uploader(
            "Upload Incident File (September)", 
            type=["xlsx", "xls"], 
            key="summary_inc_sept"
        )
        uploaded_request_sept = st.file_uploader(
            "Upload Request Item File (September)", 
            type=["xlsx", "xls"], 
            key="summary_req_sept"
        )

    # --- 2. Check and Load Files ---
    if not all([uploaded_incident_aug, uploaded_incident_sept, uploaded_request_aug, uploaded_request_sept]):
        st.info("Silakan upload keempat file untuk melanjutkan.")
        return

    try:
        df_inc_aug_raw = pd.read_excel(uploaded_incident_aug)
        df_inc_sept_raw = pd.read_excel(uploaded_incident_sept)
        df_req_aug_raw = pd.read_excel(uploaded_request_aug)
        df_req_sept_raw = pd.read_excel(uploaded_request_sept)
    except Exception as e:
        st.error(f"Gagal membaca salah satu file Excel: {e}")
        return

    st.success("Berhasil memuat 4 file.")

    # --- 3. Data Filter (Tambahan) ---
    st.markdown(
        """
        <h1 style="display: flex; align-items: center; gap: 10px;">
            <img src="https://cdn.jsdelivr.net/gh/twitter/twemoji@14.0.2/assets/svg/1f9ea.svg" 
                 width="40" height="40">
            Data Filter (Berlaku untuk semua file)
        </h1>
        """,
        unsafe_allow_html=True
    )
    
    # Gabungkan semua file (hanya untuk mencari opsi filter)
    df_combined_raw = pd.concat([
        df_inc_aug_raw, df_inc_sept_raw, df_req_aug_raw, df_req_sept_raw
    ], ignore_index=True)

    possible_loc_cols = ['Lokasi Pelapor', 'Name', 'User Name', 'Lokasi']

    loc_col = find_column(df_combined_raw.columns, possible_loc_cols)

    regional_option = "All"

    if loc_col:
        regional_option = st.radio(
            "Pilih Lokasi/Regional:",
            options=["All", "Regional 3"],
            horizontal=True,
            key="regional_filter_summary"
        )
    else:
        st.warning("Kolom Lokasi ('Lokasi Pelapor', 'Name', dll.) tidak ditemukan. Filter Regional dinonaktifkan.")

    # Terapkan filter ke setiap DataFrame
    dataframes_raw = {
        "inc_aug": df_inc_aug_raw,
        "inc_sept": df_inc_sept_raw,
        "req_aug": df_req_aug_raw,
        "req_sept": df_req_sept_raw
    }
    
    dataframes_filtered = {}
    total_rows_after_filter = 0

    for name, df in dataframes_raw.items():
        df_filtered = df.copy()
        
        # Apply Regional Filter
        if regional_option == "Regional 3":
            loc_col_in_df = find_column(df_filtered.columns, possible_loc_cols)
            if loc_col_in_df:
                df_filtered = df_filtered[
                    df_filtered[loc_col_in_df].astype(str).str.contains("Regional 3", case=False, na=False)
                ]
        
        dataframes_filtered[name] = df_filtered
        total_rows_after_filter += len(df_filtered)

    # Ganti DataFrame asli dengan yang sudah difilter
    df_inc_aug_raw = dataframes_filtered["inc_aug"]
    df_inc_sept_raw = dataframes_filtered["inc_sept"]
    df_req_aug_raw = dataframes_filtered["req_aug"]
    df_req_sept_raw = dataframes_filtered["req_sept"]

    st.markdown(f"**Total data setelah filter:** {total_rows_after_filter} baris (dari {len(df_combined_raw)} baris awal)")
    
    # Hapus df gabungan sementara
    del df_combined_raw


    # --- 4. Proses SLA untuk semua file (diperlukan untuk line graph & breach table) ---
    df_inc_aug = process_sla_dataframe(df_inc_aug_raw, "Incident August", sla_mapping_hours)
    df_inc_sept = process_sla_dataframe(df_inc_sept_raw, "Incident September", sla_mapping_hours)
    df_req_aug = process_sla_dataframe(df_req_aug_raw, "Request August", sla_mapping_hours)
    df_req_sept = process_sla_dataframe(df_req_sept_raw, "Request September", sla_mapping_hours)

    # --- 5. Total Ticket Metrics ---
    st.subheader("📈 Volume Tiket (Setelah Filter)")

    # --- Mencari kolom 'resolved' secara dinamis ---
    possible_resolved_cols = ['Resolved', 'Tiket Ditutup', 'Closed', 'Closed At', 'Tiket ditutup']
    
    inc_res_col_aug = find_column(df_inc_aug.columns, possible_resolved_cols)
    inc_res_col_sept = find_column(df_inc_sept.columns, possible_resolved_cols)
    req_res_col_aug = find_column(df_req_aug.columns, possible_resolved_cols)
    req_res_col_sept = find_column(df_req_sept.columns, possible_resolved_cols)

    # --- Menghitung Metrik ---
    
    # Total Gabungan
    total_incident = len(df_inc_aug) + len(df_inc_sept)
    total_request = len(df_req_aug) + len(df_req_sept)
    total_all = total_incident + total_request

    # Total Aktif Gabungan (pengecekan 'if inc_res_col_aug else 0' membuatnya aman)
    total_active_incident_aug = df_inc_aug[inc_res_col_aug].isna().sum() if inc_res_col_aug else 0
    total_active_incident_sept = df_inc_sept[inc_res_col_sept].isna().sum() if inc_res_col_sept else 0
    total_active_request_aug = df_req_aug[req_res_col_aug].isna().sum() if req_res_col_aug else 0
    total_active_request_sept = df_req_sept[req_res_col_sept].isna().sum() if req_res_col_sept else 0
    
    total_active_incident = total_active_incident_aug + total_active_incident_sept
    total_active_request = total_active_request_aug + total_active_request_sept

    # Total per Bulan
    total_incident_aug = len(df_inc_aug)
    total_request_aug = len(df_req_aug)
    total_all_aug = total_incident_aug + total_request_aug
    
    total_incident_sept = len(df_inc_sept)
    total_request_sept = len(df_req_sept)
    total_all_sept = total_incident_sept + total_request_sept

    
    # --- Menampilkan Metrik (Layout Baru) ---

    st.markdown("<h4>Ringkasan Total (Gabungan)</h4>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Tiket Insiden", f"{total_incident:,}")
    col2.metric("Total Tiket Request", f"{total_request:,}")
    col3.metric("Total Semua Tiket", f"{total_all:,}")

    col1, col2 = st.columns(2)
    col1.metric("Tiket Insiden Aktif/Pending", f"{total_active_incident:,}")
    col2.metric("Tiket Request Aktif/Pending", f"{total_active_request:,}")

    st.divider()

    st.markdown("<h4>Rincian per Bulan</h4>", unsafe_allow_html=True)
    col_aug, col_sept = st.columns(2)

    # --- Kolom AGUSTUS ---
    with col_aug:
        st.markdown("**Agustus**")
        
        c1, c2, c3 = st.columns(3)
        c1.metric("Insiden", f"{total_incident_aug:,}")
        c2.metric("Request", f"{total_request_aug:,}")
        c3.metric("Total", f"{total_all_aug:,}")

        c1, c2 = st.columns(2)
        c1.metric("Insiden Aktif", f"{total_active_incident_aug:,}")
        c2.metric("Request Aktif", f"{total_active_request_aug:,}")

    # --- Kolom SEPTEMBER ---
    with col_sept:
        st.markdown("**September**")

        c1, c2, c3 = st.columns(3)
        c1.metric("Insiden", f"{total_incident_sept:,}")
        c2.metric("Request", f"{total_request_sept:,}")
        c3.metric("Total", f"{total_all_sept:,}")

        c1, c2 = st.columns(2)
        c1.metric("Insiden Aktif", f"{total_active_incident_sept:,}")
        c2.metric("Request Aktif", f"{total_active_request_sept:,}")

    st.divider()

    # --- 6. Standardize and Combine Channel Data ---
    
    # Tentukan nama kolom channel
    contact_col_incident_names = ['Channel', 'Contact Type', 'ContactType', 'Contact type']
    contact_col_request_names = ['Contact Type', 'ContactType', 'Contact type', 'Channel']

    # Cari kolom channel di setiap file (yang sudah difilter)
    col_inc_aug = find_column(df_inc_aug.columns, contact_col_incident_names)
    col_inc_sept = find_column(df_inc_sept.columns, contact_col_incident_names)
    col_req_aug = find_column(df_req_aug.columns, contact_col_request_names)
    col_req_sept = find_column(df_req_sept.columns, contact_col_request_names)
    
    df_combined_channel = pd.DataFrame(columns=['Channel']) # Buat df kosong

    # Periksa apakah SEMUA kolom ditemukan (agar tidak error)
    all_channel_cols_found = all([col_inc_aug, col_inc_sept, col_req_aug, col_req_sept])

    if not all_channel_cols_found:
        st.error("Tidak dapat menemukan kolom 'Channel' atau 'Contact Type' di salah satu file (setelah filter). Analisis ESS dan Channel dibatalkan.")
    elif total_all == 0:
        st.warning("Tidak ada data tiket setelah filter, Analisis ESS dan Channel dilewati.")
    else:
        # Buat DataFrame ramping gabungan
        df_inc_aug_slim = df_inc_aug[[col_inc_aug]].copy().rename(columns={col_inc_aug: 'Channel'})
        df_inc_sept_slim = df_inc_sept[[col_inc_sept]].copy().rename(columns={col_inc_sept: 'Channel'})
        df_req_aug_slim = df_req_aug[[col_req_aug]].copy().rename(columns={col_req_aug: 'Channel'})
        df_req_sept_slim = df_req_sept[[col_req_sept]].copy().rename(columns={col_req_sept: 'Channel'})

        df_combined_channel = pd.concat([
            df_inc_aug_slim, df_inc_sept_slim, df_req_aug_slim, df_req_sept_slim
        ], ignore_index=True)
        
        df_combined_channel['Channel'] = df_combined_channel['Channel'].fillna('Unknown').astype(str).str.strip()

    
    # --- 7. ESS (Self-Service) Analysis ---
    st.subheader("💻 Analisis Self-Service (ESS)")
    
    if not df_combined_channel.empty:
        ess_keywords = ['ess', 'self-service', 'self service']
        is_ess = df_combined_channel['Channel'].str.lower().isin(ess_keywords)
        total_ess_tickets = int(is_ess.sum())

        if total_all > 0:
            ess_percentage = (total_ess_tickets / total_all) * 100
        else:
            ess_percentage = 0.0

        st.metric(
            "Tiket ESS (Self-Service)", 
            f"{total_ess_tickets:,}",
            f"{ess_percentage:.1f}% dari semua tiket"
        )
    else:
        st.warning("Tidak ada data channel untuk dianalisis.")

    # --- 8. Channel Donut Chart (Combined) ---
    st.subheader("📊 Distribusi Channel (Semua Tiket)")
    
    if not df_combined_channel.empty:
        channel_summary = df_combined_channel['Channel'].value_counts().reset_index()
        channel_summary.columns = ['Channel', 'Count']
        
        fig_channel = px.pie(
            channel_summary,
            names='Channel',
            values='Count',
            title='Distribusi Tiket Gabungan berdasarkan Channel',
            hole=0.4
        )
        fig_channel.update_traces(textinfo='percent+label')
        st.plotly_chart(fig_channel, use_container_width=True)
        
        st.dataframe(channel_summary)
    else:
        st.warning("Tidak ada data channel untuk ditampilkan.")
        
    
    # --- 9. SLA Performance Line Graph ---
    st.subheader("📉 Performa SLA dari Waktu ke Waktu")

    # Ambil ringkasan dari DataFrame yang sudah diproses
    stats_inc_aug = get_sla_summary(df_inc_aug)
    stats_inc_sept = get_sla_summary(df_inc_sept)
    stats_req_aug = get_sla_summary(df_req_aug)
    stats_req_sept = get_sla_summary(df_req_sept)

    # Periksa apakah ada data
    if stats_inc_aug['total_closed'] > 0 or stats_inc_sept['total_closed'] > 0 or stats_req_aug['total_closed'] > 0 or stats_req_sept['total_closed'] > 0:
        
        # Siapkan data untuk grafik
        data = [
            {'Month': 'August', 'Type': 'Incident', 'SLA (%)': stats_inc_aug['percent']},
            {'Month': 'September', 'Type': 'Incident', 'SLA (%)': stats_inc_sept['percent']},
            {'Month': 'August', 'Type': 'Request', 'SLA (%)': stats_req_aug['percent']},
            {'Month': 'September', 'Type': 'Request', 'SLA (%)': stats_req_sept['percent']},
        ]
        chart_df = pd.DataFrame(data)

        # Buat line chart
        fig_line = px.line(
            chart_df,
            x='Month',
            y='SLA (%)',
            color='Type',        # Garis terpisah untuk 'Incident' dan 'Request'
            markers=True,        # Tambahkan titik di setiap data point
            text='SLA (%)',      # Tampilkan nilai di atas titik
            title='Pencapaian SLA (%) vs. Bulan'
        )
        
        # Format teks agar menampilkan 1 angka desimal
        fig_line.update_traces(texttemplate='%{y:.1f}%', textposition='top center')
        fig_line.update_layout(
            yaxis_title="SLA Achievement (%)",
            xaxis_title="Bulan",
            legend_title="Tipe Tiket",
            yaxis_range=[0, 105] # Set range Y dari 0-105%
        )
        st.plotly_chart(fig_line, use_container_width=True)

        # Tampilkan data tabelnya juga
        st.dataframe(chart_df.pivot(index='Month', columns='Type', values='SLA (%)'))

    else:
        st.error("Grafik performa SLA tidak dapat dibuat karena tidak ada tiket yang ditutup (resolved) di semua file.")

    # --- 10. Top 3 Breach Time Table (New Requirement) ---
    st.subheader("🔥 Top 3 Service Offering dengan Max Breach Terbesar")
    
    # Gabungkan semua dataFrame yang telah diproses
    df_combined_full = pd.concat([
        df_inc_aug, df_inc_sept, df_req_aug, df_req_sept
    ], ignore_index=True)

    # Cari kolom Service Offering
    possible_service_cols = ['Service offering', 'Service Offering', 'ServiceOffering']
    service_col = find_column(df_combined_full.columns, possible_service_cols)
    
    if not service_col:
        st.error("Kolom 'Service Offering' tidak ditemukan di file mana pun. Tidak dapat membuat tabel Top Breach.")
    elif 'Time Breach' not in df_combined_full.columns:
        st.error("Kolom 'Time Breach' gagal dihitung. Tidak dapat membuat tabel Top Breach.")
    elif df_combined_full.empty:
        st.warning("Tidak ada data untuk dianalisis di tabel Top Breach (kemungkinan semua terfilter habis).")
    else:
        # Lakukan agregasi sesuai permintaan baru
        sla_service_agg = df_combined_full.groupby(service_col).agg(
            # Temukan nilai 'Time Breach' maksimum (paling parah)
            Max_Time_Breach=('Time Breach', 'max'),
            # Hitung semua tiket untuk service tsb
            Total_Tiket=(service_col, 'size'), 
            # Hitung hanya tiket yang SLA = 0 (breach)
            Tiket_Breach=('SLA', lambda x: (x == 0).sum())
        ).reset_index()
        
        # Ganti NaN di Max_Time_Breach (jika service tsb tidak punya tiket breach) dengan 0
        sla_service_agg['Max_Time_Breach'] = sla_service_agg['Max_Time_Breach'].fillna(0)

        # Urutkan berdasarkan Max_Time_Breach terbesar
        sla_service_agg = sla_service_agg.sort_values(by='Max_Time_Breach', ascending=False)
        
        # Buat ranking
        sla_service_agg['No'] = sla_service_agg['Max_Time_Breach'].rank(method='dense', ascending=False).astype(int)
        
        # Ambil top 3 (rank 1, 2, 3)
        bottom3_sla = sla_service_agg[sla_service_agg['No'] <= 3].sort_values(by=['No', service_col])
        
        # Sesuaikan kolom
        bottom3_sla = bottom3_sla[['No', service_col, 'Max_Time_Breach', 'Total_Tiket', 'Tiket_Breach']]
        bottom3_sla.columns = ['No', 'Service Offering', 'Max Time Breach', '∑Total Tiket', '∑ Tiket Breach']

        # --- Tampilkan Tabel HTML ---
        # (CSS didefinisikan di awal fungsi run())
        html_bottom = css_tabel + '<table class="manual-sla-table"><thead><tr>'
        html_bottom += "<th>No</th>"
        html_bottom += "<th>Service Offering</th>"
        html_bottom += "<th>Max Time Breach</th>"
        html_bottom += "<th>∑Total Tiket</th>"
        html_bottom += "<th>∑ Tiket Breach</th>"
        html_bottom += "</tr></thead><tbody>"

        if bottom3_sla.empty:
            html_bottom += "<tr><td colspan='5' style='text-align:center;'>Tidak ada data breach untuk ditampilkan.</td></tr>"
        else:
            for _, row in bottom3_sla.iterrows():
                html_bottom += "<tr>"
                html_bottom += f"<td>{row['No']}</td>"
                html_bottom += f"<td>{row['Service Offering']}</td>"
                # Format kolom Max Time Breach
                html_bottom += f"<td>{format_hari_jam_menit(row['Max Time Breach'])}</td>"
                html_bottom += f"<td>{row['∑Total Tiket']}</td>"
                html_bottom += f"<td>{row['∑ Tiket Breach']}</td>"
                html_bottom += "</tr>"

        html_bottom += "</tbody></table>"
        st.markdown(html_bottom, unsafe_allow_html=True)
    
    
    # --- 11. Top Occurrence Analysis ---
    st.subheader("🔎 Top Occurrence Analysis by Category")
    
    time_filter_occurrence = st.radio(
        "Select Time Period for Analysis:", 
        ["All", "August", "September"], 
        horizontal=True, 
        key="top_occurrence_filter"
    )

    # Tentukan slice data berdasarkan filter
    if time_filter_occurrence == "August":
        inc_df_slice = df_inc_aug
        req_df_slice = df_req_aug
    elif time_filter_occurrence == "September":
        inc_df_slice = df_inc_sept
        # **BUG FIX 2: Typo diperbaiki**
        req_df_slice = df_req_sept 
    else: # "All"
        inc_df_slice = pd.concat([df_inc_aug, df_inc_sept], ignore_index=True)
        req_df_slice = pd.concat([df_req_aug, df_req_sept], ignore_index=True)

    # Tentukan nama kolom (service_col sudah ada dari section 10)
    possible_kategori_cols = ['Kategori', 'Category', 'Tipe']
    # Kita cari di DUA dataframe incident (mentah), karena salah satunya mungkin kosong
    kategori_col = find_column(df_inc_aug_raw.columns, possible_kategori_cols) or find_column(df_inc_sept_raw.columns, possible_kategori_cols)

    # Tampilkan Tabel Incident
    st.markdown("<h4>Incident Analysis (Top 3 Occurrence)</h4>", unsafe_allow_html=True)
    if not kategori_col:
        st.error("Kolom 'Kategori' tidak ditemukan di file Incident. Tidak dapat membuat tabel.")
    elif not service_col:
        st.error("Kolom 'Service Offering' tidak ditemukan. Tidak dapat membuat tabel.")
    elif inc_df_slice.empty:
         st.warning("Tidak ada data Insiden untuk periode ini.")
    else:
        html_incident = display_occurrence_table(
            df_slice=inc_df_slice, 
            service_col=service_col, 
            group_by_col=kategori_col, 
            static_type=None
        )
        st.markdown(css_tabel + html_incident, unsafe_allow_html=True)

    # Tampilkan Tabel Request
    st.markdown("<h4>Request Analysis (Top 3 Occurrence)</h4>", unsafe_allow_html=True)
    if not service_col:
        st.error("Kolom 'Service Offering' tidak ditemukan. Tidak dapat membuat tabel.")
    elif req_df_slice.empty:
        st.warning("Tidak ada data Request untuk periode ini.")
    else:
        html_request = display_occurrence_table(
            df_slice=req_df_slice,
            service_col=service_col,
            group_by_col=None,
            static_type="Request"
        )
        st.markdown(css_tabel + html_request, unsafe_allow_html=True)

    
    # --- 12. Solved vs Active/Pending Status (FITUR BARU) ---
    st.subheader("📊 Solved vs Active/Pending Status")
    st.caption(f"Menampilkan status untuk periode: **{time_filter_occurrence}**")

    # Kolom 'kategori_col' sudah ditemukan di section 11
    # 'inc_df_slice' dan 'req_df_slice' juga sudah di-filter
    
    # Kita perlu menemukan kolom resolved *lagi* di dalam slice
    inc_res_col = find_column(inc_df_slice.columns, possible_resolved_cols)
    req_res_col = find_column(req_df_slice.columns, possible_resolved_cols)

    if not kategori_col:
        st.error("Kolom 'Kategori' tidak ditemukan di file Incident. Tidak dapat membuat tabel status.")
    elif not inc_res_col or not req_res_col:
        st.error("Kolom 'Resolved' / 'Tiket Ditutup' tidak ditemukan. Tidak dapat membuat tabel status.")
    elif inc_df_slice.empty and req_df_slice.empty:
        st.warning("Tidak ada data untuk ditampilkan di tabel status.")
    else:
        solved_data = {}
        active_data = {}
        
        # Dapatkan semua Tipe Insiden yang unik
        incident_types = []
        if not inc_df_slice.empty:
            incident_types = sorted(list(inc_df_slice[kategori_col].dropna().unique()))
        
        all_table_cols = incident_types + ["Request"]
        
        # 1. Hitung data Insiden
        if not inc_df_slice.empty:
            for type_name in incident_types:
                type_mask = (inc_df_slice[kategori_col] == type_name)
                # Solved = notna()
                solved_count = inc_df_slice[type_mask & inc_df_slice[inc_res_col].notna()].shape[0]
                # Active = isna()
                active_count = inc_df_slice[type_mask & inc_df_slice[inc_res_col].isna()].shape[0]
                
                solved_data[type_name] = solved_count
                active_data[type_name] = active_count
        
        # 2. Hitung data Request
        if not req_df_slice.empty:
            req_solved_count = req_df_slice[req_df_slice[req_res_col].notna()].shape[0]
            req_active_count = req_df_slice[req_df_slice[req_res_col].isna()].shape[0]
        else:
            req_solved_count = 0
            req_active_count = 0
            
        solved_data["Request"] = req_solved_count
        active_data["Request"] = req_active_count
        
        # 3. Buat DataFrame
        final_table_data = [solved_data, active_data]
        df_status = pd.DataFrame(final_table_data, index=["Solved", "Active/Pending"])
        
        # Pastikan urutan kolom benar
        df_status = df_status[all_table_cols]
        
        st.dataframe(df_status, use_container_width=True)


if __name__ == "__main__":
    run()