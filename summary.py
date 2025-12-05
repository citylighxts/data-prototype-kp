import streamlit as st
import pandas as pd
import plotly.express as px
import io
import re
from datetime import time, timedelta, datetime
import numpy as np
import calendar

def get_table_css():
    return """
    <style>
        .manual-sla-table {
            width: 100%;
            border-collapse: collapse;
            font-family: sans-serif;
            font-size: 13px;
        }
        /* Header Blue Style */
        .manual-sla-table th {
            background-color: #305496; /* Dark Blue */
            color: white;
            padding: 8px 10px;
            border: 1px solid white; /* White borders */
            text-align: center;
            font-weight: bold;
        }
        /* General Cell Style */
        .manual-sla-table td {
            padding: 6px 10px;
            border: 1px solid white; /* White grid lines */
            vertical-align: middle;
            color: black;
        }
        /* Grouping Column (No/Type) - Gray Background */
        .col-no {
            background-color: #D9D9D9; 
            font-weight: bold;
            text-align: center;
            vertical-align: middle;
        }
        /* Data Row - Light Gray Background */
        .row-data {
            background-color: #E9E9E9; 
        }
        /* Center align numbers */
        .text-center {
            text-align: center;
        }
    </style>
    """

def find_column(df_cols, possible_names):
    """Mencari nama kolom yang cocok pertama dalam list."""
    for col in possible_names:
        if col in df_cols:
            return col
    return None

def normalize_label(s: str) -> str:
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
    Fungsi inti untuk menghitung SLA & Time Breach.
    """
    
    possible_bc_cols = ['Businesscriticality', 'Business criticality', 'Business Criticality', 'BusinessCriticality']
    possible_sev_cols = ['Severity', 'severity', 'SEVERITY']
    possible_created_cols = ['Tiket Dibuat', 'Tiket dibuat', 'Created', 'Created Date', 'CreatedAt']
    possible_resolved_cols = ['Resolved', 'Tiket Ditutup', 'Closed', 'Closed At', 'Tiket ditutup']

    bc_col = find_column(df.columns, possible_bc_cols)
    sev_col = find_column(df.columns, possible_sev_cols)
    date_created_col = find_column(df.columns, possible_created_cols)
    date_resolved_col = find_column(df.columns, possible_resolved_cols)
    
    df_calc = df.copy() 

    if not all([bc_col, sev_col, date_created_col]):
        st.warning(f"**[{type_name}]**: Kolom penting (BC, Severity, Created) tidak ditemukan. Tidak dapat menghitung SLA.")
        return df_calc 
    
    if not date_resolved_col:
        st.warning(f"**[{type_name}]**: Kolom 'Resolved'/'Tiket Ditutup' tidak ditemukan.")
        date_resolved_col = 'Resolved_Placeholder' 
        if date_resolved_col not in df_calc.columns:
            df_calc[date_resolved_col] = pd.NaT

    try:
        df_calc[date_created_col] = pd.to_datetime(df_calc[date_created_col], errors='coerce')
        df_calc[date_resolved_col] = pd.to_datetime(df_calc[date_resolved_col], errors='coerce')

        df_calc = df_calc.dropna(subset=[date_created_col]) 
        df_calc['Month'] = df_calc[date_created_col].dt.strftime('%Y-%m (%B)')

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
        
        def calculate_time_breach(row):
            sla_val = row.get('SLA')
            if pd.isna(sla_val): return pd.NA 

            if pd.notna(row[date_resolved_col]) and pd.notna(row[date_created_col]) and pd.notna(row['Target SLA (jam)']):
                resolution_duration = row[date_resolved_col] - row[date_created_col]
                sla_timedelta = pd.to_timedelta(row['Target SLA (jam)'], unit='h')
                breach = resolution_duration - sla_timedelta
                return breach.total_seconds() / 3600 
            return pd.NA

        df_calc['Time Breach'] = df_calc.apply(calculate_time_breach, axis=1)

        return df_calc

    except Exception as e:
        st.error(f"Error saat menghitung SLA untuk {type_name}: {e}")
        return df

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

def display_occurrence_table(df_slice, data_col, group_by_col, static_type=None, limit=None):
    """
    Membuat tabel HTML untuk Occurrence dengan STYLE BARU.
    """
    
    if not data_col in df_slice.columns:
        return f"<p>Error: Kolom data utama ('{data_col}') tidak ditemukan.</p>"
    
    if not static_type and (not group_by_col or not group_by_col in df_slice.columns):
        return f"<p>Error: Kolom Kategori ('{group_by_col}') tidak ditemukan.</p>"

    try:
        df_agg = df_slice.copy()
        if static_type:
            df_agg['Type'] = static_type
            group_cols = ['Type', data_col]
        else:
            df_agg['Type'] = df_agg[group_by_col].fillna('N/A')
            group_cols = ['Type', data_col]
        
        agg = df_agg.groupby(group_cols).size().reset_index(name='Number of Case')
        agg = agg.sort_values(by=['Type', 'Number of Case'], ascending=[True, False])
        
        if limit:
            final_df = agg.groupby('Type').head(limit).reset_index(drop=True)
            header_text = f"Top {limit} Occurrence"
        else:
            final_df = agg.reset_index(drop=True)
            header_text = "Occurrence (All)"

    except Exception as e:
        return f"<p>Error saat agregasi data: {e}</p>"

    html_table = '<table class="manual-sla-table"><thead><tr>'
    html_table += "<th>Type</th>"
    html_table += f"<th>{header_text} ({data_col})</th>"
    html_table += "<th>Number of Case</th>"
    html_table += "</tr></thead><tbody>"

    if final_df.empty:
        html_table += "<tr class='row-data'><td colspan='3' style='text-align:center;'>Tidak ada data untuk ditampilkan.</td></tr>"
    else:
        for type_name in final_df['Type'].unique():
            group = final_df[final_df['Type'] == type_name]
            rowspan = len(group)
            
            for i, (_, row) in enumerate(group.iterrows()):
                html_table += "<tr class='row-data'>"
                if i == 0:
                    html_table += f"<td rowspan='{rowspan}' class='col-no'>{row['Type']}</td>"
                html_table += f"<td>{row[data_col]}</td>"
                html_table += f"<td class='text-center'>{row['Number of Case']}</td>"
                html_table += "</tr>"
    
    html_table += "</tbody></table>"
    return html_table

def make_simple_html_table(df):
    """Konversi DataFrame sederhana ke Tabel HTML dengan Style Baru."""
    html = '<table class="manual-sla-table"><thead><tr>'
    html += f"<th>Status</th>"
    for col in df.columns:
        html += f"<th>{col}</th>"
    html += "</tr></thead><tbody>"

    for idx, row in df.iterrows():
        html += "<tr class='row-data'>"
        html += f"<td class='col-no'>{idx}</td>"
        for val in row:
            html += f"<td class='text-center'>{val}</td>"
        html += "</tr>"
    html += "</tbody></table>"
    return html


def run():
    st.markdown(get_table_css(), unsafe_allow_html=True)

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

    st.subheader("Upload File")
    num_months = st.selectbox(
        "Pilih jumlah periode/bulan yang akan dianalisis:",
        options=list(range(1, 13)),
        index=1
    )

    uploaded_incident_files = []
    uploaded_request_files = []
    
    st.write("Silakan upload file Incident dan Request untuk setiap periode:")
    
    cols = st.columns(2)
    for i in range(num_months):
        with cols[0]:
            file_inc = st.file_uploader(
                f"Upload Incident File (Bulan {i+1})", 
                type=["xlsx", "xls"], 
                key=f"summary_inc_{i}"
            )
            uploaded_incident_files.append(file_inc)
        with cols[1]:
            file_req = st.file_uploader(
                f"Upload Request Item File (Bulan {i+1})", 
                type=["xlsx", "xls"], 
                key=f"summary_req_{i}"
            )
            uploaded_request_files.append(file_req)

    if not all(uploaded_incident_files) or not all(uploaded_request_files):
        st.info("Harap lengkapi semua file uploader di atas untuk melanjutkan.")
        return

    try:
        list_df_inc_raw = [pd.read_excel(f) for f in uploaded_incident_files]
        list_df_req_raw = [pd.read_excel(f) for f in uploaded_request_files]
    except Exception as e:
        st.error(f"Gagal membaca salah satu file Excel: {e}")
        return

    st.success(f"Berhasil memuat {len(list_df_inc_raw)} file Incident dan {len(list_df_req_raw)} file Request.")

    list_df_inc_processed = []
    list_df_req_processed = []

    for i, df_raw in enumerate(list_df_inc_raw):
        df_processed = process_sla_dataframe(df_raw, f"Incident (File {i+1})", sla_mapping_hours)
        list_df_inc_processed.append(df_processed)

    for i, df_raw in enumerate(list_df_req_raw):
        df_processed = process_sla_dataframe(df_raw, f"Request (File {i+1})", sla_mapping_hours)
        list_df_req_processed.append(df_processed)

    st.subheader("Data Filter")
    
    df_combined_raw = pd.concat(list_df_inc_raw + list_df_req_raw, ignore_index=True)

    possible_loc_cols = ['Lokasi Pelapor', 'Name', 'User Name', 'Lokasi']
    loc_col = find_column(df_combined_raw.columns, possible_loc_cols)
    regional_option = "All"

    if loc_col:
        regional_option = st.radio(
            "Filter Lokasi:",
            options=["All", "Regional 3 (Request)"],
            horizontal=True,
            key="regional_filter_summary"
        )
    else:
        st.warning("Kolom Lokasi tidak ditemukan. Filter Regional dinonaktifkan.")

    list_df_inc_filtered = []
    list_df_req_filtered = []
    total_rows_after_filter = 0

    for df_proc in list_df_inc_processed:
        list_df_inc_filtered.append(df_proc.copy())
        total_rows_after_filter += len(df_proc)

    for df_proc in list_df_req_processed:
        df_filtered = df_proc.copy()
        if regional_option == "Regional 3 (Request)":
            loc_col_in_df = find_column(df_filtered.columns, possible_loc_cols)
            if loc_col_in_df:
                df_filtered = df_filtered[
                    df_filtered[loc_col_in_df].astype(str).str.strip().isin(regional_3_locations)
                ]
        list_df_req_filtered.append(df_filtered)
        total_rows_after_filter += len(df_filtered)

    st.markdown(f"**Total data yang diolah:** {total_rows_after_filter} baris")
    del df_combined_raw

    df_inc_all = pd.concat(list_df_inc_filtered, ignore_index=True)
    df_req_all = pd.concat(list_df_req_filtered, ignore_index=True)
    df_combined_full = pd.concat([df_inc_all, df_req_all], ignore_index=True)

    st.subheader("Volume Tiket")

    possible_resolved_cols = ['Resolved', 'Tiket Ditutup', 'Closed', 'Closed At', 'Tiket ditutup']
    
    inc_res_col = find_column(df_inc_all.columns, possible_resolved_cols)
    req_res_col = find_column(df_req_all.columns, possible_resolved_cols)

    total_incident = len(df_inc_all)
    total_request = len(df_req_all)
    total_all = total_incident + total_request

    total_active_incident = df_inc_all[inc_res_col].isna().sum() if inc_res_col else 0
    total_active_request = df_req_all[req_res_col].isna().sum() if req_res_col else 0

    col1, col2, col3 = st.columns(3)
    col1.metric("Total Tiket Insiden", f"{total_incident:,}")
    col2.metric("Total Tiket Request", f"{total_request:,}")
    col3.metric("Total Semua Tiket", f"{total_all:,}")

    col1, col2 = st.columns(2)
    col1.metric("Tiket Insiden Aktif/Pending", f"{total_active_incident:,}")
    col2.metric("Tiket Request Aktif/Pending", f"{total_active_request:,}")

    st.divider()

    st.markdown("<h4>Rincian per Bulan</h4>", unsafe_allow_html=True)
    
    agg_inc_monthly = df_inc_all.groupby('Month').agg(
        Incident=('Month', 'size'),
        Incident_Aktif=(inc_res_col, lambda x: x.isna().sum())
    ) if inc_res_col else df_inc_all.groupby('Month').agg(Incident=('Month', 'size'))
    
    agg_req_monthly = df_req_all.groupby('Month').agg(
        Request=('Month', 'size'),
        Request_Aktif=(req_res_col, lambda x: x.isna().sum())
    ) if req_res_col else df_req_all.groupby('Month').agg(Request=('Month', 'size'))
    
    df_monthly_summary = pd.concat([agg_inc_monthly, agg_req_monthly], axis=1).fillna(0).astype(int)
    
    for col in ['Incident', 'Request', 'Incident_Aktif', 'Request_Aktif']:
        if col not in df_monthly_summary: df_monthly_summary[col] = 0

    df_monthly_summary = df_monthly_summary.sort_index() 
    
    df_monthly_chart = df_monthly_summary[['Incident', 'Request']].reset_index().melt(
        id_vars='Month', var_name='Type', value_name='Count'
    )
    df_monthly_active_chart = df_monthly_summary[['Incident_Aktif', 'Request_Aktif']].reset_index().melt(
        id_vars='Month', var_name='Type', value_name='Count'
    )
    month_order = df_monthly_summary.index.tolist()

    if df_monthly_summary.empty:
        st.warning("Tidak ada data bulanan untuk ditampilkan.")
    else:
        col_chart1, col_chart2 = st.columns(2)
        
        with col_chart1:
            st.markdown("**Total Tiket Dibuat (per Bulan)**")
            fig_monthly_total = px.bar(
                df_monthly_chart,
                x="Month",
                y="Count",
                color="Type", 
                barmode="group", 
                text_auto=True,
                category_orders={"Month": month_order} 
            )
            st.plotly_chart(fig_monthly_total, use_container_width=True)

        with col_chart2:
            st.markdown("**Total Tiket Aktif/Pending (per Bulan)**")
            fig_monthly_active = px.bar(
                df_monthly_active_chart,
                x="Month",
                y="Count",
                color="Type",
                barmode="group",
                text_auto=True,
                category_orders={"Month": month_order}
            )
            st.plotly_chart(fig_monthly_active, use_container_width=True)
            
    st.divider()

    contact_col_incident_names = ['Channel', 'Contact Type', 'ContactType', 'Contact type']
    contact_col_request_names = ['Contact Type', 'ContactType', 'Contact type', 'Channel']

    col_inc = find_column(df_inc_all.columns, contact_col_incident_names)
    col_req = find_column(df_req_all.columns, contact_col_request_names)
    
    df_combined_channel = pd.DataFrame(columns=['Channel']) 

    if total_all > 0 and col_inc and col_req:
        df_inc_slim = df_inc_all[[col_inc]].copy().rename(columns={col_inc: 'Channel'})
        df_req_slim = df_req_all[[col_req]].copy().rename(columns={col_req: 'Channel'})

        df_combined_channel = pd.concat([df_inc_slim, df_req_slim], ignore_index=True)
        df_combined_channel['Channel'] = df_combined_channel['Channel'].fillna('Unknown').astype(str).str.strip()

    c1, c2 = st.columns(2)

    with c1:
        st.subheader("Analisis Self-Service (ESS)")
        if not df_combined_channel.empty:
            ess_keywords = ['ess', 'self-service', 'self service']
            is_ess = df_combined_channel['Channel'].str.lower().isin(ess_keywords)
            total_ess_tickets = int(is_ess.sum())
            ess_percentage = (total_ess_tickets / total_all) * 100 if total_all > 0 else 0.0

            st.metric("Tiket ESS (Self-Service)", f"{total_ess_tickets:,}", f"{ess_percentage:.1f}%")
        else:
            st.warning("Tidak ada data channel untuk dianalisis.")

        st.subheader("Top 3 Service Offering dengan Max Breach Terbesar")
        
        unique_months = sorted(df_combined_full['Month'].dropna().unique().tolist())
        time_filter_options = ["All"] + unique_months
        
        time_filter_selection = st.radio(
            "Pilih periode waktu:", 
            time_filter_options, 
            horizontal=True, 
            key="time_period_filter_summary"
        )

        if time_filter_selection == "All":
            inc_df_slice = df_inc_all
            req_df_slice = df_req_all
        else:
            inc_df_slice = df_inc_all[df_inc_all['Month'] == time_filter_selection]
            req_df_slice = df_req_all[df_req_all['Month'] == time_filter_selection]

        df_combined_full_slice = pd.concat([inc_df_slice, req_df_slice], ignore_index=True)
        
        possible_service_cols = ['Service offering', 'Service Offering', 'ServiceOffering']
        service_col = find_column(df_combined_full_slice.columns, possible_service_cols)
        
        if not service_col:
            st.error("Kolom 'Service Offering' tidak ditemukan.")
        elif 'Time Breach' not in df_combined_full_slice.columns:
            st.error("Kolom 'Time Breach' gagal dihitung.")
        elif df_combined_full_slice.empty or df_combined_full_slice['Time Breach'].isna().all():
            st.warning("Tidak ada data breach untuk dianalisis.")
        else:
            sla_service_agg = df_combined_full_slice.groupby(service_col).agg(
                Max_Time_Breach=('Time Breach', 'max'),
                Total_Tiket=(service_col, 'size'), 
                Tiket_Breach=('SLA', lambda x: (x == 0).sum())
            ).reset_index()
            
            sla_service_agg['Max_Time_Breach'] = sla_service_agg['Max_Time_Breach'].fillna(0)
            sla_service_agg = sla_service_agg.sort_values(by='Max_Time_Breach', ascending=False)
            sla_service_agg['No'] = sla_service_agg['Max_Time_Breach'].rank(method='dense', ascending=False).astype(int)
            
            bottom3_sla = sla_service_agg[sla_service_agg['No'] <= 3].sort_values(by=['No', service_col])
            bottom3_sla = bottom3_sla[['No', service_col, 'Max_Time_Breach', 'Total_Tiket', 'Tiket_Breach']]
            bottom3_sla.columns = ['No', 'Service Offering', 'Max Time Breach', '∑Total Tiket', '∑ Tiket Breach']

            html_bottom = '<table class="manual-sla-table"><thead><tr>'
            html_bottom += "<th>No</th>"
            html_bottom += "<th>Service Offering</th>"
            html_bottom += "<th>Max Time Breach</th>"
            html_bottom += "<th>∑Total Tiket</th>"
            html_bottom += "<th>∑ Tiket Breach</th>"
            html_bottom += "</tr></thead><tbody>"

            if bottom3_sla.empty:
                html_bottom += "<tr class='row-data'><td colspan='5' style='text-align:center;'>Tidak ada data breach untuk ditampilkan.</td></tr>"
            else:
                for _, row in bottom3_sla.iterrows():
                    html_bottom += "<tr class='row-data'>"
                    html_bottom += f"<td class='col-no'>{row['No']}</td>"
                    html_bottom += f"<td><b>{row['Service Offering']}</b></td>"
                    html_bottom += f"<td class='text-center'>{format_hari_jam_menit(row['Max Time Breach'])}</td>"
                    html_bottom += f"<td class='text-center'>{row['∑Total Tiket']}</td>"
                    html_bottom += f"<td class='text-center'>{row['∑ Tiket Breach']}</td>"
                    html_bottom += "</tr>"

            html_bottom += "</tbody></table>"
            st.markdown(html_bottom, unsafe_allow_html=True)

    with c2:
        st.subheader("Distribusi Channel")
        
        tab_all, tab_inc, tab_req = st.tabs(["Semua Tiket", "Incident", "Request"])

        with tab_all:
            if not df_combined_channel.empty:
                channel_summary = df_combined_channel['Channel'].value_counts().reset_index()
                channel_summary.columns = ['Channel', 'Count']
                
                fig_channel_all = px.pie(
                    channel_summary,
                    names='Channel',
                    values='Count',
                    title='Semua Tiket',
                    hole=0.4
                )
                fig_channel_all.update_traces(textinfo='percent+label')
                fig_channel_all.update_layout(margin=dict(t=30, b=0, l=0, r=0))
                st.plotly_chart(fig_channel_all, use_container_width=True)
            else:
                st.warning("Tidak ada data channel.")

        with tab_inc:
            if col_inc and not df_inc_all.empty:
                inc_channel_summary = df_inc_all[col_inc].fillna('Unknown').value_counts().reset_index()
                inc_channel_summary.columns = ['Channel', 'Count']
                
                fig_channel_inc = px.pie(
                    inc_channel_summary,
                    names='Channel',
                    values='Count',
                    title='Incident',
                    hole=0.4
                )
                fig_channel_inc.update_traces(textinfo='percent+label')
                fig_channel_inc.update_layout(margin=dict(t=30, b=0, l=0, r=0))
                st.plotly_chart(fig_channel_inc, use_container_width=True)
            else:
                st.info("Tidak ada data Incident.")

        with tab_req:
            if col_req and not df_req_all.empty:
                req_channel_summary = df_req_all[col_req].fillna('Unknown').value_counts().reset_index()
                req_channel_summary.columns = ['Channel', 'Count']
                
                fig_channel_req = px.pie(
                    req_channel_summary,
                    names='Channel',
                    values='Count',
                    title='Request',
                    hole=0.4
                )
                fig_channel_req.update_traces(textinfo='percent+label')
                fig_channel_req.update_layout(margin=dict(t=30, b=0, l=0, r=0))
                st.plotly_chart(fig_channel_req, use_container_width=True)
            else:
                st.info("Tidak ada data Request.")

    st.divider()

    st.subheader("Performa SLA")

    # Hitung summary SLA per bulan
    inc_sla_monthly = df_inc_all.groupby('Month').apply(get_sla_summary)
    req_sla_monthly = df_req_all.groupby('Month').apply(get_sla_summary)

    data_points = []
    
    # Proses data Incident
    for month, stats in inc_sla_monthly.items():
        if stats['total_closed'] > 0:
            data_points.append({
                'Month': month, 
                'Type': 'Incident', 
                'SLA (%)': stats['percent'], 
                'Achieved': stats['achieved'],         
                'Total Closed': stats['total_closed']   
            })
            
    # Proses data Request
    for month, stats in req_sla_monthly.items():
        if stats['total_closed'] > 0:
            data_points.append({
                'Month': month, 
                'Type': 'Request', 
                'SLA (%)': stats['percent'],
                'Achieved': stats['achieved'],         
                'Total Closed': stats['total_closed']   
            })

    if data_points:
        chart_df = pd.DataFrame(data_points).sort_values(by='Month')
        
        # Label text di chart
        chart_df['Label'] = (
            chart_df['SLA (%)'].map('{:.1f}'.format) + "% (" + 
            chart_df['Achieved'].astype(str) + "/" + 
            chart_df['Total Closed'].astype(str) + ")"
        )
        
        # Format tampilan Bulan (Menghilangkan YYYY-MM di depan agar lebih rapi)
        # Asumsi format 'Month' adalah "YYYY-MM (NamaBulan)"
        chart_df['Month_Display'] = chart_df['Month'].str.split('(').str[1].str.replace(')', '') + " " + chart_df['Month'].str.split('-').str[0]
        
        # Mengurutkan berdasarkan month_order yang sudah dibuat di bagian Volume Tiket
        # agar urutan bulan kronologis (bukan abjad)
        valid_months = [m for m in month_order if m in chart_df['Month'].unique()]
        month_display_map = chart_df.drop_duplicates(subset=['Month']).set_index('Month')['Month_Display'].to_dict()
        month_display_order = [month_display_map[m] for m in valid_months if m in month_display_map]

        fig_line = px.line(
            chart_df,
            x='Month_Display',
            y='SLA (%)',
            color='Type',        
            markers=True,        
            text='Label',   
            hover_data=['Achieved', 'Total Closed'], 
            title='Pencapaian SLA (%) vs. Bulan'
        )
        
        fig_line.update_traces(textposition='top center')
        
        fig_line.update_layout(
            yaxis_title="SLA Achievement (%)",
            xaxis_title="Bulan",
            yaxis_range=[0, 115], # Memberi ruang untuk text label di atas 100%
            xaxis={'categoryorder':'array', 'categoryarray': month_display_order} 
        )
        st.plotly_chart(fig_line, use_container_width=True)
    else:
        st.info("Grafik performa SLA tidak dapat dibuat (data kosong atau belum ada tiket yang ditutup).")

    st.divider()
    
    st.subheader("Occurrence Analysis by Category")
    
    view_mode = st.radio(
        "Pilih Tampilan Jumlah Data:",
        ["Top 3", "All"],
        horizontal=True,
        key="view_mode_selector"
    )
    
    limit_val = 3 if view_mode == "Top 3" else None
    
    possible_kategori_cols = ['Kategori', 'Category', 'Item', 'Tipe']
    kategori_col = find_column(df_inc_all.columns, possible_kategori_cols)
    item_col = find_column(df_req_all.columns, possible_kategori_cols)

    st.markdown("<h4>Incident Analysis</h4>", unsafe_allow_html=True)
    if not kategori_col:
        st.error("Kolom 'Kategori' tidak ditemukan di file Incident.")
    elif not service_col:
        st.error("Kolom 'Service Offering' tidak ditemukan.")
    elif inc_df_slice.empty:
            st.warning(f"Tidak ada data Insiden untuk periode: {time_filter_selection}")
    else:
        html_incident = display_occurrence_table(
            df_slice=inc_df_slice, 
            data_col=service_col, 
            group_by_col=kategori_col, 
            static_type=None,
            limit=limit_val 
        )
        st.markdown(html_incident, unsafe_allow_html=True)

    st.markdown("<h4>Request Analysis</h4>", unsafe_allow_html=True)
    if not item_col:
        st.error("Kolom 'Kategori' atau 'Item' tidak ditemukan di file Request.")
    elif req_df_slice.empty:
        st.warning(f"Tidak ada data Request untuk periode: {time_filter_selection}")
    else:
        html_request = display_occurrence_table(
            df_slice=req_df_slice,
            data_col=item_col,       
            group_by_col=item_col,   
            static_type="Request",
            limit=limit_val 
        )
        st.markdown(html_request, unsafe_allow_html=True)

    st.subheader("Solved vs Active/Pending Status")

    if not kategori_col:
        st.error("Kolom 'Kategori' tidak ditemukan.")
    elif not inc_res_col or not req_res_col:
        st.error("Kolom 'Resolved' / 'Tiket Ditutup' tidak ditemukan.")
    elif inc_df_slice.empty and req_df_slice.empty:
        st.warning(f"Tidak ada data untuk ditampilkan.")
    else:
        solved_data = {}
        active_data = {}
        
        incident_types = []
        if not inc_df_slice.empty:
            incident_types = sorted(list(inc_df_slice[kategori_col].dropna().unique()))
        
        all_table_cols = incident_types + ["Request"]
        
        if not inc_df_slice.empty:
            for type_name in incident_types:
                type_mask = (inc_df_slice[kategori_col] == type_name)
                solved_count = inc_df_slice[type_mask & inc_df_slice[inc_res_col].notna()].shape[0]
                active_count = inc_df_slice[type_mask & inc_df_slice[inc_res_col].isna()].shape[0]
                
                solved_data[type_name] = solved_count
                active_data[type_name] = active_count
        
        req_solved_count = 0
        req_active_count = 0
        if not req_df_slice.empty:
            req_solved_count = req_df_slice[req_df_slice[req_res_col].notna()].shape[0]
            req_active_count = req_df_slice[req_df_slice[req_res_col].isna()].shape[0]
            
        solved_data["Request"] = req_solved_count
        active_data["Request"] = req_active_count
        
        final_table_data = [solved_data, active_data]
        df_status = pd.DataFrame(final_table_data, index=["Solved", "Active/Pending"])
        
        df_status = df_status.reindex(columns=all_table_cols, fill_value=0)
        
        st.markdown(make_simple_html_table(df_status), unsafe_allow_html=True)

if __name__ == "__main__":
    run()