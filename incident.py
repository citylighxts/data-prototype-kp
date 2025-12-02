import streamlit as st
import pandas as pd
import plotly.express as px
import io
import re
from datetime import time, timedelta, datetime
import numpy as np
import calendar

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

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

def format_hari_jam_menit(total_hours_decimal):
    if pd.isna(total_hours_decimal) or total_hours_decimal <= 0:
        return "0 hari 0 jam 0 menit"
    
    total_hours_decimal = float(total_hours_decimal)
    total_minutes = round(total_hours_decimal * 60)
    
    days = total_minutes // (24 * 60)
    minutes_remaining = total_minutes % (24 * 60)
    hours = minutes_remaining // 60
    minutes = minutes_remaining % 60
        
    return f"{days} hari {hours} jam {minutes} menit"

def format_jam_menit_saja(total_hours_decimal):
    if pd.isna(total_hours_decimal) or total_hours_decimal <= 0:
        return "0 jam 0 menit"
    
    total_hours_decimal = float(total_hours_decimal)
    total_minutes = round(total_hours_decimal * 60)
    
    hours = total_minutes // 60
    minutes = total_minutes % 60
        
    return f"{hours} jam {minutes} menit"

def calculate_time_breach(row, date_created_col, date_resolved_col):
    sla_val = row.get('SLA')
    
    if pd.isna(sla_val):
        return pd.NA

    if pd.notna(row[date_resolved_col]) and pd.notna(row[date_created_col]) and pd.notna(row['Waktu SLA']):
        resolution_duration = row[date_resolved_col] - row[date_created_col]
        sla_timedelta = pd.to_timedelta(row['Waktu SLA'], unit='h')
        breach = resolution_duration - sla_timedelta
        return breach.total_seconds() / (3600 * 24) 
    else:
        return pd.NA
    
def run():
    st.set_page_config(layout="wide") 

    st.markdown(
        """
        <h1 style="display: flex; align-items: center; gap: 10px;">
            <img src="https://cdnjs.cloudflare.com/ajax/libs/twemoji/14.0.2/svg/1f6a8.svg" 
                 width="40" height="40">
            Incident
        </h1>
        """,
        unsafe_allow_html=True
    )

    st.write("Upload file Excel Insiden. Sistem akan menambahkan kolom: `Business criticality-Severity`, `Waktu SLA`, `Target Selesai Baru`, `SLA`, dan `Time Breach`. **Perhitungan menggunakan Waktu Kalender (24/7).**")
    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"], key="incident_uploader")
    if not uploaded_file:
        st.info("Silakan upload file Excel terlebih dahulu.")
        return

    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Gagal membaca file Excel: {e}")
        return

    st.success("File berhasil dibaca.")
    st.subheader("Data Preview")
    st.dataframe(df.head(10))

    possible_bc_cols = ['Business criticality', 'Businesscriticality', 'Business criticality', 'BusinessCriticality']
    possible_sev_cols = ['Severity', 'severity', 'SEVERITY']
    bc_col = next((c for c in possible_bc_cols if c in df.columns), None)
    sev_col = next((c for c in possible_sev_cols if c in df.columns), None)

    if bc_col is None or sev_col is None:
        st.error(f"Kolom 'Business criticality' atau 'Severity' tidak ditemukan. Tidak bisa menghitung SLA.")
        st.write("Kolom yang ada di file:", list(df.columns))
        return

    date_created_col = next((c for c in ['Tiket Dibuat', 'Tiket dibuat', 'Created', 'Created Date', 'CreatedAt'] if c in df.columns), None)
    date_resolved_col = next((c for c in ['Resolved', 'Tiket Ditutup', 'Closed', 'Closed At', 'Tiket ditutup'] if c in df.columns), None)

    if date_created_col is None:
        st.error("Kolom tanggal 'Tiket Dibuat' (atau variasinya) tidak ditemukan.")
        return

    df[date_created_col] = pd.to_datetime(df[date_created_col], errors='coerce')
    if date_resolved_col:
        df[date_resolved_col] = pd.to_datetime(df[date_resolved_col], errors='coerce')
    else:
        st.warning("Kolom 'Resolved' atau 'Tiket Ditutup' tidak ditemukan. Perhitungan SLA dan Time Breach mungkin tidak akurat.")

    total_hours_in_month = 744 
    
    first_valid_date = df[date_created_col].dropna().min()
    
    if pd.notna(first_valid_date):
        ref_year = first_valid_date.year
        ref_month = first_valid_date.month
        days_in_month = calendar.monthrange(ref_year, ref_month)[1]
        total_hours_in_month = days_in_month * 24
        
        st.info(f"Bulan terdeteksi: **{first_valid_date.strftime('%B %Y')}** ({days_in_month} hari). Total jam digunakan untuk SLA%: **{total_hours_in_month} jam**.")
    else:
        st.warning("Tidak dapat mendeteksi tanggal di 'Tiket Dibuat'. Menggunakan default 744 jam (31 hari).")

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

    st.subheader("✨ SLA Mapping (hours)")
    mapping_df = pd.DataFrame(list(sla_mapping_hours.items()), columns=["Business criticality-Severity", "Target SLA (jam)"])
    st.markdown(
        mapping_df.to_html(index=False, classes='table table-sm', justify='left')
        .replace('<td>', '<td style="text-align:left;">'),
        unsafe_allow_html=True
    )

    df['_bc_raw'] = df[bc_col].astype(str).fillna('').str.strip()
    df['_sev_raw'] = df[sev_col].astype(str).fillna('').str.strip()
    df['Business criticality-Severity'] = (df['_bc_raw'] + " - " + df['_sev_raw']).apply(normalize_label)

    def map_to_hours(label: str):
        if not label:
            return None
        norm_label = normalize_label(label)
        for key in sla_mapping_hours.keys():
            if norm_label == normalize_label(key):
                return sla_mapping_hours[key]
        return None

    df['Waktu SLA'] = df['Business criticality-Severity'].apply(map_to_hours)
    
    df['Target Selesai Baru'] = df.apply(
        lambda r: (r[date_created_col] + pd.to_timedelta(r['Waktu SLA'], unit='h'))
        if pd.notna(r[date_created_col]) and pd.notna(r['Waktu SLA']) else pd.NaT,
        axis=1
    )

    if date_resolved_col:
        df['SLA'] = df.apply(
            lambda r: 1 if (pd.notna(r['Target Selesai Baru']) and pd.notna(r[date_resolved_col]) and r[date_resolved_col] <= r['Target Selesai Baru'])
            else (0 if (pd.notna(r['Target Selesai Baru']) and pd.notna(r[date_resolved_col])) else pd.NA),
            axis=1
        )
    else:
        df['SLA'] = pd.NA

    df['Time Breach'] = df.apply(
        lambda r: calculate_time_breach(r, date_created_col, date_resolved_col), axis=1
    )

    st.dataframe(df)

    sla_tercapai = int((df['SLA'] == 1).sum())
    sla_tidak_tercapai = int((df['SLA'] == 0).sum())
    if date_resolved_col:
        sla_open = int(df[date_resolved_col].isna().sum())
    else:
        sla_open = len(df)
    total_semua = len(df)

    st.subheader("✨ Rekapitulasi SLA")
    
    col_rekap_kiri, col_rekap_kanan = st.columns([1, 1])

    with col_rekap_kiri:
        st.markdown("### Ringkasan")
        st.write(f"- 🏆 SLA tercapai: **{sla_tercapai}**")
        st.write(f"- 🚨 SLA tidak tercapai: **{sla_tidak_tercapai}**")
        st.write(f"- ⏳ Tiket masih open: **{sla_open}**")
        st.write(f"- 📈 Total tiket: **{total_semua}**")
        
        if sla_tercapai + sla_tidak_tercapai > 0:
            donut_data = pd.DataFrame({
                "SLA": ["On Time", "Late"],
                "Jumlah": [sla_tercapai, sla_tidak_tercapai]
            })
            fig_donut = px.pie(donut_data, names="SLA", values="Jumlah", hole=0.4, title="Persentase On Time (Closed Tickets)")
            fig_donut.update_layout(margin=dict(t=30, b=0, l=0, r=0), height=300)
            st.plotly_chart(fig_donut, use_container_width=True)

    with col_rekap_kanan:
        rekap_data = pd.DataFrame({
            'Kategori': ['SLA Tercapai', 'SLA Tidak Tercapai', 'Open'],
            'Jumlah': [sla_tercapai, sla_tidak_tercapai, sla_open]
        })
        fig2 = px.bar(rekap_data, x='Kategori', y='Jumlah', text='Jumlah', title='Detail Jumlah Tiket')
        fig2.update_traces(textposition='outside')
        fig2.update_layout(margin=dict(t=40, b=0, l=0, r=0), height=450)
        st.plotly_chart(fig2, use_container_width=True)

    st.divider()

    st.subheader("✨ Analisis Kombinasi Business criticality-Severity")
    if 'Business criticality-Severity' in df.columns:
        top5 = df['Business criticality-Severity'].value_counts().reset_index()
        top5.columns = ['Business criticality-Severity', 'Jumlah']
        top5.index = top5.index + 1
        top5 = top5.head(5)
        st.markdown("**Top 5 Kombinasi Business criticality-Severity:**")
        st.markdown(top5.to_html(index=True, justify='left').replace('<td>', '<td style="text-align:left;">'), unsafe_allow_html=True)

    possible_contacttype_cols = ['Channel', 'Contact Type', 'ContactType', 'Contact type']
    contact_col = next((c for c in possible_contacttype_cols if c in df.columns), None)

    if contact_col:
        contact_summary = df[contact_col].value_counts(dropna=False).reset_index()
        contact_summary.columns = ['Channel', 'Jumlah']
        contact_summary.index = contact_summary.index + 1
        st.subheader("✨ Analisis Channel")
        
        col_channel_1, col_channel_2 = st.columns([1, 1])
        with col_channel_1:
            st.markdown("**Rekapitulasi Semua Channel**")
            st.markdown(contact_summary.to_html(index=True, justify='left').replace('<td>', '<td style="text-align:left;">'), unsafe_allow_html=True)
        with col_channel_2:
            fig_contact = px.pie(contact_summary, names='Channel', values='Jumlah', hole=0.4, title='Proporsi Channel')
            fig_contact.update_layout(margin=dict(t=40, b=0, l=0, r=0))
            st.plotly_chart(fig_contact, use_container_width=True)

    possible_service_cols = ['Service offering', 'Service Offering', 'ServiceOffering', 'Service offering']
    service_col = next((c for c in possible_service_cols if c in df.columns), None)

    css_tabel = """
    <style>
        .manual-sla-table {
            width: 100%; /* FIT TO LAYOUT */
            border-collapse: collapse;
            text-align: left;
            font-size: 13px;
        }
        .manual-sla-table th, .manual-sla-table td {
            padding: 6px 10px;
            border-bottom: 1px solid #EEEEEE;
            vertical-align: middle;
            white-space: nowrap;
        }
        .manual-sla-table th {
            background-color: #F0F2F6;
            font-weight: 600;
        }
        .manual-sla-table tr:hover {
            background-color: #F5F5F5;
        }
    </style>
    """

    if service_col:
        st.subheader("✨ Analisis Service Offering")
        service_summary = (
            df[service_col]
            .dropna()
            .astype(str)
            .replace(['', 'None', 'nan', 'NaN'], pd.NA)
            .dropna()
            .value_counts()
            .reset_index()
        )
        service_summary.columns = ['Service Offering', 'Jumlah']
        service_summary.index = service_summary.index + 1
        top5_service = service_summary.head(5)

        col_table, col_chart = st.columns([1, 1]) 

        with col_table:
            st.markdown("**Top 5 Service Offering dengan tiket terbanyak:**")
            st.markdown(top5_service.to_html(index=True, justify='left').replace('<td>', '<td style="text-align:left;">'), unsafe_allow_html=True)

        with col_chart:
            fig_service = px.bar(
                top5_service,
                x='Service Offering',
                y='Jumlah',
                text='Jumlah',
                title='Top 5 Service Offering',
            )
            fig_service.update_traces(textposition='outside')
            fig_service.update_layout(
                yaxis_range=[0, top5_service['Jumlah'].max() * 1.2],
                xaxis_tickangle=-45,
                margin=dict(l=20, r=20, t=40, b=20) 
            )
            st.plotly_chart(fig_service, use_container_width=True)

    if service_col and 'SLA' in df.columns:
        tiket_col = next((c for c in ['No. Tiket', 'Ticket No', 'No Ticket', 'No Tiket', 'Ticket'] if c in df.columns), 'No. Tiket')

        sla_service_agg = (
            df.groupby(service_col)
            .agg(
                Jumlah_Tiket=('SLA', 'count'),
                Total_Waktu_Breach=('Time Breach', lambda x: x[x > 0].sum()),
                SLA_Tercapai=('SLA', lambda x: (x == 1).sum()),
                Jumlah_Tiket_Breach=('SLA', lambda x: (x == 0).sum()),
                Total_Waktu_SLA_Alokasi=('Waktu SLA', 'sum') 
            )
            .reset_index()
        )
        sla_service_agg['Total Waktu Breach (jam)'] = sla_service_agg['Total_Waktu_Breach'] * 24

        def hitung_sla_pencapaian_waktu(row, total_jam):
            total_alokasi = row['Total_Waktu_SLA_Alokasi']
            total_breach_jam = row['Total Waktu Breach (jam)'] 
            if total_alokasi <= 0: return 0.0
            pencapaian = (total_jam - total_breach_jam) / total_jam
            hasil_persen = max(0, pencapaian) * 100
            return int(hasil_persen + 0.5)

        sla_service_agg['SLA_Pencapaian_%'] = sla_service_agg.apply(
            lambda r: hitung_sla_pencapaian_waktu(r, total_hours_in_month), axis=1
        )

        def hitung_sla_breach_persen(row, total_jam):
            total_breach_jam = row['Total Waktu Breach (jam)']
            if total_jam <= 0: return 0.0
            breach_persen = (total_breach_jam / total_jam) * 100
            return int(breach_persen + 0.5)

        sla_service_agg['SLA_Breach_%'] = sla_service_agg.apply(
            lambda r: hitung_sla_breach_persen(r, total_hours_in_month), axis=1
        )

        sla_service_agg['No_Top'] = sla_service_agg['SLA_Pencapaian_%'].rank(method='dense', ascending=False).astype(int)
        top3_sla = sla_service_agg[sla_service_agg['No_Top'] <= 3].sort_values(by=['No_Top', service_col])
        top3_sla = top3_sla[['No_Top', service_col, 'Jumlah_Tiket', 'Total Waktu Breach (jam)', 'SLA_Pencapaian_%']]
        top3_sla.columns = ['No', 'Service Offering', 'Σ Tiket (Closed)', 'Total Waktu Breach (jam)', 'SLA (%)']

        sla_service_agg['No_Bottom'] = sla_service_agg['SLA_Breach_%'].rank(method='dense', ascending=False).astype(int)
        bottom3_sla = sla_service_agg[sla_service_agg['No_Bottom'] <= 3].sort_values(by=['No_Bottom', service_col])
        bottom3_sla = bottom3_sla[['No_Bottom', service_col, 'Jumlah_Tiket', 'Total Waktu Breach (jam)', 'SLA_Breach_%']] 
        bottom3_sla.columns = ['No', 'Service Offering', 'Σ Tiket (Closed)', 'Total Waktu Breach (jam)', 'SLA (%)']
        
        html_top = css_tabel + '<table class="manual-sla-table"><thead><tr>'
        html_top += "<th>No</th><th>Service Offering</th><th>Σ Tiket</th><th>Total Breach (jam)</th><th>SLA (%)</th></tr></thead><tbody>"
        for no in sorted(top3_sla['No'].unique()):
            group = top3_sla[top3_sla['No'] == no]
            rowspan = len(group)
            for i, (_, row) in enumerate(group.iterrows()):
                html_top += "<tr>"
                if i == 0: html_top += f"<td rowspan='{rowspan}'>{no}</td>"
                html_top += f"<td>{row['Service Offering']}</td><td>{row['Σ Tiket (Closed)']}</td><td>{format_hari_jam_menit(row['Total Waktu Breach (jam)'])}</td><td>{row['SLA (%)']}%</td></tr>"
        html_top += "</tbody></table>"

        html_bottom = css_tabel + '<table class="manual-sla-table"><thead><tr>'
        html_bottom += "<th>No</th><th>Service Offering</th><th>Σ Tiket</th><th>Total Breach (jam)</th><th>SLA (%)</th></tr></thead><tbody>"
        for no in sorted(bottom3_sla['No'].unique()):
            group = bottom3_sla[bottom3_sla['No'] == no]
            rowspan = len(group)
            for i, (_, row) in enumerate(group.iterrows()):
                html_bottom += "<tr>"
                if i == 0: html_bottom += f"<td rowspan='{rowspan}'>{no}</td>"
                html_bottom += f"<td>{row['Service Offering']}</td><td>{row['Σ Tiket (Closed)']}</td><td>{format_jam_menit_saja(row['Total Waktu Breach (jam)'])}</td><td>-{row['SLA (%)']}%</td></tr>" 
        html_bottom += "</tbody></table>"

        st.divider()

        st.header("SLA Evaluation")
        st.markdown("Pencapaian SLA berdasarkan Service Offering :")
        col_s1_left, col_s1_right = st.columns([1, 1])

        with col_s1_left:
            st.markdown("**Top 3 SLA Service**")
            st.markdown(html_top, unsafe_allow_html=True)
            
            st.markdown("**Bottom 3 SLA Service**")
            st.markdown(html_bottom, unsafe_allow_html=True)

        with col_s1_right:
            chart_config = {'displayModeBar': True, 'modeBarButtonsToRemove': ['zoom2d', 'pan2d', 'select2d', 'lasso2d', 'zoomIn2d', 'zoomOut2d', 'autoScale2d', 'resetScale2d'], 'displaylogo': False}
            
            if 'SLA_Pencapaian_%' in sla_service_agg.columns:
                chart_data = sla_service_agg.sort_values(by=service_col, ascending=True)
                fig_sla_percent = px.bar(chart_data, x=service_col, y='SLA_Pencapaian_%', text='SLA_Pencapaian_%', title="SLA Total per Service Offering")
                fig_sla_percent.update_traces(texttemplate='%{text}%', textposition='outside')
                y_max = chart_data['SLA_Pencapaian_%'].max()
                fig_sla_percent.update_layout(
                    height=600, 
                    yaxis_range=[0, max(100, y_max * 1.25)], 
                    xaxis=dict(fixedrange=True), 
                    yaxis=dict(fixedrange=True), 
                    dragmode=False, 
                    margin=dict(t=50, b=50, l=20, r=20)
                )
                st.plotly_chart(fig_sla_percent, use_container_width=True, config=chart_config)

        all_tickets_df = df[df['Time Breach'].notna() & df[service_col].notna()].copy()
        if not all_tickets_df.empty:
            all_tickets_df = all_tickets_df.sort_values(by='Time Breach', ascending=False)
            max_breach_df = all_tickets_df.drop_duplicates(subset=[service_col], keep='first')
            
            if 'sla_service_agg' in locals():
                max_breach_df = max_breach_df.merge(sla_service_agg[[service_col, 'Total_Waktu_SLA_Alokasi']], on=service_col, how='left')
                max_breach_df['Total_Waktu_SLA_Alokasi'] = max_breach_df['Total_Waktu_SLA_Alokasi'].fillna(0.0)
            else:
                max_breach_df['Total_Waktu_SLA_Alokasi'] = 0.0 

            def hitung_sla_max_breach(row, total_jam):
                total_alokasi = row['Total_Waktu_SLA_Alokasi'] 
                max_breach_hari = max(0, row['Time Breach']) 
                if total_alokasi <= 0: return 0.0 
                max_breach_jam = max_breach_hari * 24
                pencapaian = (total_jam - max_breach_jam) / total_jam
                hasil_persen = max(0, pencapaian) * 100
                return int(hasil_persen)

            max_breach_df['SLA Service (%)'] = max_breach_df.apply(lambda r: hitung_sla_max_breach(r, total_hours_in_month), axis=1)
            max_breach_df['Time Breach (jam)'] = max_breach_df['Time Breach'] * 24 
            max_breach_df['_breach_rank_val'] = max_breach_df['Time Breach'].clip(lower=0)

            top3_min_max_breach = max_breach_df.sort_values(by=['SLA Service (%)', '_breach_rank_val', service_col], ascending=[False, True, True])
            top3_min_max_breach['No'] = top3_min_max_breach['SLA Service (%)'].rank(method='dense', ascending=False).astype(int)
            top3_min_max_breach = top3_min_max_breach[top3_min_max_breach['No'] <= 3]
            top3_min_max_breach = top3_min_max_breach.drop(columns=['_breach_rank_val'])

            bottom3_max_breach = max_breach_df.sort_values(by='Time Breach', ascending=False)
            bottom3_max_breach['No'] = bottom3_max_breach['Time Breach'].rank(method='dense', ascending=False).astype(int)
            bottom3_max_breach = bottom3_max_breach[bottom3_max_breach['No'] <= 3]

            html_top_max = css_tabel + '<table class="manual-sla-table"><thead><tr>'
            html_top_max += "<th>No</th><th>Service Offering</th><th>Waktu Breach</th><th>SLA (%)</th></tr></thead><tbody>"
            for no in sorted(top3_min_max_breach['No'].unique()):
                group = top3_min_max_breach[top3_min_max_breach['No'] == no]
                rowspan = len(group)
                for i, (_, row) in enumerate(group.iterrows()):
                    html_top_max += "<tr>"
                    if i == 0: html_top_max += f"<td rowspan='{rowspan}'>{no}</td>"
                    html_top_max += f"<td>{row[service_col]}</td><td>{format_hari_jam_menit(row['Time Breach (jam)'])}</td><td>{row['SLA Service (%)']}%</td></tr>"
            html_top_max += "</tbody></table>"

            html_bottom_max = css_tabel + '<table class="manual-sla-table"><thead><tr>'
            html_bottom_max += "<th>No</th><th>Service Offering</th><th>Waktu Breach</th><th>No Tiket</th><th>SLA (%)</th></tr></thead><tbody>"
            for no in sorted(bottom3_max_breach['No'].unique()):
                group = bottom3_max_breach[bottom3_max_breach['No'] == no]
                rowspan = len(group)
                for i, (_, row) in enumerate(group.iterrows()):
                    html_bottom_max += "<tr>"
                    if i == 0: html_bottom_max += f"<td rowspan='{rowspan}'>{no}</td>"
                    html_bottom_max += f"<td>{row[service_col]}</td><td>{format_hari_jam_menit(row['Time Breach (jam)'])}</td><td>{row[tiket_col]}</td><td>{row['SLA Service (%)']}%</td></tr>"
            html_bottom_max += "</tbody></table>"

            st.divider()
            
            st.header("SLA Tiket Max Breach per Service Offering")
            st.markdown("Pencapaian SLA Max Breach per service offering :")

            col_s2_left, col_s2_right = st.columns([1, 1])

            with col_s2_left:
                st.markdown("**Top 3 SLA Service**")
                st.markdown(html_top_max, unsafe_allow_html=True)
                
                st.markdown("**Bottom 3 SLA Service**")
                st.markdown(html_bottom_max, unsafe_allow_html=True)

            with col_s2_right:                
                chart_data_max_breach = max_breach_df.sort_values(by=service_col, ascending=True)
                fig_max_breach = px.bar(chart_data_max_breach, x=service_col, y='SLA Service (%)', text='SLA Service (%)', title="SLA Tiket Max Breach per Service Offering")
                fig_max_breach.update_traces(texttemplate='%{text}%', textposition='outside')
                y_max_2 = chart_data_max_breach['SLA Service (%)'].max()
                fig_max_breach.update_layout(
                    height=600, 
                    yaxis_range=[0, max(100, y_max_2 * 1.25)], 
                    xaxis=dict(fixedrange=True), 
                    yaxis=dict(fixedrange=True), 
                    dragmode=False, 
                    margin=dict(t=50, b=50, l=20, r=20)
                )
                st.plotly_chart(fig_max_breach, use_container_width=True, config=chart_config)

    def sla_status(row):
        sla_val = row.get("SLA")
        if not date_resolved_col or pd.isna(row.get(date_resolved_col)):
            return "Open"
        if pd.isna(sla_val):
            return "Unknown"
        if sla_val == 1:
            return "Achieved"
        elif sla_val == 0:
            return "Not Achieved"
        return "Unknown"

    df["Status SLA"] = df.apply(sla_status, axis=1)
    st.subheader("🔥 Hasil Kalkulasi")
    tiket_col = next((c for c in ['No. Tiket', 'Ticket No', 'No Ticket', 'No Tiket', 'Ticket'] if c in df.columns), None)
    show_cols = [
        tiket_col,
        date_created_col,
        date_resolved_col,
        'Business criticality-Severity',
        'Waktu SLA',
        'Target Selesai Baru',
        'SLA',
        'Time Breach',
        'Status SLA'
    ]
    show_cols = [col for col in show_cols if col in df.columns or col in ['SLA', 'Status SLA', 'Waktu SLA', 'Target Selesai Baru', 'Time Breach', 'Business criticality-Severity']]
    st.dataframe(df[show_cols].head(50))

    excel_bytes = to_excel(df)
    st.download_button(
        "Download Hasil XLSX",
        data=excel_bytes,
        file_name="incident_hasil_kalkulasi.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    run()