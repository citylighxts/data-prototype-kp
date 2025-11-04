import streamlit as st
import pandas as pd
import plotly.express as px
import io
import re

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

def run():
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

    st.write("Upload file Excel Insiden. Sistem akan menambahkan kolom: `Business criticality-Severity`, `Waktu SLA`, `Target Selesai Baru`, `SLA`, dan `Time Breach`.")
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

    possible_loc_cols = ['Lokasi Pelapor', 'Name', 'User Name', 'Lokasi']
    loc_col = next((c for c in possible_loc_cols if c in df.columns), None)

    if loc_col:
        st.markdown(
            """
            <h1 style="display: flex; align-items: center; gap: 10px;">
                <img src="https://cdn.jsdelivr.net/gh/twitter/twemoji@14.0.2/assets/svg/1f9ea.svg" 
                    width="40" height="40">
                Data Filter
            </h1>
            """,
            unsafe_allow_html=True
        )

        regional_option = st.radio(
            "Pilih Lokasi/Regional:",
            options=["All", "Regional 3"],
            horizontal=True,
            key="regional_filter_incident"
        )

        if regional_option == "Regional 3":
            df = df[df[loc_col].astype(str).str.contains("Regional 3", case=False, na=False)]

        possible_date_cols = ['Tiket Dibuat', 'Tiket dibuat', 'Created', 'Created Date', 'CreatedAt']
        date_created_col = next((c for c in possible_date_cols if c in df.columns), None)

        if date_created_col:
            df[date_created_col] = pd.to_datetime(df[date_created_col], errors='coerce')
            df['Bulan-Tahun'] = df[date_created_col].dt.strftime('%Y-%m')
            available_months = sorted(df['Bulan-Tahun'].dropna().unique())
            selected_month = st.selectbox("Pilih Bulan & Tahun:", options=["All"] + available_months, key="month_filter_incident")
            if selected_month != "All":
                df = df[df['Bulan-Tahun'] == selected_month]

        st.markdown(f"**Data setelah filter:** {len(df)} baris")
        st.dataframe(df.head(7))

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

    def calculate_time_breach(row):
        sla_val = row.get('SLA')
        if pd.isna(sla_val) or sla_val != 0:
            return pd.NA
        if pd.notna(row[date_resolved_col]) and pd.notna(row[date_created_col]) and pd.notna(row['Waktu SLA']):
            resolution_duration = row[date_resolved_col] - row[date_created_col]
            sla_timedelta = pd.to_timedelta(row['Waktu SLA'], unit='h')
            breach = resolution_duration - sla_timedelta
            return breach.total_seconds() / 3600
        else:
            return pd.NA

    df['Time Breach'] = df.apply(calculate_time_breach, axis=1)

    sla_tercapai = int((df['SLA'] == 1).sum())
    sla_tidak_tercapai = int((df['SLA'] == 0).sum())
    if date_resolved_col:
        sla_open = int(df[date_resolved_col].isna().sum())
    else:
        sla_open = len(df)
    total_semua = len(df)

    st.subheader("✨ Rekapitulasi SLA (Berdasarkan Kalkulasi Baru)")
    st.write(f"- 🏆 SLA tercapai: **{sla_tercapai}**")
    st.write(f"- 🚨 SLA tidak tercapai: **{sla_tidak_tercapai}**")
    st.write(f"- ⏳ Tiket masih open: **{sla_open}**")
    st.write(f"- 📈 Total tiket: **{total_semua}**")

    rekap_data = pd.DataFrame({
        'Kategori': ['SLA Tercapai', 'SLA Tidak Tercapai', 'Open'],
        'Jumlah': [sla_tercapai, sla_tidak_tercapai, sla_open]
    })
    fig2 = px.bar(rekap_data, x='Kategori', y='Jumlah', text='Jumlah', title='Rekapitulasi SLA')
    fig2.update_traces(textposition='outside')
    st.plotly_chart(fig2)

    if sla_tercapai + sla_tidak_tercapai > 0:
        donut_data = pd.DataFrame({
            "SLA": ["On Time", "Late"],
            "Jumlah": [sla_tercapai, sla_tidak_tercapai]
        })
        fig_donut = px.pie(donut_data, names="SLA", values="Jumlah", hole=0.4, title="Persentase On Time (dari tiket yang sudah ditutup)")
        st.plotly_chart(fig_donut)

    st.subheader("✨ Analisis Tambahan")

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
        top3_contact = contact_summary.head(3)
        st.subheader("✨ Analisis Channel")
        st.markdown("**Top 3 Channel**")
        st.markdown(top3_contact.to_html(index=True, justify='left').replace('<td>', '<td style="text-align:left;">'), unsafe_allow_html=True)
        fig_contact = px.pie(top3_contact, names='Channel', values='Jumlah', hole=0.4, title='Top 3 Channel')
        st.plotly_chart(fig_contact)

    possible_service_cols = ['Service offering', 'Service Offering', 'ServiceOffering', 'Service offering']
    service_col = next((c for c in possible_service_cols if c in df.columns), None)

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
        top3_service = service_summary.head(3)
        st.markdown("**Top 3 Service Offering dengan tiket terbanyak:**")
        st.markdown(top3_service.to_html(index=True, justify='left').replace('<td>', '<td style="text-align:left;">'), unsafe_allow_html=True)
        fig_service = px.bar(
            top3_service,
            x='Service Offering',
            y='Jumlah',
            text='Jumlah',
            title='Top 3 Service Offering',
        )
        fig_service.update_traces(textposition='outside')
        fig_service.update_layout(
            yaxis_range=[0, top3_service['Jumlah'].max() * 1.2],
            xaxis_tickangle=-45
        )
        st.plotly_chart(fig_service)

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
