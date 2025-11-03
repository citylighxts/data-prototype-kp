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
    """Convert DataFrame to Excel bytes for download."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

def run():
    st.markdown(
        """
        <h1 style="display: flex; align-items: center; gap: 10px;">
            <img src="https://cdnjs.cloudflare.com/ajax/libs/twemoji/14.0.2/svg/1f4ca.svg" 
                 width="40" height="40">
            Request Item
        </h1>
        """,
        unsafe_allow_html=True
    )

    st.write("Upload an Excel file. The system will create columns: `Businesscriticality-Severity`, `Target SLA (jam)`, `Target Selesai`, and `SLA`.")

    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"], key="reqitem_uploader")
    if not uploaded_file:
        st.info("Please upload the Excel file first.")
        return

    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Failed to read Excel file: {e}")
        return

    st.success("File has been successfully read.")
    st.subheader("Data Preview")
    st.dataframe(df.head(10))

    # === Tambahan: Filter Regional & Tanggal ===
    possible_name_cols = ['Name', 'Nama', 'User Name']
    name_col = next((c for c in possible_name_cols if c in df.columns), None)

    if name_col:
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


        # --- Filter Regional ---
        regional_option = st.radio(
            "Choose Regional:",
            options=["All", "Regional 3"],
            horizontal=True,
            key="regional_filter"
        )

        if regional_option == "Regional 3":
            df = df[df[name_col].astype(str).str.contains("Regional 3", case=False, na=False)]

        # --- Filter Bulan & Tahun ---
        possible_date_cols = ['Tiket Dibuat', 'Tiket dibuat', 'Created', 'Created Date', 'CreatedAt']
        date_created_col = next((c for c in possible_date_cols if c in df.columns), None)

        if date_created_col:
            df[date_created_col] = pd.to_datetime(df[date_created_col], errors='coerce')

            # ambil semua kombinasi bulan-tahun yang ada
            df['Bulan-Tahun'] = df[date_created_col].dt.strftime('%Y-%m')
            available_months = sorted(df['Bulan-Tahun'].dropna().unique())

            selected_month = st.selectbox("Choose Month & Year:", options=["All"] + available_months, key="month_filter")

            if selected_month != "All":
                df = df[df['Bulan-Tahun'] == selected_month]

        st.markdown(f"**Data after filter:** {len(df)} baris")
        st.dataframe(df.head(7))

    possible_bc_cols = ['Businesscriticality', 'Business criticality', 'Business Criticality', 'BusinessCriticality']
    possible_sev_cols = ['Severity', 'severity', 'SEVERITY']

    bc_col = next((c for c in possible_bc_cols if c in df.columns), None)
    sev_col = next((c for c in possible_sev_cols if c in df.columns), None)

    if bc_col is None or sev_col is None:
        st.error("Column 'Businesscriticality' or 'Severity' not found.")
        st.write("Kolom yang ada di file:", list(df.columns))
        return

    # === Mapping SLA dalam JAM ===
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
    mapping_df = pd.DataFrame(list(sla_mapping_hours.items()), columns=["Businesscriticality-Severity", "Target SLA (jam)"])
    st.markdown(
        mapping_df.to_html(index=False, classes='table table-sm', justify='left')
        .replace('<td>', '<td style="text-align:left;">'),
        unsafe_allow_html=True
    )

    # === Normalisasi label & mapping ===
    df['_bc_raw'] = df[bc_col].astype(str).fillna('').str.strip()
    df['_sev_raw'] = df[sev_col].astype(str).fillna('').str.strip()
    df['Businesscriticality-Severity'] = (df['_bc_raw'] + " - " + df['_sev_raw']).apply(normalize_label)

    def map_to_hours(label: str):
        if not label:
            return None
        for key in sla_mapping_hours.keys():
            if normalize_label(label) == normalize_label(key):
                return sla_mapping_hours[key]
        return None

    df['Target SLA (jam)'] = df['Businesscriticality-Severity'].apply(map_to_hours)

    # cek kolom tanggal
    date_created_col = next((c for c in ['Tiket Dibuat', 'Tiket dibuat', 'Created', 'Created Date', 'CreatedAt'] if c in df.columns), None)
    if date_created_col is None:
        st.error("Ticket creation date column not found.")
        return

    date_closed_col = next((c for c in ['Tiket Ditutup', 'Resolved', 'Closed', 'Closed At', 'Tiket ditutup'] if c in df.columns), None)

    df[date_created_col] = pd.to_datetime(df[date_created_col], errors='coerce')
    if date_closed_col:
        df[date_closed_col] = pd.to_datetime(df[date_closed_col], errors='coerce')

    df['Target Selesai'] = df.apply(
        lambda r: (r[date_created_col] + pd.to_timedelta(r['Target SLA (jam)'], unit='h'))
        if pd.notna(r[date_created_col]) and pd.notna(r['Target SLA (jam)']) else pd.NaT,
        axis=1
    )

    if date_closed_col:
        df['SLA'] = df.apply(
            lambda r: 1 if (pd.notna(r['Target Selesai']) and pd.notna(r[date_closed_col]) and r[date_closed_col] <= r['Target Selesai'])
            else (0 if (pd.notna(r['Target Selesai']) and pd.notna(r[date_closed_col])) else pd.NA),
            axis=1
        )
    else:
        df['SLA'] = pd.NA

    # === Rekapitulasi SLA (donut chart di sini) ===
    if date_closed_col:
        sla_tercapai = int((df['SLA'] == 1).sum())
        sla_tidak_tercapai = int((df['SLA'] == 0).sum())
        sla_open = int(df[date_closed_col].isna().sum())
        total_semua = len(df)

        st.subheader("✨ SLA Recap")
        st.write(f"- 🏆 SLA achieved: **{sla_tercapai}**")
        st.write(f"- 🚨 SLA not achieved: **{sla_tidak_tercapai}**")
        st.write(f"- ⏳ SLA still open: **{sla_open}**")
        st.write(f"- 📈 Total tickets: **{total_semua}**")

        rekap_data = pd.DataFrame({
            'Kategori': ['SLA Achieved', 'SLA Not Achieved', 'Open'],
            'Jumlah': [sla_tercapai, sla_tidak_tercapai, sla_open]
        })
        fig2 = px.bar(rekap_data, x='Kategori', y='Jumlah', text='Jumlah', title='SLA Recap')
        fig2.update_traces(textposition='outside')
        st.plotly_chart(fig2)

        if sla_tercapai + sla_tidak_tercapai > 0:
            donut_data = pd.DataFrame({
                "SLA": ["On Time", "Late"],
                "Jumlah": [sla_tercapai, sla_tidak_tercapai]
            })
            fig_donut = px.pie(donut_data, names="SLA", values="Jumlah", hole=0.4, title="On Time Percentage")
            st.plotly_chart(fig_donut)

    # === Analisis tambahan ===
    st.subheader("✨ Additional Analysis")

    # top 5 kombinasi
    top5 = df['Businesscriticality-Severity'].value_counts().reset_index()
    top5.columns = ['Businesscriticality-Severity', 'Jumlah']
    top5.index = top5.index + 1  # mulai dari 1
    top5 = top5.head(5)
    st.markdown("**Top 5 most frequent Businesscriticality-Severity combinations:**")
    st.markdown(top5.to_html(index=True, justify='left').replace('<td>', '<td style="text-align:left;">'), unsafe_allow_html=True)

    # === Contact Type & Item ===
    possible_contacttype_cols = ['Contact Type', 'ContactType', 'Contact type']
    contact_col = next((c for c in possible_contacttype_cols if c in df.columns), None)

    if contact_col:
        contact_summary = df[contact_col].value_counts(dropna=False).reset_index()
        contact_summary.columns = ['Tipe Kontak', 'Jumlah']
        contact_summary.index = contact_summary.index + 1
        top3_contact = contact_summary.head(3)

        st.subheader("✨ Contact Type Analysis")
        st.markdown("**Top 3 contact type analysis**")
        st.markdown(top3_contact.to_html(index=True, justify='left').replace('<td>', '<td style="text-align:left;">'), unsafe_allow_html=True)

        fig_contact = px.pie(top3_contact, names='Tipe Kontak', values='Jumlah', hole=0.4, title='Top 3 Contact Type')
        st.plotly_chart(fig_contact)
        st.markdown("<p style='font-style: italic; color: gray;'>Showing top 3 contact types.</p>", unsafe_allow_html=True)

        possible_item_cols = ['Item', 'item', 'ITEM']
        item_col = next((c for c in possible_item_cols if c in df.columns), None)

        if item_col:
            top5_item = df[item_col].value_counts(dropna=False).reset_index()
            top5_item.columns = ['Item', 'Jumlah']
            top5_item.index = top5_item.index + 1
            top5_item = top5_item.head(5)

            st.subheader("✨ Top 5 Most Requested Items")
            st.markdown(top5_item.to_html(index=True, justify='left').replace('<td>', '<td style="text-align:left;">'), unsafe_allow_html=True)

            fig_item = px.bar(
                top5_item,
                x='Item',
                y='Jumlah',
                text='Jumlah',
                title='Top 5 Most Requested Items',
            )
            fig_item.update_traces(textposition='outside')
            fig_item.update_layout(yaxis_range=[0, top5_item['Jumlah'].max() * 1.2])
            st.plotly_chart(fig_item)

        # === Analisis Service Offering ===
        possible_service_cols = ['Service Offering', 'ServiceOffering', 'Service offering']
        service_col = next((c for c in possible_service_cols if c in df.columns), None)

        if service_col:
            st.subheader("✨ Service Offering Analysis")
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

            st.markdown("**Top 3 Service Offerings with the most tickets:**")
            st.markdown(top3_service.to_html(index=True, justify='left').replace('<td>', '<td style="text-align:left;">'), unsafe_allow_html=True)


            # Visualisasi
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

            # Catatan di bawah grafik
            st.markdown("<p style='font-style: italic; color: gray;'", unsafe_allow_html=True)


    # tampilkan hasil akhir
    def sla_status(row):
        if pd.isna(row.get(date_closed_col)):
            return "Open"
        elif row.get("SLA") == 1:
            return "Achieved"
        elif row.get("SLA") == 0:
            return "Not Achieved"
        else:
            return "Unknown"

    df["Status SLA"] = df.apply(sla_status, axis=1)

    st.subheader("🔥 Calculation Result")
    tiket_col = next((c for c in ['No. Tiket', 'Ticket No', 'No Ticket', 'No Tiket', 'Ticket'] if c in df.columns), None)
    show_cols = [col for col in [tiket_col, 'Businesscriticality-Severity', 'Target SLA (jam)', 'Target Selesai', 'SLA'] if col]
    st.dataframe(df[show_cols].head(50))

    # Ganti ke XLSX download
    excel_bytes = to_excel(df)
    st.download_button(
        "Download Result XLSX",
        data=excel_bytes,
        file_name="reqitem_hasil.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )