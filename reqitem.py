import streamlit as st
import pandas as pd
import plotly.express as px
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

def run():
    st.title("Request Item")
    st.write("Unggah file Excel. Sistem akan membuat kolom: `Businesscriticality-Severity`, `Target SLA (jam)`, `Target Selesai`, dan `SLA`.")

    uploaded_file = st.file_uploader("Upload file Excel", type=["xlsx", "xls"], key="reqitem_uploader")
    if not uploaded_file:
        st.info("Silakan upload file Excel terlebih dahulu.")
        return

    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Gagal membaca file Excel: {e}")
        return

    st.success("File berhasil dibaca.")
    st.subheader("Preview data")
    st.dataframe(df.head(10))

    possible_bc_cols = ['Businesscriticality', 'Business criticality', 'Business Criticality', 'BusinessCriticality']
    possible_sev_cols = ['Severity', 'severity', 'SEVERITY']

    bc_col = next((c for c in possible_bc_cols if c in df.columns), None)
    sev_col = next((c for c in possible_sev_cols if c in df.columns), None)

    if bc_col is None or sev_col is None:
        st.error("Kolom 'Businesscriticality' atau 'Severity' tidak ditemukan.")
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

    st.subheader("📊 Mapping SLA (jam)")
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
        st.error("Kolom tanggal pembuatan tiket tidak ditemukan.")
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

        st.subheader("📊 Rekapitulasi SLA")
        st.write(f"- ✅ SLA tercapai: **{sla_tercapai}**")
        st.write(f"- ❌ SLA tidak tercapai: **{sla_tidak_tercapai}**")
        st.write(f"- ⏳ SLA masih open: **{sla_open}**")
        st.write(f"- 🧮 Total tiket: **{total_semua}**")

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
            fig_donut = px.pie(donut_data, names="SLA", values="Jumlah", hole=0.4, title="Persentase On Time")
            st.plotly_chart(fig_donut)

    # === Analisis tambahan ===
    st.subheader("📈 Analisis Tambahan")

    # top 5 kombinasi
    top5 = (
        df['Businesscriticality-Severity']
        .value_counts()
        .reset_index()
        .rename(columns={'index': 'Businesscriticality-Severity', 'Businesscriticality-Severity': 'Jumlah'})
        .head(5)
    )
    st.markdown("**Top 5 kombinasi Businesscriticality-Severity terbanyak:**")
    st.dataframe(top5)

    # === Contact Type & Item ===
    possible_contact_cols = ['Contact Type', 'ContactType', 'Contact type']
    contact_col = next((c for c in possible_contact_cols if c in df.columns), None)

    if contact_col:
        contact_summary = df[contact_col].value_counts(dropna=False).reset_index()
        contact_summary.columns = ['Tipe Kontak', 'Jumlah']
        top3_contact = contact_summary.head(3)

        st.subheader("📞 Analisis Contact Type")
        st.dataframe(top3_contact)
        fig_contact = px.pie(top3_contact, names='Tipe Kontak', values='Jumlah', hole=0.4, title='Top 3 Contact Type')
        st.plotly_chart(fig_contact)
        st.markdown("<p style='font-style: italic; color: gray;'>Menampilkan 3 Contact Type teratas.</p>", unsafe_allow_html=True)

        possible_item_cols = ['Item', 'item', 'ITEM']
        item_col = next((c for c in possible_item_cols if c in df.columns), None)

        if item_col:
            top5_item = df[item_col].value_counts(dropna=False).reset_index()
            top5_item.columns = ['Item', 'Jumlah']
            top5_item = top5_item.head(5)

            st.subheader("🧾 Top 5 Item Terbanyak")
            st.dataframe(top5_item)

            fig_item = px.bar(
                top5_item,
                x='Item',
                y='Jumlah',
                text='Jumlah',
                title='Top 5 Item Terbanyak',
            )
            fig_item.update_traces(textposition='outside')
            fig_item.update_layout(yaxis_range=[0, top5_item['Jumlah'].max() * 1.2])
            st.plotly_chart(fig_item)

    # tampilkan hasil akhir
    def sla_status(row):
        if pd.isna(row.get(date_closed_col)):
            return "Open"
        elif row.get("SLA") == 1:
            return "Tercapai"
        elif row.get("SLA") == 0:
            return "Tidak Tercapai"
        else:
            return "Unknown"

    df["Status SLA"] = df.apply(sla_status, axis=1)

    tiket_col = next((c for c in ['No. Tiket', 'Ticket No', 'No Ticket', 'No Tiket', 'Ticket'] if c in df.columns), None)
    show_cols = [col for col in [tiket_col, 'Businesscriticality-Severity', 'Target SLA (jam)', 'Target Selesai', 'SLA'] if col]
    st.subheader("Hasil Perhitungan")
    st.dataframe(df[show_cols].head(50))

    csv = df.to_csv(index=False).encode('utf-8')
    st.download_button("Download CSV hasil", data=csv, file_name="reqitem_hasil.csv", mime="text/csv")
