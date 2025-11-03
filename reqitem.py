import streamlit as st
import pandas as pd
import plotly.express as px
import re

def normalize_label(s: str) -> str:
    if pd.isna(s):
        return ""
    # ubah ke string, hilangkan leading/trailing space
    s = str(s).strip()
    # ganti banyak whitespace dengan single space
    s = re.sub(r'\s+', ' ', s)
    # sisipkan spasi di sekitar tanda '-'
    s = re.sub(r'\s*-\s*', ' - ', s)
    # jika format seperti "3 - Medium3 - Low" jadi "3 - Medium - 3 - Low" (pisahkan angka+kata jika menyatu)
    # insert space between digit and letter (digit followed by letter) and vice versa if diperlukan
    s = re.sub(r'(?<=\d)(?=[A-Za-z])', ' ', s)
    s = re.sub(r'(?<=\D)(?=\d)', ' ', s)
    # replace multiple spaces again
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def run():
    st.title("Request Item")
    st.write("Unggah file Excel. Sistem akan membuat kolom: `Businesscriticality-Severity`, `Target SLA (jam)`, `Target Selesai`, `SLA`.")

    uploaded_file = st.file_uploader("Upload file Excel", type=["xlsx", "xls"], key="reqitem_uploader")
    if not uploaded_file:
        st.info("Silakan upload file Excel terlebih dahulu.")
        return

    # baca file
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Gagal membaca file Excel: {e}")
        return

    st.success("File berhasil dibaca.")
    st.subheader("Preview data (top 10)")
    st.dataframe(df.head(10))

    # cek nama kolom untuk Businesscriticality & Severity (beberapa kemungkinan)
    possible_bc_cols = ['Businesscriticality', 'Business criticality', 'Business Criticality', 'BusinessCriticality']
    possible_sev_cols = ['Severity', 'severity', 'SEVERITY']

    bc_col = next((c for c in possible_bc_cols if c in df.columns), None)
    sev_col = next((c for c in possible_sev_cols if c in df.columns), None)

    if bc_col is None or sev_col is None:
        st.error("Kolom 'Businesscriticality' atau 'Severity' tidak ditemukan. Cek nama kolom di file Excel.")
        st.write("Kolom yang ada di file:", list(df.columns))
        return

    # --- mapping SLA dalam satuan JAM (hours) untuk menghindari kebingungan hari/float ---
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

    # buat kolom normalized untuk mempermudah mapping
    df['_bc_raw'] = df[bc_col].astype(str).fillna('').str.strip()
    df['_sev_raw'] = df[sev_col].astype(str).fillna('').str.strip()

    # gabungkan lalu normalisasi
    df['Businesscriticality-Severity'] = (df['_bc_raw'] + " - " + df['_sev_raw']).apply(normalize_label)

    # Tampilkan beberapa contoh hasil normalisasi agar kamu bisa cek
    st.write("Contoh kombinasi `Businesscriticality-Severity` setelah normalisasi:")
    st.dataframe(df['Businesscriticality-Severity'].head(10))

    # mapping: karena dataset bisa punya format lain, kita coba beberapa varian:
    def map_to_hours(label: str):
        if not label:
            return None
        # coba langsung
        if label in sla_mapping_hours:
            return sla_mapping_hours[label]
        # kadang dataset punya format tanpa spasi: coba normalisasi alternatif
        alt = label.replace(' - ', '-')
        if alt in sla_mapping_hours:
            return sla_mapping_hours[alt]
        # coba compress spaces
        alt2 = re.sub(r'\s+', ' ', label)
        if alt2 in sla_mapping_hours:
            return sla_mapping_hours[alt2]
        # coba ganti single dash no spaces
        alt3 = label.replace(' - ', ' - ').strip()
        if alt3 in sla_mapping_hours:
            return sla_mapping_hours[alt3]
        # terakhir: coba hapus semua spasi di antara parts (jelaskan jika perlu)
        alt4 = label.replace(' ', '')
        if alt4 in sla_mapping_hours:
            return sla_mapping_hours[alt4]
        return None

    df['Target SLA (jam)'] = df['Businesscriticality-Severity'].apply(map_to_hours)

    # tampilkan kombinasi yang gagal dimapping (untuk debug)
    unmapped = df[df['Target SLA (jam)'].isna()]['Businesscriticality-Severity'].drop_duplicates().tolist()
    if len(unmapped) > 0:
        st.warning(f"Ada kombinasi Businesscriticality-Severity yang TIDAK TERMAPPING ({len(unmapped)}):")
        st.write(unmapped[:30])  # tampilkan hingga 30 unik

    # jika ada mapping yang null, jangan otomatis jadi 0 — biarkan null atau beri nilai default
    # kita lanjutkan perhitungan hanya untuk baris yang punya Target SLA
    # pastikan kolom tanggal ada (Tiket Dibuat)
    date_created_col = None
    for c in ['Tiket Dibuat', 'Tiket dibuat', 'Created', 'Created Date', 'CreatedAt']:
        if c in df.columns:
            date_created_col = c
            break
    if date_created_col is None:
        st.error("Kolom tanggal pembuatan tiket (mis. 'Tiket Dibuat') tidak ditemukan.")
        return

    # coba cari kolom closed/resolved / tiket ditutup
    date_closed_col = None
    for c in ['Tiket Ditutup', 'Resolved', 'Closed', 'Closed At', 'Tiket ditutup']:
        if c in df.columns:
            date_closed_col = c
            break
    if date_closed_col is None:
        st.warning("Kolom tanggal penutupan/Resolved tidak ditemukan. SLA tidak dapat dihitung tanpa tanggal penutupan.")
        # kita tetap buat kolom Target Selesai tapi SLA akan NaN
    # konversi tanggal
    df[date_created_col] = pd.to_datetime(df[date_created_col], errors='coerce')
    if date_closed_col:
        df[date_closed_col] = pd.to_datetime(df[date_closed_col], errors='coerce')

    # hitung Target Selesai = Tiket Dibuat + Target SLA (jam)
    df['Target Selesai'] = df.apply(
        lambda r: (r[date_created_col] + pd.to_timedelta(r['Target SLA (jam)'], unit='h')) if pd.notna(r[date_created_col]) and pd.notna(r['Target SLA (jam)']) else pd.NaT,
        axis=1
    )

    # Hitung SLA jika ada tanggal closed
    if date_closed_col:
        df['SLA'] = df.apply(
            lambda r: 1 if (pd.notna(r['Target Selesai']) and pd.notna(r[date_closed_col]) and r[date_closed_col] <= r['Target Selesai']) else (0 if (pd.notna(r['Target Selesai']) and pd.notna(r[date_closed_col])) else pd.NA),
            axis=1
        )
    else:
        df['SLA'] = pd.NA

    # Statistik ringkas (hitung hanya baris yang punya SLA 0/1)
    if df['SLA'].notna().any():
        total = df['SLA'].dropna().shape[0]
        ontime = int((df['SLA'] == 1).sum())
        late = int((df['SLA'] == 0).sum())
        st.subheader("Statistik SLA")
        st.metric("Total baris dengan SLA dihitung", total)
        st.metric("Tepat waktu (1)", ontime)
        st.metric("Terlambat (0)", late)
        st.write(f"Persentase on time: {ontime/total*100:.2f}%")
        # donut chart
        sla_counts = df['SLA'].value_counts().rename(index={1:'On Time', 0:'Late'}).reset_index()
        sla_counts.columns = ['SLA', 'Jumlah']
        fig = px.pie(sla_counts, names='SLA', values='Jumlah', hole=0.4)
        st.plotly_chart(fig)
    else:
        st.info("Belum ada baris yang SLA-nya bisa dihitung (periksa 'Tiket Ditutup' dan mapping).")
    # ==========================================================
    # 🔹 Tambahan: Analisis tambahan & ringkasan akhir
    # ==========================================================

    st.subheader("📈 Analisis Tambahan")

    # --- Top 5 kombinasi Businesscriticality-Severity terbanyak ---
    top5 = (
        df['Businesscriticality-Severity']
        .value_counts()
        .reset_index()
        .rename(columns={'index': 'Businesscriticality-Severity', 'Businesscriticality-Severity': 'Jumlah'})
        .head(5)
    )
    st.markdown("**Top 5 kombinasi Businesscriticality-Severity terbanyak:**")
    st.dataframe(top5)

    # --- Hitung SLA tercapai / tidak tercapai / open ---
    if date_closed_col:
        sla_tercapai = int((df['SLA'] == 1).sum())
        sla_tidak_tercapai = int((df['SLA'] == 0).sum())
        sla_open = int(df[date_closed_col].isna().sum())
        total_semua = len(df)

        st.markdown("**📊 Rekapitulasi SLA:**")
        st.write(f"- ✅ SLA tercapai: **{sla_tercapai}**")
        st.write(f"- ❌ SLA tidak tercapai: **{sla_tidak_tercapai}**")
        st.write(f"- ⏳ SLA masih open (belum ditutup): **{sla_open}**")
        st.write(f"- 🧮 Total tiket: **{total_semua}**")

        # Visualisasi bar chart ringkas
        rekap_data = pd.DataFrame({
            'Kategori': ['SLA Tercapai', 'SLA Tidak Tercapai', 'Open'],
            'Jumlah': [sla_tercapai, sla_tidak_tercapai, sla_open]
        })
        fig2 = px.bar(rekap_data, x='Kategori', y='Jumlah', text='Jumlah', title='Rekapitulasi SLA')
        st.plotly_chart(fig2)
    else:
        st.info("Kolom tanggal penutupan tidak ditemukan, tidak bisa menghitung status open SLA.")
    def sla_status(row):
        if pd.isna(row.get(date_closed_col)):
            return "Open"
        elif row.get("SLA") == 1:
            return "Tercapai"
        elif row.get("SLA") == 0:
            return "Tidak Tercapai"
        else:
            return "Unknown"
        
    # ==========================================================
    # 🔹 Tambahan Visualisasi Contact Type & Item
    # ==========================================================
    st.subheader("📞 Analisis Contact Type")

    # Cek apakah kolom contact type ada
    possible_contact_cols = ['Contact Type', 'ContactType', 'Contact type']
    contact_col = next((c for c in possible_contact_cols if c in df.columns), None)

    if contact_col:
        # Hitung jumlah tiap contact type
        contact_summary = df[contact_col].value_counts(dropna=False).reset_index()
        contact_summary.columns = ['Tipe Kontak', 'Jumlah']

        # Pastikan kolom 'Jumlah' numerik
        contact_summary['Jumlah'] = pd.to_numeric(contact_summary['Jumlah'], errors='coerce').fillna(0)

        # Hitung persentase
        total_jumlah = contact_summary['Jumlah'].sum()
        if total_jumlah > 0:
            contact_summary['Persentase'] = (contact_summary['Jumlah'] / total_jumlah * 100).round(2)
        else:
            contact_summary['Persentase'] = 0

        st.markdown("**Persentase penggunaan Contact Type:**")
        st.dataframe(contact_summary)

        # Pie chart persentase Contact Type
        fig_contact = px.pie(
            contact_summary,
            names='Tipe Kontak',
            values='Jumlah',
            hole=0.4,
            title='Distribusi Contact Type (dalam Persentase)',
        )
        st.plotly_chart(fig_contact)

        # ==========================================================
        # 🔹 Top 5 Item terbanyak
        # ==========================================================
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
            st.plotly_chart(fig_item)
        else:
            st.info("Kolom 'Item' tidak ditemukan di file Excel.")

    else:
        st.info("Kolom 'Contact Type' tidak ditemukan di file Excel.")


    df["Status SLA"] = df.apply(sla_status, axis=1)
    # tampilkan hasil (kolom yang diminta)
    show_cols = []
    # kolom tiket id (cari nama yg ada)
    tiket_col = next((c for c in ['No. Tiket', 'Ticket No', 'No Ticket', 'No Tiket', 'Ticket'] if c in df.columns), None)
    if tiket_col:
        show_cols.append(tiket_col)
    show_cols += ['Businesscriticality-Severity', 'Target SLA (jam)', 'Target Selesai', 'SLA']
    st.subheader("Hasil perhitungan")
    st.dataframe(df[show_cols].head(50))

    # Download hasil
    st.subheader("Download hasil")
    csv = df.to_csv(index=False).encode('utf-8')
    st.download_button("Download CSV hasil", data=csv, file_name="reqitem_hasil.csv", mime="text/csv")
