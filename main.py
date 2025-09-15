import streamlit as st
import pandas as pd

st.title("📊 Upload & Tampilkan Data Excel")

# Mapping Business criticality-Severity ke Waktu SLA
sla_mapping = {
    '1 - Critical - 1 - High': '04.00',
    '1 - Critical - 2 - Medium': '06.00',
    '1 - Critical - 3 - Low': '08.00',
    '2 - High - 1 - High': '06.00',
    '2 - High - 2 - Medium': '08.00',
    '2 - High - 3 - Low': '12.00',
    '3 - Medium - 1 - High': '08.00',
    '3 - Medium - 2 - Medium': '12.00',
    '3 - Medium - 3 - Low': '16.00',
    '4 - Low - 1 - High': '16.00',
    '4 - Low - 2 - Medium': '1',
    '4 - Low - 3 - Low': '2'
}

# Buat DataFrame untuk tabel mapping SLA
df_sla_mapping = pd.DataFrame([
    {"Business criticality-Severity": k, "Waktu SLA (jam)": v}
    for k, v in sla_mapping.items()
])

st.subheader("📋 Tabel Mapping SLA")
st.dataframe(df_sla_mapping)

# Upload file
uploaded_file = st.file_uploader("Upload file Excel", type=["xlsx", "xls"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    # Buat kolom baru "Business criticality-Severity"
    if 'Business criticality' in df.columns and 'Severity' in df.columns:
        df['Business criticality-Severity'] = (
            df['Business criticality'].astype(str) + " - " + df['Severity'].astype(str)
        )

    # Buat kolom baru "Waktu SLA"
    if 'Business criticality-Severity' in df.columns:
        df['Waktu SLA'] = df['Business criticality-Severity'].map(sla_mapping)

    # Buat kolom baru "Target Selesai Baru"
    if 'Tiket Dibuat' in df.columns and 'Waktu SLA' in df.columns:
        df['Target Selesai Baru'] = pd.to_datetime(df['Tiket Dibuat']) + pd.to_timedelta(
            df['Waktu SLA'].astype(float), unit='h'
        )

    # Buat kolom baru "SLA"
    if 'Target Selesai Baru' in df.columns and 'Resolved' in df.columns:
        df['SLA'] = (pd.to_datetime(df['Target Selesai Baru']) <= pd.to_datetime(df['Resolved'])).astype(int)
        # 1 berarti on time, 0 berarti terlambat

    # Buat kolom baru "Time Breach"
    # if SLA = 0, hitung keterlambatan dengan Resolved - Created - Waktu SLA. if SLA = 1, keterlambatan = 0
    if {'SLA','Resolved','Created','Waktu SLA'}.issubset(df.columns):
        df['Time Breach'] = df.apply(
            lambda row: (
                (pd.to_datetime(row['Resolved'])
                - pd.to_datetime(row['Created'])
                - pd.to_timedelta(float(row['Waktu SLA']), unit='h')
                ).total_seconds() / 3600
            ) if row['SLA'] == 0 else 0,
            axis=1
        )

    st.subheader("📊 Data dari Excel:")
    st.dataframe(df)

    # Download CSV hasil
    csv = df.to_csv(index=False).encode('utf-8')
    st.download_button(
        "Download CSV hasil",
        data=csv,
        file_name='data_gabungan.csv',
        mime='text/csv'
    )
else:
    st.info("Silakan upload file Excel di atas")
