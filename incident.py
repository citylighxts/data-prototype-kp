import streamlit as st
import pandas as pd

def run():
    st.markdown(
            """
            <h1 style="display: flex; align-items: center; gap: 10px;">
                <img src="https://cdnjs.cloudflare.com/ajax/libs/twemoji/14.0.2/svg/1f4ca.svg" 
                    width="40" height="40">
                Incident Data
            </h1>
            """,
            unsafe_allow_html=True
        )

    st.write("Upload an Excel file. The system will create columns: `Businesscriticality-Severity`, `Target SLA (jam)`, `Target Selesai`, `SLA`.")

    # Upload file
    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])
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

    possible_name_cols = ['Name', 'Nama', 'User Name', 'Assigned To', 'Assignee']
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
            key="incident_regional_filter"
        )

        if regional_option == "Regional 3":
            df = df[df[name_col].astype(str).str.contains("Regional 3", case=False, na=False)]

        # --- Filter Bulan & Tahun ---
        possible_date_cols = ['Created Date', 'Created', 'Tiket Dibuat', 'Opened', 'CreatedAt']
        date_created_col = next((c for c in possible_date_cols if c in df.columns), None)

        if date_created_col:
            df[date_created_col] = pd.to_datetime(df[date_created_col], errors='coerce')
            df['Bulan-Tahun'] = df[date_created_col].dt.strftime('%Y-%m')
            available_months = sorted(df['Bulan-Tahun'].dropna().unique())

            selected_month = st.selectbox("Choose Month & Year:", options=["All"] + available_months, key="incident_month_filter")
            if selected_month != "All":
                df = df[df['Bulan-Tahun'] == selected_month]

        st.markdown(f"**Data after filter:** {len(df)} rows")
        st.dataframe(df.head(7))


    if uploaded_file is not None:
        progress = st.progress(0)
        status_text = st.empty()
        total_steps = 6
        current_step = 0

        status_text.text("Reading Excel file...")
        df = pd.read_excel(uploaded_file)
        current_step += 1
        progress.progress(int(current_step/total_steps*100))

        # Buat kolom baru "Sum of Time Breach (jam)"
        status_text.text("Creating column Sum of Time Breach (hours)...")
        if 'Sum of Time Breach' in df.columns:
            df['Sum of Time Breach (jam)'] = (
                (df['Sum of Time Breach'] * 24).astype(int).astype(str) + " jam " +
                ((df['Sum of Time Breach'] * 24 * 60) % 60).astype(int).astype(str) + " menit"
            )
        current_step += 1
        progress.progress(int(current_step/total_steps*100))

        # Buat kolom baru "Average of Time Breach (hari)"
        # status_text.text("Membuat kolom Average of Time Breach (hari)...")
        # if 'Average of Time Breach' in df.columns:
        #     df['Average of Time Breach (hari)'] = (
        #         (df['Average of Time Breach'].astype(int).astype(str) + " hari ") +
        #         (df['Average of Time Breach'] * 24).astype(int).astype(str) + " jam " +
        #         ((df['Average of Time Breach'] * 24 * 60) % 60).astype(int).astype(str) + " menit"
        #     )
        # current_step += 1
        # progress.progress(int(current_step/total_steps*100))

        # Buat kolom baru "Max of Time Breach (hari)"
        status_text.text("Creating column Max of Time Breach (days)...")
        if 'Max of Time Breach' in df.columns:
            df['Max of Time Breach (hari)'] = (
                (df['Max of Time Breach'].astype(int).astype(str) + " hari ") +
                (df['Max of Time Breach'] * 24).astype(int).astype(str) + " jam " +
                ((df['Max of Time Breach'] * 24 * 60) % 60).astype(int).astype(str) + " menit"
            )
        current_step += 1
        progress.progress(int(current_step/total_steps*100))

        # Buat kolom baru "Aplikasi"
        status_text.text("Creating column Application...")
        if 'Row Labels' in df.columns:
            df['Aplikasi'] = df['Row Labels']
        current_step += 1
        progress.progress(int(current_step/total_steps*100))

        # Buat kolom baru "SLA Total Tiket setiap Service Offering"
        status_text.text("Creating column SLA Total Tickets per Service Offering...")
        # =IF(((744-(G4*24))/744)<0;0;((744-(G4*24))/744))
        if (((744-(df['Sum of Time Breach']*24))/744).astype(float).all() < 0):
            sla_raw_total = 0
        else:
            sla_raw_total = (
                (((744-(df['Sum of Time Breach']*24))/744).clip(lower=0)).round(4)
            )
        sla_percent_total = (sla_raw_total * 100).round(2).astype(int).astype(str) + "%"
        df['SLA Total Tickets per Service Offering'] = sla_percent_total
        current_step += 1
        progress.progress(int(current_step/total_steps*100))

        # Buat kolom baru "SLA Tiket Max breach setiap Service Offering"
        status_text.text("Creating column SLA Max Breach Tickets per Service Offering...")
        # =(744-(J4*24))/744
        sla_raw_max = (
            (((744-(df['Max of Time Breach']*24))/744).clip(lower=0)).round(4)
        )
        sla_percent_max = (sla_raw_max * 100).round(2).astype(int).astype(str) + "%"
        df['SLA Tiket Max breach setiap Service Offering'] = sla_percent_max
        current_step += 1
        progress.progress(int(current_step/total_steps*100))

        status_text.text("Done!")

        # Tampilkan top 3 dengan time breach terendah
        if 'Sum of Time Breach' in df.columns and 'Aplikasi' in df.columns and 'SLA Total Tickets per Service Offering' in df.columns:
            # Ambil nilai unik terkecil (top 3) dari time breach
            unique_breach = df['Sum of Time Breach'].dropna().unique()
            top_breach = sorted(unique_breach)[:3]
            # Filter baris yang termasuk dalam top 3 breach
            top3 = df[df['Sum of Time Breach'].isin(top_breach)][['Aplikasi', 'Sum of Time Breach', 'SLA Total Tickets per Service Offering']].copy()
            # Buat nomor berdasarkan urutan nilai time breach
            breach_to_no = {v: i+1 for i, v in enumerate(top_breach)}
            top3['No'] = top3['Sum of Time Breach'].map(breach_to_no)
            # Urutkan tabel berdasarkan nomor dan time breach
            top3 = top3.sort_values(['No', 'Sum of Time Breach']).reset_index(drop=True)
            # Siapkan data untuk HTML table dengan rowspan
            html = """
            <table border='1' style='border-collapse:collapse;width:100%;text-align:center;'>
                <thead>
                    <tr>
                        <th>No</th>
                        <th>Service Offering</th>
                        <th>Time Breach</th>
                        <th>SLA</th>
                    </tr>
                </thead>
                <tbody>
            """
            # Hitung rowspan untuk setiap nomor
            for no in sorted(top3['No'].unique()):
                group = top3[top3['No'] == no]
                rowspan = len(group)
                for i, (_, row) in enumerate(group.iterrows()):
                    html += "<tr>"
                    if i == 0:
                        html += f"<td rowspan='{rowspan}'>{no}</td>"
                    # ...existing columns...
                    html += f"<td>{row['Aplikasi']}</td>"
                    html += f"<td>{row['Sum of Time Breach']}</td>"
                    html += f"<td>{row['SLA Total Tickets per Service Offering']}</td>"
                    html += "</tr>"
            html += "</tbody></table>"
            st.subheader("Top 3 Service Offerings with the Lowest Time Breach")
            st.markdown(html, unsafe_allow_html=True)

        # Tampilkan bottom 3 dengan time breach tertinggi
        if 'Sum of Time Breach' in df.columns and 'Aplikasi' in df.columns and 'SLA Total Tickets per Service Offering' in df.columns:
            # Ambil nilai unik terbesar (bottom 3) dari time breach
            unique_breach = df['Sum of Time Breach'].dropna().unique()
            bottom_breach = sorted(unique_breach, reverse=True)[:3]
            # Filter baris yang termasuk dalam bottom 3 breach
            bottom3 = df[df['Sum of Time Breach'].isin(bottom_breach)][['Aplikasi', 'Sum of Time Breach', 'SLA Total Tickets per Service Offering']].copy()
            # Buat nomor berdasarkan urutan nilai time breach
            breach_to_no = {v: i+1 for i, v in enumerate(sorted(bottom_breach, reverse=True))}
            bottom3['No'] = bottom3['Sum of Time Breach'].map(breach_to_no)
            # Urutkan tabel berdasarkan nomor dan time breach
            bottom3 = bottom3.sort_values(['No', 'Sum of Time Breach'], ascending=[True, False]).reset_index(drop=True)
            # Siapkan data untuk HTML table dengan rowspan
            html = """
            <table border='1' style='border-collapse:collapse;width:100%;text-align:center;'>
                <thead>
                    <tr>
                        <th>No</th>
                        <th>Service Offering</th>
                        <th>Time Breach</th>
                        <th>SLA</th>
                    </tr>
                </thead>
                <tbody>
            """
            # Hitung rowspan untuk setiap nomor
            for no in sorted(bottom3['No'].unique()):
                group = bottom3[bottom3['No'] == no]
                rowspan = len(group)
                for i, (_, row) in enumerate(group.iterrows()):
                    html += "<tr>"
                    if i == 0:
                        html += f"<td rowspan='{rowspan}'>{no}</td>"
                    # ...existing columns...
                    html += f"<td>{row['Aplikasi']}</td>"
                    html += f"<td>{row['Sum of Time Breach']}</td>"
                    html += f"<td>{row['SLA Total Tickets per Service Offering']}</td>"
                    html += "</tr>"
            html += "</tbody></table>"
            st.subheader("Bottom 3 Service Offerings with the Highest Time Breach")
            st.markdown(html, unsafe_allow_html=True)

        # bar chart judul "SLA Total Tiket setiap Service Offering", x = Aplikasi, y = SLA Total Tiket setiap Service Offering
        if 'Aplikasi' in df.columns and 'SLA Total Tickets per Service Offering' in df.columns:
            sla_chart = df[['Aplikasi', 'SLA Total Tickets per Service Offering']].dropna().copy()
            # Ubah kolom SLA ke format numerik
            sla_chart['SLA Numeric'] = sla_chart['SLA Total Tickets per Service Offering'].str.rstrip('%').astype(float)
            sla_chart = sla_chart.sort_values('SLA Numeric', ascending=False)
            st.subheader("SLA Total Tickets per Service Offering")
            st.bar_chart(sla_chart.set_index('Aplikasi')['SLA Numeric'])

        # bar chart judul "SLA Tiket Max breach setiap Service Offering", x = Aplikasi, y = SLA Tiket Max breach setiap Service Offering
        if 'Aplikasi' in df.columns and 'SLA Tiket Max breach setiap Service Offering' in df.columns:
            sla_max_chart = df[['Aplikasi', 'SLA Tiket Max breach setiap Service Offering']].dropna().copy()
            # Ubah kolom SLA ke format numerik
            sla_max_chart['SLA Max Numeric'] = sla_max_chart['SLA Tiket Max breach setiap Service Offering'].str.rstrip('%').astype(float)
            sla_max_chart = sla_max_chart.sort_values('SLA Max Numeric', ascending=False)
            st.subheader("SLA Tiket Max breach setiap Service Offering")
            st.bar_chart(sla_max_chart.set_index('Aplikasi')['SLA Max Numeric'])

        st.subheader("Data from Excel:")
        st.dataframe(df)

        # Download CSV hasil
        csv = df.to_csv(index=False).encode('utf-8')
        st.download_button(
            "Download Result CSV",
            data=csv,
            file_name='Hasil_Analisis.csv',
            mime='text/csv'
        )
    else:
        st.info("Please upload an Excel file above")
