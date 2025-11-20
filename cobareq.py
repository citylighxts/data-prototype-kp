# app.py
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import io

def run():
    st.set_page_config(page_title="Incident/Request Visualizer", layout="wide")

    st.title("Incident / Request Visualizer — Deploy untuk IT Dept (Regional 3 filter)")

    st.markdown(
        """
        Upload data incident/request (Excel/CSV).  
        Jika Anda punya sheet/Excel `Mapping SLA`, upload juga agar lookup Target SLA akurat.
        """
    )

    # --- Regional 3 lokasi list (dari user) ---
    REGIONAL3_LIST_RAW = """
    P. Lembar,Regional 3,P. Batulicin,R. Jawa,Terminal Celukan Bawang,Sub Regional BBN,P. Tg. Emas,P. Bumiharjo,Tanjung Perak,R. Bali Nusra,P. Badas,TANJUNGPERAK,TANJUNGEMAS/KEUANGAN,TANJUNGEMAS,P. Tg. Intan, BANJARMASIN/TPK, KOTABARU/MEKARPUTIH,P. Waingapu, R. Kalimantan,Terminal Nilam, Terminal Kumai,P. Kalimas, P. Tg. Wangi,P. Gresik, P. Kotabaru,BANJARMASIN/KOMERSIAL,TANJUNGWANGI/TEKNIK,Sub Regional Kalimantan,GRESIK/TERMINAL,Terminal Kota Baru,P. Sampit,BANJARMASIN/TMP,P. Bagendang,BANJARMASIN/PDS,TENAU/KALABAHI,P. Bima,P. Tenau Kupang,Terminal Lembar,P. Tegal,Terminal Trisakti,BENOA/OPKOM,P. Benoa,BANJARMASIN/TEKNIK,BANJARMASIN/PBJ,TANJUNGINTAN,KOTABARU,TENAU,Sub Regional Jawa Timur,KUMAI/OPKOM,Terminal Batulicin,Terminal Gresik, KUMAI/KEUPER,LEMBAR/KEUPER,P. Kalabahi,BIMA/BADAS,Terminal Jamrud,TENAU/WAINGAPU,Terminal Benoa,P. Tg. Tembaga,BIMA/PDS,BENOA/SUK,P. Clk. Bawang,KUMAI/BUMIHARJO,P. Pulang Pisau,Terminal Labuan Bajo,P. Maumere,BENOA/KEUANGAN,BENOA/PKWT,Terminal Kalimas,BANJARMASIN/KEUANGAN,BENOA/PEMAGANG,GRESIK/KEUANGAN,Terminal Petikemas Banjarmasin,CELUKANBAWANG,P. Ende-Ippi,SAMPIT/BAGENDANG,Terminal Bima,KOTABARU/KEPANDUAN,Terminal Sampit,Terminal Kupang, BENOA/TEKNIK, Terminal Maumere, PROBOLINGGO/PLS, SAMPIT/PKWT, P. Labuan Bajo, P. Kalianget, Banjarmasin, Terminal Waingapu, MAUMERE/ENDE
    """
    REGIONAL3_SET = {s.strip().upper() for s in REGIONAL3_LIST_RAW.split(",") if s.strip()}

    # helper: apakah lokasi termasuk regional 3? (case-insensitive, substring match)
    def is_regional3(location_str):
        if not isinstance(location_str, str):
            return False
        s = location_str.upper()
        # check exact tokens or substring
        for token in REGIONAL3_SET:
            if token == "":
                continue
            if token in s:
                return True
        return False

    # --- Upload inputs ---
    col1, col2 = st.columns([1,1])
    with col1:
        uploaded = st.file_uploader("Upload dataset (Excel/CSV) — kolom sesuai deskripsi", type=["xlsx","xls","csv"], accept_multiple_files=False)

    with col2:
        mapping_file = st.file_uploader("Opsional: Upload file Mapping SLA (Excel) jika ada", type=["xlsx","xls","csv"], accept_multiple_files=False)

    # --- Parse uploaded data ---
    @st.cache_data
    def read_input_file(f):
        if f is None:
            return None
        name = getattr(f, "name", "")
        if name.lower().endswith(".csv"):
            df = pd.read_csv(f)
        else:
            # try excel first sheet
            try:
                df = pd.read_excel(f)
            except Exception as e:
                # try to read all sheets and concat
                x = pd.read_excel(f, sheet_name=None)
                # pick first
                df = list(x.values())[0]
        return df

    @st.cache_data
    def read_mapping_file(f):
        if f is None:
            return None
        name = getattr(f, "name", "")
        if name.lower().endswith(".csv"):
            df = pd.read_csv(f)
        else:
            try:
                # try to find sheet 'Mapping SLA' or take first sheet
                x = pd.read_excel(f, sheet_name=None)
                if 'Mapping SLA' in x:
                    df = x['Mapping SLA']
                else:
                    df = list(x.values())[0]
            except Exception as e:
                return None
        return df

    df = read_input_file(uploaded)
    mapping_df = read_mapping_file(mapping_file)

    if df is None:
        st.info("Unggah dataset Anda terlebih dahulu (Excel/CSV). Saya sudah menyiapkan template processing.")
        st.stop()

    st.success(f"Data terupload — {len(df)} baris")

    # Normalize column names: strip spaces
    df.columns = [c.strip() for c in df.columns]

    # Ensure key columns exist (best-effort)
    required_cols = ["No. Tiket","Tiket Dibuat","Tiket Ditutup","Businesscriticality","Severity","Judul Permasalahan","Item","Lokasi Pelapor"]
    # Some users may have slightly different naming; try to detect alternative names
    col_lookup = {c.lower(): c for c in df.columns}
    def get_col(name_variants):
        for v in name_variants:
            if v.lower() in col_lookup:
                return col_lookup[v.lower()]
        return None

    col_no = get_col(["No. Tiket","No Tiket","Tiket"])
    col_created = get_col(["Tiket Dibuat","Tiket dibuat","Tiket Dibuat " ,"Tiket Dibuat"])
    col_closed = get_col(["Tiket Ditutup","Tiket ditutup","Tiket Ditutup "])
    col_business = get_col(["Businesscriticality","Businesscriticality ","Businesscriticality " ,"Businesscriticality"])
    col_severity = get_col(["Severity","severity"])
    col_judul = get_col(["Judul Permasalahan","Judul permasalahan","Judul Permasalahan "])
    col_item = get_col(["Item","Item "])
    col_lokasi = get_col(["Lokasi Pelapor","Lokasi Pelapor " ,"Lokasi Pelapor"])

    # If some not found, show a warning but proceed
    missing = [name for name,var in [
        ("No. Tiket", col_no),
        ("Tiket Dibuat", col_created),
        ("Tiket Ditutup", col_closed),
        ("Businesscriticality", col_business),
        ("Severity", col_severity),
        ("Judul Permasalahan", col_judul),
        ("Item", col_item),
        ("Lokasi Pelapor", col_lokasi),
    ] if var is None]

    if missing:
        st.warning(f"Terdapat kolom penting yang tidak terdeteksi otomatis: {missing}. Aplikasi akan tetap mencoba praktik terbaik, namun hasil lookup SLA mungkin tidak sempurna.")

    # Work on a copy
    data = df.copy()

    # Parse datetime fields
    def try_parse_datetime(s):
        if pd.isna(s) or s=="":
            return pd.NaT
        # try common formats
        for fmt in ("%d-%m-%Y %H:%M:%S","%d-%m-%Y %H:%M","%Y-%m-%d %H:%M:%S","%Y-%m-%d","%d/%m/%Y %H:%M:%S","%d/%m/%Y"):
            try:
                return datetime.strptime(str(s).strip(), fmt)
            except:
                continue
        # try pandas
        try:
            return pd.to_datetime(s, dayfirst=True, errors='coerce')
        except:
            return pd.NaT

    # Apply parsing if columns exist
    if col_created:
        data["__Tiket_Dibuat_parsed"] = data[col_created].apply(try_parse_datetime)
    else:
        data["__Tiket_Dibuat_parsed"] = pd.NaT

    if col_closed:
        data["__Tiket_Ditutup_parsed"] = data[col_closed].apply(try_parse_datetime)
    else:
        data["__Tiket_Ditutup_parsed"] = pd.NaT

    # Fill businesscriticality/severity choices
    if col_business:
        data["Businesscriticality"] = data[col_business].astype(str).fillna("")
    else:
        data["Businesscriticality"] = ""

    if col_severity:
        data["Severity"] = data[col_severity].astype(str).fillna("")
    else:
        data["Severity"] = ""

    # Businesscriticality-Severity column
    data["Businesscriticality-Severity"] = data["Businesscriticality"].str.strip() + data["Severity"].str.strip()

    # --- Mapping SLA logic ---
    # If user provided mapping_df we will try to use it:
    # We'll search for rows where Judul Permasalahan or Item matches and Business Critical - Severity matches.
    def parse_sla_value(val):
        """Return SLA in hours (float) from common string formats like '00:30', '02:00', '5', '30' (minutes?), or '30:00'."""
        if pd.isna(val):
            return np.nan
        s = str(val).strip()
        if s == "":
            return np.nan
        # If contains colon -> hh:mm
        if ":" in s:
            try:
                parts = s.split(":")
                if len(parts) == 2:
                    hh = int(parts[0])
                    mm = int(parts[1])
                    return hh + mm/60.0
                elif len(parts) == 3:
                    hh = int(parts[0]); mm = int(parts[1]); ss=int(parts[2])
                    return hh + mm/60.0 + ss/3600.0
            except:
                pass
        # If string looks like '30' but earlier mapping uses '00:30' to mean 30 minutes -> assume minutes if value <= 60 AND contains no decimal but mapping often uses 00:30 meaning 30 minutes;
        # treat numbers <= 24 as hours? ambiguous. We'll assume:
        # - if numeric and <= 24 -> treat as hours
        # - if numeric and > 24 -> treat as minutes (rare)
        try:
            v = float(s.replace(",","."))
            if v <= 24:
                return v
            else:
                # treat as minutes
                return v/60.0
        except:
            pass
        return np.nan

    # Build fallback mapping dict from the snippet for common BusinessCritical-Severity to minutes/hours
    FALLBACK_BC_SEV_TO_SLA = {
        # from snippet: many combos map to '00:30' => 0.5 hours
        # sample:
        "1-Critical1-High": "00:30",
        "1-Critical2-Medium": "00:30",
        "1-Critical3-Low": "00:30",
        "2-High1-High": "00:30",
        "2-High2-Medium": "00:30",
        "2-High3-Low": "00:30",
        "3-Medium1-High": "00:30",
        "3-Medium2-Medium": "00:30",
        "3-Medium3-Low": "00:30",
        "4-Low1-High": "00:30",
        "4-Low2-Medium": "00:30",
        "4-Low3-Low": "00:30",
        # Other blocks: Request Akses (Pemberian Hak Akses) had 02:00,06:00,08:00 depending on severity combos
        # We'll include some common ones:
        "1-Critical1-High_reqaccess": "02:00",
        # For Penyediaan Data block: examples: 04:00,06:00,08:00,12:00,16:00
        # We'll add a simple default: if item contains "Penyediaan Data" use 4h
    }

    # Convert fallback strings to float hours
    FALLBACK_BC_SEV_HOURS = {k: parse_sla_value(v) for k,v in FALLBACK_BC_SEV_TO_SLA.items()}

    # Function to find SLA hours for a row using mapping_df if available, else fallback heuristics
    def lookup_target_sla_hours(row, mapping_df=None):
        # try user mapping_df first
        if mapping_df is not None:
            # normalize mapping_df columns lower
            m = mapping_df.copy()
            cols = {c.lower(): c for c in m.columns}
            # try to find columns for Judul Permasalahan, Item, Business Critical - Severity, SLA
            jud_col = cols.get("judul permasalahan") or cols.get("judul_permasalahan") or cols.get("judul permasalahan ".strip()) or None
            item_col = cols.get("item") or None
            bcsev_col = None
            for candidate in ["business critical - severity","businesscriticality-severity","business critical - severity ","business criticality - severity"]:
                if candidate in cols:
                    bcsev_col = cols[candidate]
                    break
            sla_col = None
            for candidate in ["sla","s l a","00:30","s t r"] :
                if "sla" in cols:
                    sla_col = cols["sla"]
                    break
            # fallback try to guess column that contains time strings (COLUMNS with ":" or "00:30")
            if sla_col is None:
                for c in m.columns:
                    sample = str(m[c].dropna().astype(str).head(5).values)
                    if ":" in sample or any(d in sample for d in ["00:30","02:00","04:00","08:00"]):
                        sla_col = c
                        break
            # Try to find matching rows
            candidates = m
            # Filter by judul if present in mapping
            if jud_col is not None and pd.notna(row.get(col_judul, None)):
                mask = candidates[jud_col].astype(str).str.strip().str.upper() == str(row.get(col_judul,"")).strip().upper()
                matches = candidates[mask]
                if len(matches):
                    # Further filter by bcsev if available
                    if bcsev_col is not None:
                        bcval = str(row.get("Businesscriticality-Severity","")).strip()
                        matches2 = matches[matches[bcsev_col].astype(str).str.strip()==bcval]
                        if len(matches2):
                            val = matches2.iloc[0].get(sla_col, np.nan)
                            return parse_sla_value(val)
                    # else return sla from first match
                    val = matches.iloc[0].get(sla_col, np.nan)
                    return parse_sla_value(val)
            # try match by Item
            if item_col is not None and pd.notna(row.get(col_item, None)):
                mask = candidates[item_col].astype(str).str.strip().str.upper() == str(row.get(col_item,"")).strip().upper()
                matches = candidates[mask]
                if len(matches):
                    # try bcsev
                    if bcsev_col is not None:
                        bcval = str(row.get("Businesscriticality-Severity","")).strip()
                        matches2 = matches[matches[bcsev_col].astype(str).str.strip()==bcval]
                        if len(matches2):
                            val = matches2.iloc[0].get(sla_col, np.nan)
                            return parse_sla_value(val)
                    val = matches.iloc[0].get(sla_col, np.nan)
                    return parse_sla_value(val)
            # if nothing found, continue to fallback
        # fallback heuristics:
        bcsev = str(row.get("Businesscriticality-Severity","")).strip()
        item = str(row.get(col_item,"")).lower() if col_item else ""
        # direct map
        if bcsev in FALLBACK_BC_SEV_HOURS:
            return FALLBACK_BC_SEV_HOURS[bcsev]
        # some heuristics:
        if "penyediaan data" in item or "penyediaan data" in str(row.get("Permintaan","")).lower():
            return 4.0  # 4 jam default
        if "video" in item or "video conference" in item.lower():
            return 0.5
        if "reset password" in item or "akun" in item.lower() or "akses" in item.lower():
            return 0.5
        if "pendampingan" in item or "implementasi" in item:
            return 30.0/60.0  # fallback 0.5
        # default fallback
        return np.nan

    # calculate Target SLA hours for each row (vectorized apply)
    st.info("Menghitung Target SLA untuk setiap baris — gunakan mapping file jika ada untuk hasil terbaik.")
    data["Target SLA (hours)"] = data.apply(lambda r: lookup_target_sla_hours(r, mapping_df), axis=1)

    # Target Selesai = Tiket Dibuat + Target SLA (hours)
    def compute_target_selesai(created_dt, sla_hours):
        if pd.isna(created_dt) or pd.isna(sla_hours):
            return pd.NaT
        try:
            return created_dt + timedelta(hours=float(sla_hours))
        except:
            return pd.NaT

    data["Target Selesai (computed)"] = data.apply(lambda r: compute_target_selesai(r["__Tiket_Dibuat_parsed"], r["Target SLA (hours)"]), axis=1)

    # SLA column: if Tiket Ditutup exists -> 1 if closed <= target selesai else 0; else 'WP'
    def compute_sla_status(closed_dt, target_dt):
        if pd.isna(closed_dt):
            return "WP"
        if pd.isna(target_dt):
            return np.nan
        try:
            return 1 if closed_dt <= target_dt else 0
        except:
            return np.nan

    data["SLA"] = data.apply(lambda r: compute_sla_status(r["__Tiket_Ditutup_parsed"], r["Target Selesai (computed)"]), axis=1)

    # Add Location regional3 flag
    if col_lokasi:
        data["Lokasi Pelapor (raw)"] = data[col_lokasi].astype(str)
    else:
        data["Lokasi Pelapor (raw)"] = ""

    data["Is_Regional3"] = data["Lokasi Pelapor (raw)"].apply(is_regional3)

    # Build final output columns order requested by user (attempt)
    output_cols = [
        col_no or "No. Tiket","Tiket Dibuat","Disetujui","Status","Item","Permintaan","Requested for","Target Selesai","Target Selesai(due_date)","Tahapan",
        "Dibuka Oleh","Jumlah","Name","PIC","Comments and Work notes","Deskripsi Permasalahan","Judul Permasalahan","Komentar Tambahan","Root Cause and Solution",
        "Service offering","Lokasi Pelapor","Deskripsi Permasalahan","Tiket Ditutup","Businesscriticality","Contact type","Severity","Data Reg3",
        "Businesscriticality-Severity","Target SLA (hours)","Target Selesai (computed)","SLA"
    ]
    # Many of these columns may not exist in df; create final_df with columns that exist
    final_df = pd.DataFrame()
    for c in output_cols:
        if c in data.columns:
            final_df[c] = data[c]
        else:
            # special mapping synonyms:
            if c == "Businesscriticality-Severity":
                final_df[c] = data["Businesscriticality-Severity"]
            elif c == "Target SLA (hours)":
                final_df[c] = data["Target SLA (hours)"]
            elif c == "Target Selesai (computed)":
                final_df[c] = data["Target Selesai (computed)"]
            elif c == "SLA":
                final_df[c] = data["SLA"]
            elif c == "Lokasi Pelapor":
                final_df[c] = data.get("Lokasi Pelapor (raw)", "")
            else:
                # create empty column if not exist
                final_df[c] = data.get(c, np.nan)

    # Convert datetime columns to strings for display
    if "__Tiket_Dibuat_parsed" in data.columns:
        final_df["Tiket Dibuat_parsed"] = data["__Tiket_Dibuat_parsed"]
    if "Target Selesai (computed)" in final_df.columns:
        final_df["Target Selesai (computed)"] = pd.to_datetime(final_df["Target Selesai (computed)"])
    if "Tiket Ditutup" in final_df.columns:
        try:
            final_df["Tiket Ditutup_parsed"] = pd.to_datetime(data["__Tiket_Ditutup_parsed"])
        except:
            pass

    # --- Sidebar: filters ---
    st.sidebar.header("Filter")
    loc_option = st.sidebar.selectbox("Lokasi Pelapor", options=["All","Regional 3"], index=1)
    only_reg3 = (loc_option == "Regional 3")

    # Apply filters
    filtered = final_df.copy()

    if only_reg3:
        # use Is_Regional3 from data
        filtered = filtered[data["Is_Regional3"]].copy()

    st.subheader("Preview Data (setelah processing & filter)")
    st.write(f"Baris: {len(filtered)}")

    # Show table (limit to first 200 rows to avoid heavy UI)
    st.dataframe(filtered.head(200), use_container_width=True)

    # Summary metrics
    st.markdown("### Ringkasan")
    colA, colB, colC = st.columns(3)
    with colA:
        st.metric("Total baris (setelah filter)", len(filtered))
    with colB:
        # count closed
        if "Tiket Ditutup" in data.columns:
            closed_count = data["__Tiket_Ditutup_parsed"].notna().sum()
        else:
            closed_count = final_df["Tiket Ditutup_parsed"].notna().sum() if "Tiket Ditutup_parsed" in final_df.columns else 0
        st.metric("Total Closed (ada tanggal Tutup)", int(closed_count))
    with colC:
        # SLA compliance rate (only closed rows)
        sla_vals = filtered["SLA"].dropna()
        if len(sla_vals[sla_vals.isin([0,1])])>0:
            num_ok = int((sla_vals==1).sum())
            denom = int(sla_vals.isin([0,1]).sum())
            st.metric("SLA Met (%)", f"{num_ok/denom*100:.1f}%", delta=f"{num_ok}/{denom}")
        else:
            st.metric("SLA Met (%)","-")

    # Simple plots
    st.markdown("### Visualisasi sederhana")
    plot_col1, plot_col2 = st.columns(2)
    with plot_col1:
        st.markdown("Count per Status")
        if "Status" in filtered.columns:
            st.bar_chart(filtered["Status"].astype(str).value_counts())
        else:
            st.info("Kolom Status tidak ditemukan untuk plotting.")

    with plot_col2:
        st.markdown("Average Target SLA (hours) by Item (top 10)")
        if "Item" in data.columns:
            grp = filtered.groupby("Item")["Target SLA (hours)"].mean().dropna().sort_values(ascending=False).head(10)
            st.bar_chart(grp)
        else:
            st.info("Kolom Item tidak ditemukan untuk plotting.")

    # Download processed CSV
    buf = io.StringIO()
    filtered.to_csv(buf, index=False)
    buf.seek(0)
    st.download_button("Download hasil (CSV)", buf.getvalue(), file_name="processed_incident_request.csv", mime="text/csv")

    st.markdown("---")
    st.markdown("**Note / Petunjuk**:")
    st.markdown("""
    - Untuk hasil *Target SLA* paling akurat, upload file `Mapping SLA` yang struktur kolomnya mirip dengan sheet `'Mapping SLA'` yang Anda gunakan di Excel.  
    - App ini menggunakan heuristik bila mapping file tidak disediakan (bisa disesuaikan lebih jauh).  
    - Jika ada kolom yang tidak terdeteksi otomatis, ganti nama kolom di file input agar cocok (mis. `Businesscriticality`, `Severity`, `Tiket Dibuat`, `Tiket Ditutup`, `Judul Permasalahan`, `Item`, `Lokasi Pelapor`).  
    - Format tanggal yang dikenali: `DD-MM-YYYY HH:MM:SS`, `DD-MM-YYYY HH:MM`, `YYYY-MM-DD HH:MM:SS`, dan beberapa format umum lainnya.
    """)

    st.success("Selesai. Kalau mau saya tambah: visualisasi bulanan, pivot table interaktif, atau export Excel lengkap dengan sheet summary — beri tahu saya fitur mana yang ingin ditambahkan.")
