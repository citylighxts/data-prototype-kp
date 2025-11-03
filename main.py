import streamlit as st
import importlib

st.set_page_config(page_title="Visualisasi SLA Incident dan Request", page_icon="📊", layout="wide")

# Judul halaman utama di tengah
st.markdown(
    """
    <div style="text-align:center;">
        <h1>Visualisasi SLA Incident dan Request</h1>
    </div>
    """,
    unsafe_allow_html=True
)

# Pilihan halaman di tengah
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    menu = st.selectbox("Pilih salah satu halaman di bawah ini:", ["Home", "Reqitem", "Incident"])

# Tampilkan konten tengah berdasarkan pilihan menu
if menu == "Home":
    st.markdown(
        """
        <div style="text-align:center; margin-top: 50px;">
            <h2>Selamat datang!</h2>
            <p>Gunakan menu di atas untuk berpindah halaman.</p>
        </div>
        """,
        unsafe_allow_html=True
    )

elif menu == "Reqitem":
    reqitem = importlib.import_module("reqitem")
    reqitem.run()

elif menu == "Incident":
    incident = importlib.import_module("incident")
    incident.run()
