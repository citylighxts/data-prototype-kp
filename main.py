import streamlit as st
import importlib

# Konfigurasi halaman Streamlit
st.set_page_config(page_title="SLA Incident and Request Visualization", page_icon="📊", layout="wide")

# Judul halaman utama di tengah
st.markdown(
    """
    <div style="text-align:center;">
        <h1>SLA Incident and Request Visualization</h1>
    </div>
    """,
    unsafe_allow_html=True
)

# Pilihan halaman di tengah
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    menu = st.selectbox("Choose one page below:", ["Home", "Reqitem", "Incident"])

# Tampilkan konten tengah berdasarkan pilihan menu
if menu == "Home":
    st.markdown(
        """
        <div style="text-align:center; margin-top: 50px;">
            <h2>Welcome!</h2>
            <p>Please use the menu above to move between pages.</p>
        </div>
        """,
        unsafe_allow_html=True
    )

elif menu == "Reqitem":
    # Import file reqitem dan jalankan fungsi utamanya
    reqitem = importlib.import_module("reqitem")
    reqitem.run()

elif menu == "Incident":
    # Import file incident dan jalankan fungsi utamanya
    incident = importlib.import_module("incident")
    incident.run()
