import streamlit as st
import importlib

st.set_page_config(page_title="Data Prototype KP", page_icon="📊", layout="wide")

st.title("📊 Data Prototype KP")
st.write("Pilih salah satu halaman di bawah:")

# Sidebar navigation
menu = st.sidebar.selectbox("Pilih halaman:", ["Main Menu", "Portaverse", "Incident"])

if menu == "Main Menu":
    st.header("Selamat datang! 👋")
    st.write("Gunakan menu di sebelah kiri untuk berpindah halaman.")
elif menu == "Portaverse":
    portaverse = importlib.import_module("portaverse")
    portaverse.run()
elif menu == "Incident":
    incident = importlib.import_module("incident")
    incident.run()
