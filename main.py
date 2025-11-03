import streamlit as st
import importlib

st.set_page_config(page_title="Visualisasi SLA Incident dan Request", page_icon="📊", layout="wide")

st.title("Visualisasi SLA Incident dan Request")
st.write("Pilih salah satu halaman di bawah:")

# Sidebar navigation
menu = st.sidebar.selectbox("Pilih halaman:", ["Main Menu", "Reqitem", "Incident"])

if menu == "Main Menu":
    st.header("Selamat datang!")
    st.write("Gunakan menu di sebelah kiri untuk berpindah halaman.")
elif menu == "Reqitem":
    reqitem = importlib.import_module("reqitem")
    reqitem.run()
elif menu == "Incident":
    incident = importlib.import_module("incident")
    incident.run()
