import streamlit as st
from kartovani_page import render_kartovani_page
from albi_page import render_albi_page

st.set_page_config(
    page_title="Kartování / ALBI",
    layout="wide",
)

engine = st.radio("Vyber sekci", ["Kartování", "ALBI"], key="app_engine")

if engine == "Kartování":
    render_kartovani_page()
else:
    render_albi_page()