import streamlit as st
from kartovani_page import render_kartovani_page
from albi_page import render_albi_page

st.set_page_config(
    page_title="GW kartování / ALBI",
    layout="wide",
)

engine = st.radio(
    "Vyber sekci",
    ["GW - kartování bez podkladů", "ALBI"],
    key="app_engine"
)

if engine == "GW - kartování bez podkladů":
    render_kartovani_page()
else:
    render_albi_page()