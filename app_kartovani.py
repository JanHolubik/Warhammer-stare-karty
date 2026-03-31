import streamlit as st
from kartovani_page import render_kartovani_page
from albi_page import render_albi_page

engine = st.radio("Vyber sekci", ["Kartování", "ALBI"])

if engine == "Kartování":
    render_kartovani_page()
else:
    render_albi_page()