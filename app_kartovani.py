import streamlit as st
from kartovani_page import render_kartovani_page

st.set_page_config(
    page_title="Kartování Content App",
    layout="wide",
    initial_sidebar_state="collapsed",
)

render_kartovani_page()