"""ELMOLTQA Developments dashboard - entry point.

Run from the `dashboard/` folder:  streamlit run app.py
"""
from pathlib import Path

import streamlit as st
from theme import inject_css, DEPARTMENTS

# Resolve assets from this file's location so it works locally and on Streamlit Cloud.
BASE_DIR = Path(__file__).resolve().parent
LOGO = str(BASE_DIR / "assets" / "logo.jpg")

st.set_page_config(
    page_title="ELMOLTQA Developments | Dashboard",
    page_icon=LOGO,
    layout="wide",
    initial_sidebar_state="expanded",
)
inject_css()
st.logo(LOGO, size="large")

home = st.Page("views/home.py", title="Home", icon=":material/home:", default=True)
operations = st.Page("views/operations.py", title="Operations", icon=DEPARTMENTS["Operations"]["icon"])
customer_service = st.Page("views/customer_service.py", title="Customer Service", icon=DEPARTMENTS["Customer Service"]["icon"])
marketing = st.Page("views/marketing.py", title="Marketing", icon=DEPARTMENTS["Marketing"]["icon"])
hr = st.Page("views/hr.py", title="HR", icon=DEPARTMENTS["HR"]["icon"])

pg = st.navigation(
    {
        "Overview": [home],
        "Departments": [operations, customer_service, marketing, hr],
    }
)
pg.run()
