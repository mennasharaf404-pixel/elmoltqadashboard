import streamlit as st
from theme import inject_css, DEPARTMENTS

st.set_page_config(
    page_title="ELMOLTQA Developments | Dashboard",
    page_icon="assets/logo.jpg",
    layout="wide",
    initial_sidebar_state="expanded",
)
inject_css()
st.logo("assets/logo.jpg", size="large")

home = st.Page("views/home.py", title="Home", icon=":material/home:", default=True)
operations = st.Page("views/operations.py", title="Operations", icon=DEPARTMENTS["Operations"]["icon"])
customer_service = st.Page("views/customer_service.py", title="Customer Service", icon=DEPARTMENTS["Customer Service"]["icon"])
financial = st.Page("views/financial.py", title="Financial", icon=DEPARTMENTS["Financial"]["icon"])
marketing = st.Page("views/marketing.py", title="Marketing", icon=DEPARTMENTS["Marketing"]["icon"])
hr = st.Page("views/hr.py", title="HR", icon=DEPARTMENTS["HR"]["icon"])

pg = st.navigation(
    {
        "Overview": [home],
        "Departments": [operations, customer_service, financial, marketing, hr],
    }
)
pg.run()