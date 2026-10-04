"""
Shared visual theme for the ELMOLTQA Developments multi-department dashboard.
Import `inject_css()` and `render_header()` at the top of every page.
"""

import base64
import os
import streamlit as st

# ---- Brand palette (sampled from the company logo) -------------------------
NAVY        = "#173B4C"   # primary brand color (logo ink)
NAVY_DARK   = "#0E2530"   # deep shade for gradients / sidebar
NAVY_LIGHT  = "#2E5A70"   # lighter tint for secondary elements
ACCENT_GOLD = "#C79A3A"   # warm accent for highlights / KPIs
ACCENT_TEAL = "#3FA7A0"   # cool accent for secondary charts
BG          = "#F5F6F7"   # page background (matches logo canvas)
CARD_BG     = "#FFFFFF"
TEXT_MUTED  = "#5B6B73"

DEPARTMENTS = {
    "Operations":        {"icon": ":material/precision_manufacturing:", "color": NAVY},
    "Customer Service":  {"icon": ":material/support_agent:", "color": ACCENT_TEAL},
    "Financial":         {"icon": ":material/payments:", "color": ACCENT_GOLD},
    "Marketing":         {"icon": ":material/campaign:", "color": "#8A5A9E"},
    "HR":                {"icon": ":material/groups:", "color": "#C0563B"},
}

LOGO_PATH = os.path.join(os.path.dirname(__file__), "assets", "logo.jpg")


def _logo_base64() -> str:
    if not os.path.exists(LOGO_PATH):
        return ""
    with open(LOGO_PATH, "rb") as f:
        return base64.b64encode(f.read()).decode()


def inject_css():
    st.markdown(
        f"""
        <style>
        .stApp {{
            background-color: {BG};
        }}
        section[data-testid="stSidebar"] {{
            background: linear-gradient(180deg, {NAVY_DARK} 0%, {NAVY} 100%);
        }}
        section[data-testid="stSidebar"] * {{
            color: #EAF1F4 !important;
        }}
        section[data-testid="stSidebar"] hr {{
            border-color: rgba(255,255,255,0.15);
        }}
        /* KPI cards */
        div[data-testid="stMetric"] {{
            background-color: {CARD_BG};
            border: 1px solid #E4E8EA;
            border-left: 5px solid {NAVY};
            border-radius: 10px;
            padding: 16px 18px 10px 18px;
            box-shadow: 0 1px 3px rgba(20,40,50,0.06);
        }}
        div[data-testid="stMetricLabel"] {{
            color: {TEXT_MUTED};
            font-weight: 600;
        }}
        /* Section headers */
        h1, h2, h3 {{
            color: {NAVY};
        }}
        /* Tabs */
        button[data-baseweb="tab"] {{
            font-weight: 600;
        }}
        /* Top banner */
        .brand-banner {{
            display: flex;
            align-items: center;
            gap: 16px;
            padding: 14px 20px;
            background-color: {CARD_BG};
            border: 1px solid #E4E8EA;
            border-radius: 12px;
            margin-bottom: 22px;
            box-shadow: 0 1px 3px rgba(20,40,50,0.06);
        }}
        .brand-banner img {{
            height: 46px;
        }}
        .brand-title {{
            font-size: 1.35rem;
            font-weight: 700;
            color: {NAVY};
            line-height: 1.15;
        }}
        .brand-subtitle {{
            font-size: 0.9rem;
            color: {TEXT_MUTED};
        }}
        .dept-pill {{
            display: inline-block;
            padding: 3px 12px;
            border-radius: 999px;
            font-size: 0.78rem;
            font-weight: 700;
            color: white;
            margin-left: auto;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_header(department: str | None = None):
    """Top banner with logo + optional department pill."""
    b64 = _logo_base64()
    img_tag = f'<img src="data:image/jpeg;base64,{b64}" />' if b64 else ""
    pill = ""
    if department:
        color = DEPARTMENTS.get(department, {}).get("color", NAVY)
        pill = f'<span class="dept-pill" style="background-color:{color}">{department}</span>'

    st.markdown(
        f"""
        <div class="brand-banner">
            {img_tag}
            <div>
                <div class="brand-title">ELMOLTQA Developments</div>
                <div class="brand-subtitle">Performance &amp; Analytics Dashboard</div>
            </div>
            {pill}
        </div>
        """,
        unsafe_allow_html=True,
    )


def sidebar_brand():
    b64 = _logo_base64()
    with st.sidebar:
        if b64:
            st.markdown(
                f"""
                <div style="text-align:center; padding: 6px 0 2px 0;">
                    <img src="data:image/jpeg;base64,{b64}" style="width:80%; border-radius:6px; background:#F5F6F7; padding:8px;" />
                </div>
                """,
                unsafe_allow_html=True,
            )
        st.markdown(
            "<div style='text-align:center; font-size:0.78rem; opacity:0.8; margin-bottom:10px;'>",
            unsafe_allow_html=True,
        )
        st.markdown("---")
