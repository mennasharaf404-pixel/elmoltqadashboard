import base64
from datetime import date
from pathlib import Path

import streamlit as st
from theme import DEPARTMENTS

# Color Palette 
NAVY = "#18528D"
ACCENT = "#8BA7BD"
BORDER = "#DBE2EA"
TEXT_MUTED = "#4A6580"


@st.cache_data
def logo_b64() -> str:
    return base64.b64encode(Path("assets/logo.jpg").read_bytes()).decode()


st.markdown(
    f"""
    <style>
      /* Hide default Streamlit chrome for a cleaner, app-like look */
      #MainMenu, footer {{ visibility: hidden; }}
      .block-container {{ padding-top: 4.5rem; max-width: 1200px; }}

      /* Sidebar logo: rounded white tile */
      [data-testid="stSidebarHeader"] img {{
        border-radius: 8px; background: #fff; padding: 3px;
      }}

      /* ---------- Header: logo is the banner background ---------- */
      .hero {{
        position: relative; isolation: isolate; overflow: hidden;
        background-color: #fff;
        border: 1px solid {BORDER};
        border-bottom: 5px solid {NAVY};
        border-radius: 14px;
        min-height: 220px;
        padding: 2rem 2.75rem;
        margin-bottom: 2.25rem;
        display: flex; flex-direction: column; justify-content: center; gap: .5rem;
        box-shadow: 0 2px 10px rgba(8, 45, 82, .08);
      }}
      /* Logo layer: scaled past the image's blank margins, then brightened and
         multiplied so its off-white background disappears into the banner */
      .hero::after {{
        content: ""; position: absolute; inset: 0; z-index: 0;
        background: url("data:image/jpeg;base64,{logo_b64()}") no-repeat
                    right 1rem center / auto 165%;
        filter: brightness(1.1) contrast(1.15);
        mix-blend-mode: multiply;
        pointer-events: none;
      }}
      .hero-title, .hero-subtitle {{ position: relative; z-index: 1; }}
      .hero-title {{
        color: {NAVY} !important; font-size: 2.4rem !important;
        font-weight: 700 !important; letter-spacing: -.01em;
        line-height: 1.15; margin: 0 !important; max-width: 55%;
      }}
      .hero-subtitle {{
        color: {TEXT_MUTED} !important; font-size: 1.05rem !important;
        margin: 0 !important; max-width: 50%;
      }}

      /* ---------- Department cards ---------- */
      [data-testid="stPageLink"] a {{
        background: #fff;
        border: 1px solid {BORDER};
        border-top: 3px solid {ACCENT};
        border-radius: 10px;
        min-height: 96px;
        justify-content: center;
        padding: 1rem .5rem;
        box-shadow: 0 1px 3px rgba(8, 45, 82, .08);
        transition: background .2s ease, border-color .2s ease, box-shadow .2s ease;
      }}
      [data-testid="stPageLink"] a p {{
        color: {NAVY} !important; font-size: 1rem !important;
        font-weight: 600 !important; transition: color .2s ease;
        white-space: normal !important; overflow: visible !important;
        text-overflow: clip !important; text-align: center;
      }}
      /* Larger icons */
      [data-testid="stPageLink"] a span[data-testid^="stIcon"] {{
        font-size: 1.6rem !important; width: 1.6rem; height: 1.6rem;
      }}
      [data-testid="stPageLink"] a:hover {{
        background: {NAVY}; border-color: {NAVY};
        box-shadow: 0 6px 16px rgba(8, 45, 82, .18);
      }}
      [data-testid="stPageLink"] a:hover p {{ color: #fff !important; }}
      [data-testid="stPageLink"] a:focus-visible {{
        outline: 3px solid {ACCENT}; outline-offset: 2px;
      }}

      /* Footer line */
      .page-footer {{
        margin-top: 3rem; padding-top: 1rem; border-top: 1px solid {BORDER};
        color: {TEXT_MUTED}; font-size: .85rem;
        display: flex; justify-content: space-between; flex-wrap: wrap; gap: .5rem;
      }}

      /* ---------- Small screens: logo moves below the text ---------- */
      @media (max-width: 760px) {{
        .hero {{ padding: 1.5rem 1.5rem 9rem; justify-content: flex-start; }}
        .hero::after {{ background-position: center bottom -1rem;
                        background-size: auto 85%; }}
        .hero-title {{ font-size: 1.6rem !important; max-width: 100%; }}
        .hero-subtitle {{ max-width: 100%; }}
      }}
      @media (prefers-reduced-motion: reduce) {{
        [data-testid="stPageLink"] a, [data-testid="stPageLink"] a p {{ transition: none; }}
      }}
    </style>

    <div class="hero">
      <p class="hero-title">Departments Dashboard</p>
      <p class="hero-subtitle">Select a department to view its performance and reports.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

CARDS = [
    ("Operations", "views/operations.py"),
    ("Customer Service", "views/customer_service.py"),
    ("Financial", "views/financial.py"),
    ("Marketing", "views/marketing.py"),
    ("HR", "views/hr.py"),
]

for col, (name, path) in zip(st.columns(len(CARDS), gap="small"), CARDS):
    col.page_link(
        path,
        label=name,
        icon=DEPARTMENTS[name]["icon"],
        use_container_width=True,
    )

st.markdown(
    f"""
    <div class="page-footer">
      <span>Elmoltaqa Developments</span>
      <span>Last updated: {date.today():%d %b %Y}</span>
    </div>
    """,
    unsafe_allow_html=True,
)