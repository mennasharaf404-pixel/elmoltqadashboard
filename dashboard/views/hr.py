import pandas as pd
import streamlit as st

# COLOR PALETTE
NAVY = "#082D52"
ACCENT = "#8BA7BD"
BORDER = "#DBE2EA"
TEXT_MUTED = "#4A6580"
GOOD, BAD = "#1E8E5A", "#C0392B"

#  DATA  --  SAMPLE VALUES.  
@st.cache_data
def load_data() -> dict:
    months = ["Nov", "Dec", "Jan", "Feb", "Mar", "Apr",
              "May", "Jun", "Jul", "Aug", "Sep", "Oct"]
    return {
        "headcount_by_dept": pd.DataFrame({
            "Department": ["Sales", "Engineering & Site", "Operations",
                           "Customer Service", "Finance", "Marketing", "HR"],
            "Employees": [64, 118, 36, 28, 17, 14, 8],
        }),
        "monthly": pd.DataFrame({
            "Month": months,
            "Hires": [6, 3, 8, 5, 7, 9, 4, 6, 10, 5, 7, 8],
            "Leavers": [4, 5, 3, 6, 4, 5, 7, 4, 5, 6, 3, 4],
            "Attendance %": [94.1, 93.5, 95.0, 94.6, 95.4, 94.8,
                             93.9, 95.2, 95.6, 94.4, 95.1, 95.7],
        }),
        "pipeline": pd.DataFrame({
            "Stage": ["1 Applied", "2 Screened", "3 Interviewed",
                      "4 Offered", "5 Hired"],
            "Candidates": [240, 120, 55, 18, 14],
        }),
        "vacancies": pd.DataFrame({
            "Position": ["Sales Executive", "Site Engineer", "Quantity Surveyor",
                         "Customer Service Agent", "Accountant"],
            "Department": ["Sales", "Engineering & Site", "Engineering & Site",
                           "Customer Service", "Finance"],
            "Openings": [4, 3, 2, 2, 1],
            "Days open": [21, 35, 48, 12, 9],
            "Stage": ["Interviewing", "Offer sent", "Screening",
                      "Interviewing", "Screening"],
        }),
        "sales_team": pd.DataFrame({
            "Agent": ["Team A", "Team B", "Team C", "Team D", "Team E"],
            "Target (EGP M)": [30, 28, 25, 22, 20],
            "Achieved (EGP M)": [34.5, 26.1, 22.0, 23.1, 12.4],
        }),
        "exit_reasons": pd.DataFrame({
            "Reason": ["Better offer", "Relocation", "Career change",
                       "Work conditions", "Personal", "Other"],
            "Leavers": [18, 9, 7, 6, 5, 3],
        }),
        "expiring": pd.DataFrame({
            "Item": ["Fixed-term contract", "Site safety certificate",
                     "Professional license", "Health insurance",
                     "Fixed-term contract"],
            "Employee / Group": ["A. Hassan", "Site crew - Project 2",
                                 "M. Salem", "All staff", "R. Adel"],
            "Expires in (days)": [9, 14, 21, 30, 38],
        }),
    }


d = load_data()
total = int(d["headcount_by_dept"]["Employees"].sum())
m = d["monthly"]
hires_now, leavers_now = int(m["Hires"].iloc[-1]), int(m["Leavers"].iloc[-1])
turnover = round(m["Leavers"].sum() / total * 100, 1)
open_positions = int(d["vacancies"]["Openings"].sum())
attendance = float(m["Attendance %"].iloc[-1])

# ---------- Styles ----------
st.markdown(
    f"""
    <style>
      .block-container {{ padding-top: 4.5rem; max-width: 1200px; }}
      .page-head {{ border-bottom: 4px solid {NAVY}; padding-bottom: .75rem;
                    margin-bottom: 1.5rem; }}
      .page-head h1 {{ color: {NAVY}; font-size: 2.1rem; margin: 0; padding: 0; }}
      .page-head p {{ color: {TEXT_MUTED}; margin: .25rem 0 0; }}
      .kpi {{ background:#fff; border:1px solid {BORDER}; border-top:3px solid {ACCENT};
              border-radius:10px; padding:1rem 1.1rem;
              box-shadow:0 1px 3px rgba(8,45,82,.08); }}
      .kpi .label {{ color:{TEXT_MUTED}; font-size:.85rem; font-weight:600; }}
      .kpi .value {{ color:{NAVY}; font-size:1.9rem; font-weight:700; line-height:1.2; }}
      .kpi .delta {{ font-size:.82rem; font-weight:600; }}
      .section {{ color:{NAVY}; font-size:1.15rem; font-weight:700;
                  margin:.25rem 0 .5rem; }}
      .note {{ background:#F3F6F9; border-left:4px solid {ACCENT}; color:{TEXT_MUTED};
               border-radius:6px; padding:.6rem .9rem; font-size:.88rem;
               margin-bottom:1.25rem; }}
    </style>
    <div class="page-head">
      <h1>Human Resources</h1>
      <p>Workforce, recruitment, attendance and sales team overview</p>
    </div>
    <div class="note">Showing sample data. Connect your real HR data in
    <b>load_data()</b> to replace it.</div>
    """,
    unsafe_allow_html=True,
)


def kpi(col, label, value, delta="", good=True):
    color = GOOD if good else BAD
    col.markdown(
        f'<div class="kpi"><div class="label">{label}</div>'
        f'<div class="value">{value}</div>'
        f'<div class="delta" style="color:{color}">{delta}&nbsp;</div></div>',
        unsafe_allow_html=True,
    )


def section(title):
    st.markdown(f'<div class="section">{title}</div>', unsafe_allow_html=True)


# ---------- KPI row ----------
c = st.columns(6, gap="small")
kpi(c[0], "Total headcount", total, f"+{hires_now - leavers_now} vs last month")
kpi(c[1], "New hires (month)", hires_now, "Target: 8", hires_now >= 8)
kpi(c[2], "Turnover (12 mo)", f"{turnover}%", "Target: under 12%", turnover < 12)
kpi(c[3], "Open positions", open_positions, "Across 5 roles", True)
kpi(c[4], "Avg. time to hire", "28 days", "Target: 30 days", True)
kpi(c[5], "Attendance", f"{attendance}%", "Target: 95%", attendance >= 95)

st.write("")

# ---------- Tabs ----------
tab1, tab2, tab3, tab4 = st.tabs(
    ["Workforce", "Recruitment", "Attendance & turnover", "Sales team"]
)

with tab1:
    left, right = st.columns(2, gap="large")
    with left:
        section("Headcount by department")
        st.bar_chart(d["headcount_by_dept"], x="Department", y="Employees",
                     color=NAVY, horizontal=True)
    with right:
        section("Hires vs. leavers (last 12 months)")
        st.bar_chart(m, x="Month", y=["Hires", "Leavers"], color=[NAVY, ACCENT],
                     stack=False)
    section("Contracts, licenses and certificates expiring soon")
    st.dataframe(d["expiring"], hide_index=True, use_container_width=True)

with tab2:
    left, right = st.columns(2, gap="large")
    with left:
        section("Hiring pipeline")
        st.bar_chart(d["pipeline"], x="Stage", y="Candidates", color=NAVY)
    with right:
        section("Open vacancies")
        st.dataframe(d["vacancies"], hide_index=True, use_container_width=True)

with tab3:
    left, right = st.columns(2, gap="large")
    with left:
        section("Attendance rate (%)")
        st.line_chart(m, x="Month", y="Attendance %", color=NAVY)
    with right:
        section("Why people leave")
        st.bar_chart(d["exit_reasons"], x="Reason", y="Leavers", color=ACCENT)

with tab4:
    sales = d["sales_team"].copy()
    sales["Achievement %"] = (sales["Achieved (EGP M)"]
                              / sales["Target (EGP M)"] * 100).round(0)
    section("Target vs. achieved by sales team")
    st.dataframe(
        sales, hide_index=True, use_container_width=True,
        column_config={
            "Achievement %": st.column_config.ProgressColumn(
                "Achievement", format="%d%%", min_value=0, max_value=120),
        },
    )
    st.bar_chart(sales, x="Agent", y=["Target (EGP M)", "Achieved (EGP M)"],
                 color=[ACCENT, NAVY], stack=False)