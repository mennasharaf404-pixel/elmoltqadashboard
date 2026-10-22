from __future__ import annotations
import hashlib
import io
import re
from datetime import date, datetime
from pathlib import Path
import altair as alt
import numpy as np
import pandas as pd
import streamlit as st
from theme import render_header

# Configuration
_HERE = Path(__file__).resolve().parent
DATA_CANDIDATES = [
    _HERE.parent / "data" / "master.xlsx",
    _HERE.parent / "master.xlsx",
    Path("data/master.xlsx"),
    Path("master.xlsx"),
]
# data column names
NAME, NID, GENDER = "الاسم", "الرقم القومي", "M/F"
CODE, FP = "كود وظيفي", "كود البصمة"
BIRTH, HIRE, END = "تاريخ الميلاد", "تاريخ التعيين", "تاريخ أنهاء التعاقد"
JOB, DEPT, MGMT, SECTOR = "الوظيفة", "القسم", "الإدارة", "القطاع"
SALARY, BANK = "المرتب الأساسي", "الحساب البنكي"
HEALTH_INS, SOC_INS = "رقم التأمين الصحي", "التأمين الاجتماعي"
PHONE, PHONE2, EMAIL, ADDR = "رقم الهاتف", "رقم الهاتف2", "البريد الإلكتروني", "العنوان"
EDU = "المؤهل الدراسي"

# Derived columns
AGE, TENURE = "العمر", "سنوات الخدمة"

# Checklist columns: "complete" when "مكتمل".
DOC_COLS = [
    "صورة البطاقة", "شهادة الجيش", "شهادة الميلاد", "الصور الشخصية",
    "فيش جنائي", "كعب العمل", "صورة من برينت التأمينات", "المؤهل",
]
ARMY_DOC = "شهادة الجيش"  # male employees
DONE = "مكتمل"

ID_COLS = [NID, FP, BANK, HEALTH_INS, SOC_INS, PHONE, PHONE2, CODE]

DEL = "🗑️ Delete"
DIR_COLS = [CODE, NAME, GENDER, JOB, DEPT, MGMT, SECTOR, HIRE, TENURE, AGE, SALARY,
            NID, BANK, HEALTH_INS, SOC_INS, PHONE, PHONE2, EMAIL, ADDR, EDU] + DOC_COLS

TYPO_FIXES = {SECTOR: {"قطاع التشقطاع التشغيل والعملياتغيل": "قطاع التشغيل والعمليات"}}

# Cleaning functions
def _clean_text(v):
    """Trim, collapse whitespace, turn empty strings into None."""
    if isinstance(v, str):
        t = re.sub(r"\s+", " ", v.replace("\u00a0", " ")).strip()
        return t or None
    return v


def _to_id(v):
    """Identifiers are text, never numbers (keeps all digits, no 1.2E+13)."""
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return None
    if isinstance(v, (int, np.integer)):
        return str(int(v))
    if isinstance(v, (float, np.floating)):
        return str(int(v)) if float(v).is_integer() else str(v)
    return _clean_text(v)


def _fix_phone(v):
    """Excel dropped the leading zero from Egyptian mobile numbers."""
    if isinstance(v, str) and len(v) == 10 and v.startswith("1") and v.isdigit():
        return "0" + v
    return v


def _to_date(v):
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return pd.NaT
    if isinstance(v, (datetime, date, pd.Timestamp)):
        return pd.Timestamp(v)
    return pd.to_datetime(str(v), dayfirst=True, errors="coerce")  # e.g. "13-7-2026"


def _birth_from_nid(n):
    if not isinstance(n, str) or len(n) != 14 or not n.isdigit():
        return pd.NaT
    year = (1900 if n[0] == "2" else 2000) + int(n[1:3])
    try:
        return pd.Timestamp(year, int(n[3:5]), int(n[5:7]))
    except ValueError:
        return pd.NaT


@st.cache_data(show_spinner="Cleaning data...")
def load_data(raw: bytes, as_of: str) -> pd.DataFrame:
    """Read + clean the master sheet. `as_of` makes the cache roll over daily."""
    today = pd.Timestamp(as_of)
    df = pd.read_excel(io.BytesIO(raw), sheet_name=0, dtype=object)

    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    df.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in df.columns]

    df = df.apply(lambda col: col.map(_clean_text))
    df = df[df[NAME].notna()].reset_index(drop=True)

    for col, fixes in TYPO_FIXES.items():
        df[col] = df[col].replace(fixes)

    for col in ID_COLS:
        df[col] = df[col].map(_to_id)
    df[PHONE] = df[PHONE].map(_fix_phone)
    df[PHONE2] = df[PHONE2].map(_fix_phone)
    df[FP] = df[FP].where(df[FP] != "Not Assigned", None)

    df[HIRE] = df[HIRE].map(_to_date)
    df[END] = df[END].map(_to_date)
    from_nid = df[NID].map(_birth_from_nid)
    df[BIRTH] = from_nid.where(from_nid.notna(), df[BIRTH].map(_to_date))
    df[SALARY] = pd.to_numeric(df[SALARY], errors="coerce")

    df[AGE] = ((today - df[BIRTH]).dt.days / 365.25).round(1)
    df[TENURE] = ((today - df[HIRE]).dt.days / 365.25).round(1)
    return df

# Charts #
def sector_chart(view: pd.DataFrame):
    """Headcount share vs payroll share per sector."""
    g = view.groupby(SECTOR).agg(hc=(NAME, "size"), pay=(SALARY, "sum"))
    if g.empty:
        return None
    order = g.sort_values("hc", ascending=False).index.tolist()
    long = pd.DataFrame({"sector": g.index, "Headcount": g["hc"] / g["hc"].sum()})
    if g["pay"].sum() > 0:
        long["Payroll"] = g["pay"] / g["pay"].sum()
    long = long.melt("sector", var_name="metric", value_name="share")

    base = alt.Chart(long).encode(
        y=alt.Y("sector:N", sort=order, title=None, axis=alt.Axis(labelLimit=280)),
        yOffset="metric:N",
        x=alt.X("share:Q", title=None, axis=alt.Axis(format="%")),
    )
    bars = base.mark_bar().encode(
        color=alt.Color("metric:N", title=None, legend=alt.Legend(orient="top")),
        tooltip=["sector:N", "metric:N", alt.Tooltip("share:Q", format=".1%")],
    )
    labels = base.mark_text(align="left", dx=3, fontSize=10).encode(text=alt.Text("share:Q", format=".0%"))
    return (bars + labels).properties(
        title="Headcount share vs payroll share, by sector",
        height=max(140, 52 * len(order)), width="container",
    )

def missing_docs_chart(view: pd.DataFrame):
    """How many employees are still missing each document (anything not marked complete)."""
    rows = []
    for col in DOC_COLS:
        if col not in view.columns:
            continue
        eligible = view[view[GENDER] == "Male"] if col == ARMY_DOC else view  # males only
        if eligible.empty:
            continue
        missing = int((eligible[col] != DONE).sum())
        rows.append({"document": col, "missing": missing, "share": missing / len(eligible),
                     "eligible": len(eligible)})
    if not rows:
        return None
    d = pd.DataFrame(rows)
    d["txt"] = d["missing"].astype(str) + " (" + (d["share"] * 100).round(0).astype(int).astype(str) + "%)"
    base = alt.Chart(d).encode(
        y=alt.Y("document:N", sort="-x", title=None, axis=alt.Axis(labelLimit=280)),
        x=alt.X("missing:Q", title=None),
    )
    bars = base.mark_bar(color="crimson").encode(
        tooltip=["document:N", alt.Tooltip("missing:Q", title="Missing"),
                 alt.Tooltip("share:Q", title="Share of eligible", format=".0%"),
                 alt.Tooltip("eligible:Q", title="Eligible employees")]
    )
    labels = base.mark_text(align="left", dx=3, fontSize=10).encode(text="txt:N")
    return (bars + labels).properties(
        title="Missing documents (employees without a completed item)",
        height=max(160, 32 * len(d)), width="container",
    )

def hiring_chart(view: pd.DataFrame, today: pd.Timestamp):
    """Monthly hires, last 24 months, with a 3-month average line."""
    h = view[HIRE].dropna()
    if h.empty:
        return None
    periods = pd.period_range(end=today.to_period("M"), periods=24, freq="M")
    s = h.dt.to_period("M").value_counts().reindex(periods, fill_value=0).sort_index()
    d = pd.DataFrame({"month": s.index.to_timestamp(), "hires": s.values})
    d["avg3"] = d["hires"].rolling(3, min_periods=1).mean().round(1)
    base = alt.Chart(d).encode(x=alt.X("month:T", title=None, axis=alt.Axis(format="%b %y")))
    bars = base.mark_bar(opacity=0.6).encode(
        y=alt.Y("hires:Q", title=None),
        tooltip=[alt.Tooltip("month:T", format="%b %Y"), "hires:Q", "avg3:Q"],
    )
    line = base.mark_line(color="darkorange", size=3).encode(y="avg3:Q")
    return (bars + line).properties(
        title="Monthly hires, last 24 months (line = 3-month average)", height=300, width="container"
    )

def pay_tenure_chart(view: pd.DataFrame):
    """Each dot is an employee; the red line shows whether pay rises with service."""
    d = view[[TENURE, SALARY, SECTOR, JOB]].dropna(subset=[TENURE, SALARY]).copy()
    if len(d) < 5:
        return None
    d.columns = ["tenure", "salary", "sector", "job"]
    pts = alt.Chart(d).mark_circle(size=55, opacity=0.55).encode(
        x=alt.X("tenure:Q", title="Years of service"),
        y=alt.Y("salary:Q", title=None),
        tooltip=["job:N", "sector:N", "tenure:Q", alt.Tooltip("salary:Q", format=",.0f")],
    )
    trend = pts.transform_regression("tenure", "salary").mark_line(color="crimson", size=3)
    return (pts + trend).properties(title="Salary vs years of service", height=300, width="container")


def show(chart):
    if chart is not None:
        st.altair_chart(chart)

# Page

def _load_source():
    """Uploaded file wins; otherwise look for master.xlsx in the usual places."""
    with st.sidebar:
        st.subheader("Data")
        up = st.file_uploader("Upload HR Data (.xlsx file)", type=["xlsx"])
    if up is not None:
        return up.getvalue(), up.name
    for p in DATA_CANDIDATES:
        if p.exists():
            return p.read_bytes(), str(p)
    return None, None

def main():
    render_header("HR")

    raw, _ = _load_source()
    if raw is None:
        st.warning("upload the Data")
        st.stop()

    today = pd.Timestamp.today().normalize()
    src_key = (hashlib.md5(raw).hexdigest(), str(today.date()))
    if st.session_state.get("src_key") != src_key:  # new file -> reload
        try:
            st.session_state.data = load_data(raw, str(today.date())).copy()
        except Exception as exc: 
            st.error(f"Could not read the file ({exc}).")
            st.stop()
        st.session_state.src_key = src_key
        st.session_state.ver = 0
    df = st.session_state.data

    # Filters #
    with st.sidebar:
        st.subheader("Filters")
        sectors = st.multiselect("Sector", sorted(df[SECTOR].dropna().unique()))
        mgmts = st.multiselect("Management", sorted(df[MGMT].dropna().unique()))
        genders = st.multiselect("Gender", sorted(df[GENDER].dropna().unique()))
        if st.button("↩️ Reset all edits"):
            st.session_state.data = load_data(raw, str(today.date())).copy()
            st.session_state.ver += 1
            st.rerun()
    view = df
    if sectors:
        view = view[view[SECTOR].isin(sectors)]
    if mgmts:
        view = view[view[MGMT].isin(mgmts)]
    if genders:
        view = view[view[GENDER].isin(genders)]
    if view.empty:
        st.info("No employees match the current filters.")
        st.stop()

    #KPIs#
    sal = view[SALARY].dropna()
    hired_12m = int((view[HIRE] >= today - pd.DateOffset(months=12)).sum())
    k = st.columns(5)
    k[0].metric("Headcount", f"{len(view):,}")
    k[1].metric("Hired last 12 months", f"{hired_12m:,}")
    k[2].metric("Median salary", f"{sal.median():,.0f}" if len(sal) else "-")
    k[3].metric("Avg tenure (yrs)", f"{view[TENURE].mean():.1f}" if view[TENURE].notna().any() else "-")
    k[4].metric("Female share", f"{(view[GENDER] == 'Female').mean():.0%}")

    #Charts#
    c1, c2 = st.columns(2)
    with c1:
        show(sector_chart(view))
    with c2:
        show(missing_docs_chart(view))

    c3, c4 = st.columns(2)
    with c3:
        show(hiring_chart(view, today))
    with c4:
        show(pay_tenure_chart(view))

    #Employee directory#
    st.subheader("Employee directory")
    st.caption(
        "✏️ Double-click to edit · 🗑️ Tick a row to delete · ➕ Add new staff · 💾 Save to update charts. "
        "📥 download before you leave."
    )
    q = st.text_input("Search name or job code")
    tbl = view
    if q:
        tbl = view[view[NAME].fillna("").str.contains(q, case=False, regex=False)
                   | view[CODE].fillna("").str.contains(q, regex=False)]
    cols = [c for c in DIR_COLS if c in df.columns]
    shown = tbl[cols].copy()
    shown.insert(0, DEL, False)

    sig = abs(hash((q, tuple(sectors), tuple(mgmts), tuple(genders))))
    edited = st.data_editor(
        shown,
        hide_index=True,
        key=f"dir_{st.session_state.ver}_{sig}",
        disabled=[TENURE, AGE],
        column_config={
            DEL: st.column_config.CheckboxColumn(DEL, help="click to delete this employee", default=False),
            HIRE: st.column_config.DateColumn(HIRE, format="YYYY-MM-DD"),
            SALARY: st.column_config.NumberColumn(SALARY, format="%.0f"),
        },
    )

    n_del = int(edited[DEL].sum())
    b1, b2, b3 = st.columns([1, 1, 1])
    if b1.button(f"💾 Save changes ({n_del} to delete)", type="primary"):
        base = st.session_state.data.copy()
        for c in cols:
            if c in (TENURE, AGE):
                continue
            if c == HIRE:
                base.loc[edited.index, c] = pd.to_datetime(edited[c], errors="coerce")
            elif c == SALARY:
                base.loc[edited.index, c] = pd.to_numeric(edited[c], errors="coerce")
            else:
                base.loc[edited.index, c] = edited[c]
        base = base.drop(index=edited.index[edited[DEL]])
        base[TENURE] = ((today - base[HIRE]).dt.days / 365.25).round(1)
        st.session_state.data = base
        st.session_state.ver += 1
        st.rerun()

    with b2.popover("➕ Add employee"):
        with st.form("add_employee", clear_on_submit=True):
            f_name = st.text_input("Name *")
            f_gender = st.selectbox("Gender", ["Male", "Female"])
            f_code = st.text_input("Job code")
            f_job = st.text_input("Job title")
            f_sector = st.selectbox("Sector", [""] + sorted(df[SECTOR].dropna().unique()))
            f_mgmt = st.selectbox("Management", [""] + sorted(df[MGMT].dropna().unique()))
            f_dept = st.selectbox("Department", [""] + sorted(df[DEPT].dropna().unique()))
            f_hire = st.date_input("Hire date", value=today.date())
            f_salary = st.number_input("Basic salary", min_value=0.0, step=100.0)
            f_nid = st.text_input("National ID (14 digits)")
            f_phone = st.text_input("Phone")
            f_email = st.text_input("Email")
            submitted = st.form_submit_button("Add employee", type="primary")
        if submitted:
            if not f_name.strip():
                st.error("Name is required.")
            else:
                base = st.session_state.data
                row = {c: None for c in base.columns}
                nid = _to_id(f_nid)
                birth = _birth_from_nid(nid)
                hire = pd.Timestamp(f_hire)
                row.update({
                    NAME: _clean_text(f_name), GENDER: f_gender, CODE: _to_id(f_code),
                    JOB: _clean_text(f_job), SECTOR: f_sector or None, MGMT: f_mgmt or None,
                    DEPT: f_dept or None, HIRE: hire, END: pd.NaT, BIRTH: birth,
                    SALARY: f_salary if f_salary > 0 else np.nan, NID: nid,
                    PHONE: _fix_phone(_to_id(f_phone)), EMAIL: _clean_text(f_email),
                    TENURE: round((today - hire).days / 365.25, 1),
                    AGE: round((today - birth).days / 365.25, 1) if pd.notna(birth) else np.nan,
                })
                new_idx = int(base.index.max()) + 1 if len(base) else 0
                st.session_state.data = pd.concat(
                    [base, pd.DataFrame([row], index=[new_idx])]
                )
                st.session_state.ver += 1
                st.rerun()

    export = st.session_state.data[cols].copy()
    export[HIRE] = export[HIRE].dt.date
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        export.to_excel(xw, index=False, sheet_name="Employees")
    b3.download_button(
        "⬇️ Download Excel",
        buf.getvalue(),
        file_name="hr_employees.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
main()