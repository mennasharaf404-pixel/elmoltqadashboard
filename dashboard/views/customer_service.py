"""
Customer service dashboard (Streamlit page).

Migrated from the JS dashboard "سجل عقود الوحدات والأقساط" (index.html / app.js / style.css).
Adds a Project filter (All Projects / Gardenia 3 / Gardenia Town) that is applied to the data
BEFORE any KPI, chart or table is calculated.

Data is never hardcoded: it is read from DATA_DIR (see CONFIGURATION).
"""
from __future__ import annotations

import os
import re
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

from theme import ACCENT_GOLD, ACCENT_TEAL, NAVY, NAVY_LIGHT, render_header

# =====================================================================
# CONFIGURATION  (change paths / projects here only)
# =====================================================================
# <project root>/data  (this file lives in <project root>/views/)
DATA_DIR = Path(os.environ.get("CUSTOMER_SERVICE_DATA_DIR", Path(__file__).resolve().parent.parent / "data"))

# Preferred source: the combined file produced by Power Query (Gardenia 3 + Gardenia Town).
COMBINED_FILE = DATA_DIR / "Customer_Service_All.xlsx"
COMBINED_SHEET = 0

# Fallback source (used only if COMBINED_FILE does not exist): one entry per project.
#   contracts    -> contract register (contract / handover dates, floor, area, value ...)
#   installments -> optional installment follow-up sheet (remaining installments + notes),
#                   matched to the contracts by Building + Unit.
PROJECT_SOURCES = {
    "Gardenia 3": {
        "contracts": DATA_DIR / "G3-DB-Customers.xlsx",
        "installments": DATA_DIR / "dataexcel.xlsx",
    },
    # "Gardenia Town": {"contracts": DATA_DIR / "GT-DB-Customers.xlsx", "installments": DATA_DIR / "GT-installments.xlsx"},
}

PROJECTS = ["Gardenia 3", "Gardenia Town"]  # options shown in the Project filter
ALL_PROJECTS = "All Projects"
ENDING_SOON_DAYS = 90
ARABIC_DIGITS = False  # True -> show ٠١٢٣ like the original JS (ar-EG) dashboard
DEPARTMENT = "Customer Service"

# =====================================================================
# CONSTANTS
# =====================================================================
REQUIRED_COLUMNS = [
    "project", "client_name", "building_no", "unit_no", "area_sqm",
    "contract_date", "handover_date", "remaining_installments", "notes",
]
OPTIONAL_COLUMNS = ["floor", "unit_value", "delivery_term", "contract_notes"]

CONTRACT_COLS = {  # Arabic header in the contract register -> canonical name
    "أسم العميل": "client_name", "اسم العميل": "client_name",
    "تاريخ التعاقد": "contract_date", "رقم الوحدة": "unit_no", "رقم العمارة": "building_no",
    "الدور": "floor", "المساحة": "area_sqm", "قيمة الوحدة": "unit_value",
    "تاريخ الإستلام": "handover_date", "مهلة التسليم": "delivery_term", "ملاحظات": "contract_notes",
}
INSTALLMENT_COLS = {
    "العمارة": "building_no", "الوحدة": "unit_no",
    "عدد الأقساط المتبقية": "remaining_installments", "ملاحظات": "notes",
}

PROJECT_ALIASES = {
    "gardenia 3": "Gardenia 3", "gardenia3": "Gardenia 3", "g3": "Gardenia 3", "جاردينيا 3": "Gardenia 3",
    "gardenia town": "Gardenia Town", "gardenia-town": "Gardenia Town", "town": "Gardenia Town",
    "جاردينيا تاون": "Gardenia Town",
}
FLOOR_FIXES = {"الارضي": "الأرضي", "الاول": "الأول", "الاولى": "الأولى"}  # spelling variants only

STATUS_ORDER = ["منتظم", "ملتزم", "إلى حد ما منتظم", "غير منتظم", "غير ملتزم", "متأخر", "أخرى"]
STATUS_COLORS = {
    "منتظم": "#27a85a", "ملتزم": "#17a398", "إلى حد ما منتظم": "#ee9716",
    "غير منتظم": "#d94343", "غير ملتزم": "#9b2c2c", "متأخر": "#7141c7", "أخرى": "#9aa1ae",
}
BUCKET_ORDER = ["0 (fully paid)", "1–6", "7–12", "13–24", "More than 24"]
DANGER = "#C0563B"

# =====================================================================
# PURE HELPERS (ported 1:1 from app.js where applicable)
# =====================================================================
def parse_number(value) -> float:
    """JS parseNumber: strip everything but digits . - ; blank/invalid -> NaN (JS returned 0, see notes)."""
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return float("nan")
    if isinstance(value, (int, float)):
        return float(value)
    cleaned = re.sub(r"[^\d.\-]", "", str(value).replace(",", ""))
    try:
        return float(cleaned)
    except ValueError:
        return float("nan")


def get_status(notes) -> str:
    """JS getStatus(notes): same keywords, same priority order."""
    text = "" if notes is None or (not isinstance(notes, str) and pd.isna(notes)) else str(notes)
    text = re.sub(r"\s+", " ", text).strip().lower()
    if not text:
        return "أخرى"
    if "غير ملتزم" in text:
        return "غير ملتزم"
    if any(k in text for k in ("متأخر", "متاخر", "متأخرات")):
        return "متأخر"
    if "غير منتظم" in text or "مش منتظم" in text:
        return "غير منتظم"
    if any(k in text for k in ("إلى حد ما منتظم", "الى حد ما منتظم", "الي حد ما منتظم",
                               "إلى حد ما ملتزم", "الى حد ما ملتزم", "الي حد ما ملتزم")):
        return "إلى حد ما منتظم"
    if "ملتزم" in text:
        return "ملتزم"
    if "منتظم" in text:
        return "منتظم"
    return "أخرى"


def _digits(text: str) -> str:
    return text.translate(str.maketrans("0123456789,.", "٠١٢٣٤٥٦٧٨٩٬٫")) if ARABIC_DIGITS else text


def fmt_int(x) -> str:
    return "—" if pd.isna(x) else _digits(f"{x:,.0f}")


def fmt_dec(x, digits: int = 1) -> str:
    return "—" if pd.isna(x) else _digits(f"{x:,.{digits}f}")


def _key_part(value) -> str:
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return ""
    text = re.sub(r"\s+", " ", str(value)).strip()
    return text.split(".")[0] if re.fullmatch(r"\d+\.0+", text) else text


def _ending_soon_mask(dates: pd.Series, today: pd.Timestamp) -> pd.Series:
    return dates.between(today, today + pd.Timedelta(days=ENDING_SOON_DAYS))  # NaT -> False


def _unit_label(df: pd.DataFrame) -> pd.Series:
    return "عمارة " + df["building_no"] + " – وحدة " + df["unit_no"]


# =====================================================================
# DATA LOADING
# =====================================================================
def _read_table(path: Path, anchor: str) -> pd.DataFrame:
    """Read a sheet whose header row is not necessarily the first row (finds the row containing `anchor`)."""
    raw = pd.read_excel(path, header=None, dtype=object)
    header_idx = next(
        (i for i, row in raw.iterrows() if anchor in [str(v).strip() for v in row.values]), None
    )
    if header_idx is None:
        raise ValueError(f"Header column '{anchor}' not found in {path.name}")
    table = raw.iloc[header_idx + 1:].copy()
    table.columns = [str(c).strip() for c in raw.iloc[header_idx]]
    return table.reset_index(drop=True)


def _select_renamed(table: pd.DataFrame, mapping: dict) -> pd.DataFrame:
    renamed = table.rename(columns=mapping)
    wanted = list(dict.fromkeys(mapping.values()))
    renamed = renamed.loc[:, ~renamed.columns.duplicated()]
    return renamed[[c for c in wanted if c in renamed.columns]]


def _load_project_sources(project: str, cfg: dict) -> tuple[pd.DataFrame, list[str]]:
    notes: list[str] = []
    contracts = _select_renamed(_read_table(cfg["contracts"], "تاريخ التعاقد"), CONTRACT_COLS)
    contracts = contracts.dropna(subset=["building_no", "unit_no"], how="all").reset_index(drop=True)
    contracts["_key"] = contracts["building_no"].map(_key_part) + "|" + contracts["unit_no"].map(_key_part)

    installments_path = cfg.get("installments")
    if installments_path and Path(installments_path).exists():
        inst = _select_renamed(_read_table(Path(installments_path), "عدد الأقساط المتبقية"), INSTALLMENT_COLS)
        inst = inst.dropna(subset=["building_no", "unit_no"], how="all")
        inst["_key"] = inst["building_no"].map(_key_part) + "|" + inst["unit_no"].map(_key_part)
        dup = int(inst["_key"].duplicated().sum())
        inst_unique = inst.drop_duplicates("_key", keep="first")
        merged = contracts.merge(
            inst_unique[["_key", "remaining_installments", "notes"]], on="_key", how="left"
        )
        unmatched = int(merged["remaining_installments"].isna().sum())
        orphan = len(set(inst_unique["_key"]) - set(contracts["_key"]))
        notes.append(f"{project}: {len(contracts)} عقد؛ {unmatched} عقد بدون بيانات أقساط مطابقة؛ "
                     f"{orphan} وحدة في ملف الأقساط بدون عقد؛ {dup} وحدة مكررة في ملف الأقساط (أُخذ أول سجل).")
        contracts = merged
    else:
        contracts["remaining_installments"] = pd.NA
        contracts["notes"] = pd.NA
        notes.append(f"{project}: لا يوجد ملف أقساط — الأقساط والحالة غير متاحة.")

    contracts["project"] = project
    return contracts.drop(columns="_key"), notes


def _load_combined(path: Path) -> tuple[pd.DataFrame, list[str]]:
    df = pd.read_excel(path, sheet_name=COMBINED_SHEET)
    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"{path.name} is missing required columns: {', '.join(missing)}")
    return df, [f"المصدر: {path.name}"]


@st.cache_data(show_spinner="جاري تحميل البيانات...")
def load_data(signature: tuple) -> tuple[pd.DataFrame, list[str]]:
    """`signature` (file paths + modified times) only exists so the cache refreshes when files change."""
    if COMBINED_FILE.exists():
        df, notes = _load_combined(COMBINED_FILE)
    else:
        frames, notes = [], []
        for project, cfg in PROJECT_SOURCES.items():
            if Path(cfg["contracts"]).exists():
                frame, frame_notes = _load_project_sources(project, cfg)
                frames.append(frame)
                notes += frame_notes
        if not frames:
            raise FileNotFoundError("No data source found")
        df = pd.concat(frames, ignore_index=True)
    return clean_data(df), notes


def _source_signature() -> tuple:
    paths = [COMBINED_FILE] + [Path(p) for cfg in PROJECT_SOURCES.values() for p in cfg.values()]
    return tuple((str(p), p.stat().st_mtime if p.exists() else None) for p in paths)


def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """Safe, meaning-preserving cleaning only (whitespace, spelling variants, types). Dates are NOT altered."""
    df = df.copy()
    for col in OPTIONAL_COLUMNS:
        if col not in df.columns:
            df[col] = pd.NA

    for col in ["project", "client_name", "building_no", "unit_no", "floor", "notes", "contract_notes"]:
        df[col] = df[col].map(_key_part)
    df["project"] = df["project"].map(lambda p: PROJECT_ALIASES.get(p.lower(), p))
    df["floor"] = df["floor"].replace(FLOOR_FIXES)

    for col in ["contract_date", "handover_date"]:
        df[col] = pd.to_datetime(df[col], errors="coerce")  # already real dates in the source
    df["area_sqm"] = df["area_sqm"].map(parse_number)
    df["unit_value"] = pd.to_numeric(df["unit_value"], errors="coerce")
    df["remaining_installments"] = df["remaining_installments"].map(parse_number)

    df["status_category"] = df["notes"].map(get_status)
    df["unit_label"] = _unit_label(df)
    df["search_text"] = (df["client_name"] + " " + df["notes"] + " " + df["contract_notes"]).str.lower()
    return df


# =====================================================================
# FILTERING
# =====================================================================
def _scope(df: pd.DataFrame, project: str) -> pd.DataFrame:
    return df if project == ALL_PROJECTS else df[df["project"] == project]


def _sorted_buildings(values) -> list[str]:
    return sorted({v for v in values if v}, key=lambda v: (0, int(v)) if v.isdigit() else (1, v))


def render_filters(df: pd.DataFrame) -> dict:
    """Draw the filter bar; options are derived from the selected project so they never show stale choices."""
    project = st.radio("Project", [ALL_PROJECTS] + PROJECTS, horizontal=True, key="cs_project")
    scope = _scope(df, project)
    k = project  # widget keys include the project so selections never clash when options change

    c1, c2, c3, c4 = st.columns(4)
    present = [s for s in STATUS_ORDER if s in set(scope["status_category"])]
    statuses = c1.multiselect("Status", present, key=f"cs_status_{k}", placeholder="All statuses")
    buildings = c2.multiselect("Building", _sorted_buildings(scope["building_no"]),
                               key=f"cs_building_{k}", placeholder="All buildings")
    floors = c3.multiselect("Floor", sorted({f for f in scope["floor"] if f}),
                            key=f"cs_floor_{k}", placeholder="All floors")
    years = sorted(scope["handover_date"].dropna().dt.year.unique())
    year = c4.selectbox("Handover year", ["All"] + [int(y) for y in years], key=f"cs_year_{k}")

    c5, c6 = st.columns([3, 1])
    search = c5.text_input("Search client name or notes", key=f"cs_search_{k}", placeholder="Type to search…")
    soon_only = c6.checkbox(f"Handover within {ENDING_SOON_DAYS} days only", key=f"cs_soon_{k}")
    return {"project": project, "statuses": statuses, "buildings": buildings, "floors": floors,
            "year": year, "search": search.strip().lower(), "soon_only": soon_only}


def filter_data(df: pd.DataFrame, f: dict, today: pd.Timestamp) -> pd.DataFrame:
    """Project and every other filter are applied here, before ANY metric is computed."""
    out = _scope(df, f["project"])
    if f["statuses"]:
        out = out[out["status_category"].isin(f["statuses"])]
    if f["buildings"]:
        out = out[out["building_no"].isin(f["buildings"])]
    if f["floors"]:
        out = out[out["floor"].isin(f["floors"])]
    if f["year"] != "All":
        out = out[out["handover_date"].dt.year == f["year"]]
    if f["soon_only"]:
        out = out[_ending_soon_mask(out["handover_date"], today)]
    if f["search"]:
        out = out[out["search_text"].str.contains(f["search"], regex=False, na=False)]
    return out


# =====================================================================
# KPIs
# =====================================================================
def calculate_kpis(df: pd.DataFrame, today: pd.Timestamp) -> dict:
    total = len(df)
    paid_off = int((df["remaining_installments"] == 0).sum())
    return {
        "total": total,
        "area": df["area_sqm"].sum(),
        "avg_installments": df["remaining_installments"].mean(),  # unknown values are skipped
        "buildings": len(df.loc[df["building_no"] != "", ["project", "building_no"]].drop_duplicates()),
        "completion": paid_off / total * 100 if total else 0.0,
        "ending_soon": int(_ending_soon_mask(df["handover_date"], today).sum()),
    }


def render_kpis(k: dict) -> None:
    """KPI row using st.metric so the cards pick up the shared theme styling."""
    cards = [
        ("Total Contracts", fmt_int(k["total"]), "All registered contracts"),
        ("Total Area (m²)", fmt_int(k["area"]), "Combined unit area"),
        ("Avg. Remaining Installments", fmt_dec(k["avg_installments"]),
         "Average across contracts with installment data"),
        ("Buildings", fmt_int(k["buildings"]), "Distinct buildings"),
        ("Fully Paid", fmt_int(k["completion"]) + "%", "Contracts with no remaining installments"),
        (f"Handover in {_digits(str(ENDING_SOON_DAYS))} Days", fmt_int(k["ending_soon"]),
         "Contracts with an upcoming handover date"),
    ]
    for start in (0, 3):  # two rows of three so labels and numbers are never truncated
        for col, (label, value, hint) in zip(st.columns(3), cards[start:start + 3]):
            col.metric(label, value, help=hint)


# =====================================================================
# CHARTS
# =====================================================================
def _style(fig, height: int = 300):
    fig.update_layout(
        height=height, margin=dict(l=8, r=8, t=8, b=8), paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)", font=dict(size=13, color=NAVY), legend=dict(title=None),
    )
    fig.update_xaxes(showgrid=False, title=None)
    fig.update_yaxes(gridcolor="#E4E8EA", title=None)
    return fig


def _chart_card(title: str, subtitle: str, fig=None, empty_msg: str = "Not enough data for this chart") -> None:
    with st.container(border=True):
        st.markdown(f"**{title}**")
        st.caption(subtitle)
        if fig is None:
            st.info(empty_msg)
        else:
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def _status_fig(df):
    counts = df["status_category"].value_counts().rename_axis("Status").reset_index(name="Contracts")
    fig = px.pie(counts, names="Status", values="Contracts", hole=0.62, color="Status",
                 color_discrete_map=STATUS_COLORS)
    fig.update_traces(textinfo="value+percent", sort=False)
    return _style(fig).update_layout(legend=dict(orientation="v", x=1.0))


def _building_fig(df):
    multi = df["project"].nunique() > 1
    label = (df["project"] + " · " + df["building_no"]) if multi else df["building_no"].replace("", "N/A")
    counts = df.assign(label=label).groupby(["label", "project"]).size().reset_index(name="Units")
    order = counts.groupby("label")["Units"].sum().sort_values(ascending=False).index.tolist()
    fig = px.bar(counts, x="label", y="Units", color="project" if multi else None, text="Units",
                 category_orders={"label": order}, color_discrete_sequence=[NAVY, ACCENT_GOLD],
                 labels={"label": "Building", "project": "Project"})
    fig.update_traces(textposition="outside", cliponaxis=False)
    if not multi:
        fig.update_traces(marker_color=NAVY)
    return _style(fig).update_xaxes(type="category")


def _bucket(n: float) -> str | None:
    if pd.isna(n):
        return None
    if n == 0:
        return "0 (fully paid)"
    return "1–6" if n <= 6 else "7–12" if n <= 12 else "13–24" if n <= 24 else "More than 24"


def _installment_fig(df):
    buckets = df["remaining_installments"].map(_bucket).dropna()
    if buckets.empty:
        return None
    counts = buckets.value_counts().reindex(BUCKET_ORDER, fill_value=0).rename_axis("Installments").reset_index(name="Contracts")
    fig = px.bar(counts, x="Installments", y="Contracts", text="Contracts")
    fig.update_traces(marker_color=NAVY_LIGHT, textposition="outside", cliponaxis=False)
    return _style(fig)


def _quarter_fig(df):
    dates = df["handover_date"].dropna()
    if dates.empty:
        return None
    counts = dates.dt.to_period("Q").value_counts().sort_index()
    data = pd.DataFrame({"Quarter": [f"{p.year} Q{p.quarter}" for p in counts.index], "Contracts": counts.values})
    fig = px.bar(data, x="Quarter", y="Contracts", text="Contracts")
    fig.update_traces(marker_color=ACCENT_TEAL, textposition="outside", cliponaxis=False)
    return _style(fig).update_xaxes(type="category")


def _top_installments_fig(df):
    top = df[df["remaining_installments"] > 0].nlargest(8, "remaining_installments")
    if top.empty:
        return None
    fig = px.bar(top, x="remaining_installments", y="unit_label", orientation="h",
                 text="remaining_installments", hover_data={"project": True, "unit_label": False},
                 labels={"remaining_installments": "Remaining installments", "project": "Project"})
    fig.update_traces(marker_color=DANGER, textposition="outside", cliponaxis=False)
    return _style(fig).update_yaxes(autorange="reversed", gridcolor="rgba(0,0,0,0)")


def render_charts(df: pd.DataFrame) -> None:
    c1, c2 = st.columns(2)
    with c1:
        _chart_card("Contract Status", "Based on the status recorded in the notes", _status_fig(df))
    with c2:
        _chart_card("Remaining Installments", "Contracts grouped by installments left", _installment_fig(df),
                    "No installment data for the selected contracts")
    _chart_card("Units per Building", "Distribution of contracts across buildings", _building_fig(df))
    c3, c4 = st.columns(2)
    with c3:
        _chart_card("Handovers by Quarter", "Contracts by handover quarter", _quarter_fig(df),
                    "No handover dates for the selected contracts")
    with c4:
        _chart_card("Top 8 Units by Remaining Installments", "Units with the most installments left",
                    _top_installments_fig(df), "No remaining installments for the selected contracts")


# =====================================================================
# TABLES & INSIGHTS
# =====================================================================
def _mini_table(df: pd.DataFrame, columns: dict, empty_msg: str) -> None:
    if df.empty:
        st.caption(empty_msg)
        return
    view = df[list(columns)].rename(columns=columns)
    st.dataframe(view, hide_index=True, use_container_width=True, height=250, column_config={
        "Handover": st.column_config.DateColumn(format="YYYY-MM-DD"),
        "Remaining": st.column_config.NumberColumn(format="%d"),
    })


def render_summary_tables(df: pd.DataFrame, today: pd.Timestamp) -> None:
    with st.container(border=True):
        st.markdown("**Upcoming Handovers**")
        st.caption("Next handover dates (past dates excluded)")
        upcoming = df[df["handover_date"] >= today].nsmallest(6, "handover_date")
        _mini_table(upcoming, {"client_name": "Client", "building_no": "Building", "unit_no": "Unit",
                               "handover_date": "Handover", "remaining_installments": "Remaining"},
                    "No upcoming handovers")
    with st.container(border=True):
        st.markdown("**Highest Remaining Installments**")
        st.caption("Top 6 contracts")
        top = df[df["remaining_installments"] > 0].nlargest(6, "remaining_installments")
        _mini_table(top, {"client_name": "Client", "building_no": "Building", "unit_no": "Unit",
                          "remaining_installments": "Remaining"}, "No remaining installments")


def render_insights(df: pd.DataFrame) -> None:
    b = df.loc[df["building_no"] != ""].groupby(["project", "building_no"]).size().sort_values(ascending=False)
    largest = df.loc[df["area_sqm"].idxmax()] if df["area_sqm"].notna().any() else None
    top_status = df["status_category"].value_counts()
    hi = df.loc[df["remaining_installments"].idxmax()] if df["remaining_installments"].notna().any() else None
    items = [
        ("Largest Building", f"Building {b.index[0][1]}" if len(b) else "—", f"{fmt_int(b.iloc[0])} units" if len(b) else ""),
        ("Largest Unit by Area", f"{fmt_int(largest['area_sqm'])} m²" if largest is not None else "—",
         largest["unit_label"] if largest is not None else ""),
        ("Most Common Status", top_status.index[0] if len(top_status) else "—",
         f"{fmt_int(top_status.iloc[0])} contracts" if len(top_status) else ""),
        ("Most Installments Left", fmt_int(hi["remaining_installments"]) if hi is not None else "—",
         hi["unit_label"] if hi is not None else ""),
    ]
    for start in (0, 2):
        for col, (label, value, sub) in zip(st.columns(2), items[start:start + 2]):
            col.metric(label, value)
            col.caption(sub)


def render_data_table(df: pd.DataFrame, today: pd.Timestamp) -> None:
    st.subheader("Contract Register")
    columns = {
        "project": "Project", "client_name": "Client", "building_no": "Building", "unit_no": "Unit",
        "floor": "Floor", "area_sqm": "Area (m²)", "remaining_installments": "Remaining Installments",
        "contract_date": "Contract Date", "handover_date": "Handover Date",
        "status_category": "Status", "notes": "Notes",
    }
    view = df.sort_values(["project", "building_no", "unit_no"]).reset_index(drop=True)
    soon = _ending_soon_mask(view["handover_date"], today)
    view = view[list(columns)].rename(columns=columns)
    styled = view.style.apply(
        lambda row: ["background-color: rgba(192,86,59,.10)" if soon[row.name] else ""] * len(row), axis=1)
    st.caption(f"{fmt_int(len(view))} contracts · highlighted rows hand over within {ENDING_SOON_DAYS} days")
    st.dataframe(styled, hide_index=True, use_container_width=True, height=460, column_config={
        "Contract Date": st.column_config.DateColumn(format="YYYY-MM-DD"),
        "Handover Date": st.column_config.DateColumn(format="YYYY-MM-DD"),
        "Area (m²)": st.column_config.NumberColumn(format="%d"),
        "Remaining Installments": st.column_config.NumberColumn(format="%d"),
    })
    st.download_button("⬇ Export CSV", view.to_csv(index=False).encode("utf-8-sig"),
                       file_name="contracts_filtered.csv", mime="text/csv")


# =====================================================================
# PAGE
# =====================================================================
# Page config, global CSS and the logo are handled once in app.py / theme.py.
def main() -> None:
    render_header(DEPARTMENT)
    try:
        df, notes = load_data(_source_signature())
    except FileNotFoundError:
        st.error(f"Data file not found. Expected {COMBINED_FILE} or the project files listed in PROJECT_SOURCES.")
        st.stop()
    except ValueError as exc:
        st.error(f"Problem with the data file structure: {exc}")
        st.stop()

    today = pd.Timestamp.today().normalize()
    filters = render_filters(df)
    data = filter_data(df, filters, today)

    if data.empty:
        has_project_rows = filters["project"] == ALL_PROJECTS or (df["project"] == filters["project"]).any()
        st.info("No records match the selected filters." if has_project_rows
                else f"No data loaded yet for {filters['project']}.")
        st.stop()

    render_kpis(calculate_kpis(data, today))
    st.write("")
    render_charts(data)
    st.write("")
    render_summary_tables(data, today)
    st.write("")
    render_insights(data)
    st.write("")
    render_data_table(data, today)

    with st.expander("Data quality"):
        for note in notes:
            st.write(note)


main()