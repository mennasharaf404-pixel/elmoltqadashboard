from __future__ import annotations
import ast
import difflib
import math
import operator
import os
import re
import shutil
from copy import copy
from io import BytesIO
from pathlib import Path
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from theme import render_header

# CONFIGURATION  
DEPARTMENT = "Operations"
DATA_DIR = Path(os.environ.get("OPERATIONS_DATA_DIR", Path(__file__).resolve().parent.parent / "data"))
DATA_FILE = DATA_DIR / "All Inventory Project(1).xlsx"
SHEET = "Gardenia Town"
HEADER_ROW = 5  # Excel row that holds the column headers (pandas header=4)

TEXT_COLS = ["Phase", "Unit Code", "Unit Type", "Building", "Apartment NO.", "Type", "STATUS", "Name of client"]
NUMERIC_COLS = ["Floor", "Rooms Num", "In/Area", "NEW M.PRICE", "M.PRICE", "Total Unit"]
CHOICE_COLS = ["Phase", "Type", "STATUS", "Unit Type", "Building"]
REQUIRED_COLS = ["Phase", "Unit Code", "Unit Type", "Building", "Floor", "STATUS"]
REQUIRED_ON_SAVE = ["Phase", "Unit Code", "Unit Type", "Building", "STATUS"]
DISPLAY_COLS = ["Phase", "Unit Code", "Unit Type", "Building", "Floor", "Apartment NO.", "Type",
                "In/Area", "NEW M.PRICE", "STATUS", "Rooms Num", "Name of client"]

STATUS_LABELS = {"SOLD": "Sold", "AVAILABLE": "Available", "HOLD": "Hold", "RESERVED": "Reserved"}
STATUS_ORDER = ["Sold", "Reserved", "Hold", "Available"]
STATUS_COLORS = {"Sold": "#1F4E9A", "Reserved": "#7B3FBF", "Hold": "#F28C0F", "Available": "#12A150"}
UNSOLD = ["Available", "Hold", "Reserved"]
FILTER_PREFIX = "ops_f_"
NEW_OPTION = "➕ Add new…"
DEL = "🗑️ Delete"  # checkbox column in the editable table
CANONICAL_STATUS = list(STATUS_LABELS)  # SOLD, AVAILABLE, HOLD, RESERVED

#HELPERS
def _has(df: pd.DataFrame, col: str) -> bool:
    return col in df.columns and bool(df[col].notna().any())


def _clean(value):
    """NaN / pd.NA / numpy scalars -> plain Python values."""
    if value is None or value is pd.NA:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    return value.item() if isinstance(value, np.generic) else value


def _same(a, b) -> bool:
    a, b = _clean(a), _clean(b)
    if a is None or (isinstance(a, str) and not a.strip()):
        return b is None or (isinstance(b, str) and not b.strip())
    if b is None:
        return False
    try:
        return bool(np.isclose(float(a), float(b)))
    except (TypeError, ValueError):
        return str(a).strip() == str(b).strip()


def _nat_key(value) -> list:
    """Natural sort so 'B2' comes before 'B10'."""
    return [int(p) if p.isdigit() else p.lower() for p in re.split(r"(\d+)", str(value))]


def fmt_int(x) -> str:
    return "—" if pd.isna(x) else f"{x:,.0f}"


def fmt_money(x) -> str:
    if pd.isna(x):
        return "—"
    if abs(x) >= 1e9:
        return f"{x / 1e9:,.2f}B"
    return f"{x / 1e6:,.2f}M" if abs(x) >= 1e6 else f"{x / 1e3:,.0f}K" if abs(x) >= 1e3 else f"{x:,.0f}"


def _price_col(df: pd.DataFrame) -> str | None:
    return next((c for c in ("NEW M.PRICE", "M.PRICE") if _has(df, c)), None)


# =====================================================================
# DATA LOADING & CLEANING
# =====================================================================
def clean_inventory(df: pd.DataFrame) -> pd.DataFrame:
    """Whitespace / type cleaning only. Values are never changed."""
    df = df.dropna(how="all").copy()
    df.columns = [str(c).strip() for c in df.columns]
    empty_unnamed = [c for c in df.columns if c.startswith("Unnamed") and df[c].isna().all()]
    df = df.drop(columns=empty_unnamed)
    for col in TEXT_COLS:
        if col in df.columns:
            df[col] = df[col].astype("string").str.strip()
    if "Phase" in df.columns:
        df["Phase"] = df["Phase"].str.upper()
    for col in NUMERIC_COLS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    return df


@st.cache_data(show_spinner="Loading inventory")
def load_inventory(path: str, mtime: float) -> pd.DataFrame:
    df = clean_inventory(pd.read_excel(path, sheet_name=SHEET, header=HEADER_ROW - 1))
    return _fill_formula_values(Path(path), df, formula_columns(path, mtime))


@st.cache_data(show_spinner="Reading uploaded file")
def load_inventory_upload(content: bytes) -> pd.DataFrame:
    return clean_inventory(pd.read_excel(BytesIO(content), sheet_name=SHEET, header=HEADER_ROW - 1))

def _status_group(value) -> str:
    if not isinstance(value, str):
        return "Unknown"
    key = value.strip().upper()
    if key in STATUS_LABELS:
        return STATUS_LABELS[key]
    close = difflib.get_close_matches(key, list(STATUS_LABELS), n=1, cutoff=0.8)
    return STATUS_LABELS[close[0]] if close else value.strip().title()

def add_derived(df: pd.DataFrame) -> pd.DataFrame:
    """Analysis-only columns. These are never written back to Excel."""
    out = df.copy()
    status = out["STATUS"].astype("string").str.strip().str.upper()
    out["status_group"] = status.map(_status_group)
    for label in STATUS_ORDER:
        out[f"is_{label.lower()}"] = out["status_group"] == label
    return out


def get_inventory() -> tuple[pd.DataFrame, Path | None] | None:
    """Server file if present, otherwise an in-memory upload (sidebar). Returns (data, editable_path)."""
    try:
        if DATA_FILE.exists():
            return load_inventory(str(DATA_FILE), DATA_FILE.stat().st_mtime), DATA_FILE
        st.info("Inventory file not found. Upload it in the sidebar to continue (used for this session only).")
        with st.sidebar:
            st.subheader("Data")
            upload = st.file_uploader(f"Inventory workbook (sheet '{SHEET}')", type="xlsx", key="ops_upload")
        return (load_inventory_upload(upload.getvalue()), None) if upload else None
    except ValueError as exc:
        st.error(f"Could not read the '{SHEET}' sheet: {exc}")
        return None

# EXCEL WRITE-BACK 
def _excel_row(index_label) -> int:
    """pandas keeps the original row positions as index labels, so label 0 = first data row."""
    return HEADER_ROW + 1 + int(index_label)


def _to_cell(value):
    v = _clean(value)
    if isinstance(v, str):
        v = v.strip() or None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return v


def _open_sheet(path: Path):
    wb = load_workbook(path)
    if SHEET not in wb.sheetnames:
        raise ValueError(f"Sheet '{SHEET}' was not found.")
    ws = wb[SHEET]
    headers = {str(ws.cell(HEADER_ROW, c).value).strip(): c
               for c in range(1, ws.max_column + 1) if ws.cell(HEADER_ROW, c).value is not None}
    return wb, ws, headers


def _save_workbook(wb, path: Path) -> None:
    backup = path.with_name(f"{path.stem}.backup{path.suffix}")
    shutil.copy2(path, backup)  # one rolling backup, refreshed before every save
    try:
        wb.calculation.fullCalcOnLoad = True  # Excel recalculates every formula when it opens the file
    except Exception:
        pass
    wb.save(path)
    load_inventory.clear()


def _check_columns(values: dict, headers: dict) -> None:
    unknown = [c for c in values if c not in headers]
    if unknown:
        raise ValueError("Column(s) not found in the Excel sheet: " + ", ".join(unknown))

#formulas 
@st.cache_data(show_spinner=False)
def formula_columns(path: str, mtime: float) -> list[str]:
    """Names of the columns that contain formulas (e.g. a 'Total Unit' = area x price)."""
    wb = load_workbook(path, read_only=True, data_only=False)
    try:
        ws = wb[SHEET]
        header_row = next(ws.iter_rows(min_row=HEADER_ROW, max_row=HEADER_ROW, values_only=True), ())
        names = [None if v is None else str(v).strip() for v in header_row]
        found = set()
        for row in ws.iter_rows(min_row=HEADER_ROW + 1, values_only=True):
            for name, value in zip(names, row):
                if name and isinstance(value, str) and value.startswith("="):
                    found.add(name)
        return sorted(found)
    finally:
        wb.close()


_REF = re.compile(r"\b([A-Za-z]{1,3})(\d+)\b")
_BINOPS = {ast.Add: operator.add, ast.Sub: operator.sub, ast.Mult: operator.mul, ast.Div: operator.truediv,
           ast.Pow: operator.pow}


def _xl_round(x, n=0):
    f = 10 ** int(n)
    return math.copysign(math.floor(abs(x) * f + 0.5) / f, x)


_FUNCS = {"ROUND": _xl_round, "ABS": abs, "MIN": min, "MAX": max, "SUM": lambda *a: sum(a)}


def _eval_formula(ws, formula: str, depth: int = 0):
    """Evaluate simple arithmetic formulas such as =H6*I6 or =ROUND(H6*I6,0). Returns None if it can't."""
    if depth > 6:
        return None
    expr = _REF.sub(lambda m: f"_c('{m.group(1).upper()}{m.group(2)}')", formula[1:].replace("$", "").replace("^", "**"))

    def cell(addr):
        v = ws[addr].value
        if v is None or v == "":
            return 0.0
        if isinstance(v, bool):
            raise ValueError
        if isinstance(v, (int, float)):
            return float(v)
        if isinstance(v, str) and v.startswith("="):
            r = _eval_formula(ws, v, depth + 1)
            if r is None:
                raise ValueError
            return r
        raise ValueError

    def ev(node):
        if isinstance(node, ast.Expression):
            return ev(node.body)
        if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)) and not isinstance(node.value, bool):
            return float(node.value)
        if isinstance(node, ast.BinOp) and type(node.op) in _BINOPS:
            return _BINOPS[type(node.op)](ev(node.left), ev(node.right))
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.USub, ast.UAdd)):
            return -ev(node.operand) if isinstance(node.op, ast.USub) else ev(node.operand)
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name):
            if node.func.id == "_c":
                return cell(node.args[0].value)
            func = _FUNCS.get(node.func.id.upper())
            if func:
                return func(*[ev(a) for a in node.args])
        raise ValueError("unsupported")

    try:
        return float(ev(ast.parse(expr.strip(), mode="eval")))
    except (ValueError, SyntaxError, ZeroDivisionError, TypeError, OverflowError, KeyError):
        return None


def _fill_formula_values(path: Path, df: pd.DataFrame, fcols: list[str]) -> pd.DataFrame:
    """A file saved by this app has formulas but no stored results, so pandas reads those cells as empty.
    Recompute the simple ones so totals keep working. Anything it can't evaluate stays empty
    (Excel recalculates everything the next time the file is opened)."""
    fcols = [c for c in fcols if c in df.columns]
    if not fcols or not df[fcols].isna().any().any():
        return df
    try:
        wb = load_workbook(path, data_only=False)
        ws = wb[SHEET]
        headers = {str(ws.cell(HEADER_ROW, c).value).strip(): c
                   for c in range(1, ws.max_column + 1) if ws.cell(HEADER_ROW, c).value is not None}
        out = df.copy()
        for col in fcols:
            for idx in out.index[out[col].isna()]:
                raw = ws.cell(_excel_row(idx), headers[col]).value
                if isinstance(raw, str) and raw.startswith("="):
                    value = _eval_formula(ws, raw)
                    if value is not None:
                        out.at[idx, col] = value
        wb.close()
        return out
    except Exception:  
        return df

#writing 
def save_edits(path: Path, changes: dict, deletes: list) -> None:
    """Write many edits in ONE save: changes = {index_label: {column: value}}, deletes = [index_label]."""
    wb, ws, headers = _open_sheet(path)
    _check_columns({c: None for ch in changes.values() for c in ch}, headers)
    gone = set(deletes)
    for index_label, cells in changes.items():
        if index_label in gone:
            continue
        row = _excel_row(index_label)
        for column, value in cells.items():
            ws.cell(row, headers[column]).value = _to_cell(value)
    for index_label in deletes:  # clears the row
        row = _excel_row(index_label)
        for column in headers.values():
            ws.cell(row, column).value = None
    _save_workbook(wb, path)


def append_unit(path: Path, values: dict) -> None:
    """Add a row under the last used one, copying its formatting and formulas (e.g. totals)."""
    wb, ws, headers = _open_sheet(path)
    _check_columns(values, headers)
    cols = list(headers.values())
    last = next((r for r in range(ws.max_row, HEADER_ROW, -1)
                 if any(ws.cell(r, c).value not in (None, "") for c in cols)), HEADER_ROW)
    new = last + 1
    if last > HEADER_ROW:
        for col in cols:
            src, dst = ws.cell(last, col), ws.cell(new, col)
            if src.has_style:
                dst._style = copy(src._style)
            if isinstance(src.value, str) and src.value.startswith("="):
                try:
                    dst.value = Translator(src.value, origin=src.coordinate).translate_formula(dst.coordinate)
                except Exception:
                    pass
    for column, value in values.items():
        ws.cell(new, headers[column]).value = _to_cell(value)
    _save_workbook(wb, path)

# SIDEBAR FILTERS
def _options(df: pd.DataFrame, col: str) -> list[str]:
    return sorted(df[col].dropna().astype(str).unique().tolist(), key=_nat_key) if col in df.columns else []


def _reset_filters() -> None:
    for key in [k for k in st.session_state if k.startswith(FILTER_PREFIX)]:
        del st.session_state[key]


def render_filters(df: pd.DataFrame) -> dict:
    """Filters live in the sidebar, like the HR page. Empty selection = no restriction."""
    k = FILTER_PREFIX
    phases = _options(df, "Phase")  
    statuses = [s for s in STATUS_ORDER if s in set(df["status_group"])] + \
               sorted(s for s in df["status_group"].unique() if s not in STATUS_ORDER)
    with st.sidebar:
        st.divider()
        st.subheader("Filters")
        f = {
            "phase": st.multiselect("Phase", phases, key=k + "phase"),
            "status": st.multiselect("Status", statuses, key=k + "status"),
            "building": st.multiselect("Building", _options(df, "Building"), key=k + "building"),
            "search": st.text_input("Search unit code or client", key=k + "search").strip().lower(),
        }
        with st.expander("More filters"):
            f["type"] = st.multiselect("Type", _options(df, "Type"), key=k + "type")
            f["unit_type"] = st.multiselect("Unit type", _options(df, "Unit Type"), key=k + "unit_type")
            rooms = df["Rooms Num"].dropna() if "Rooms Num" in df else pd.Series(dtype=float)
            f["rooms"] = st.multiselect("Rooms", sorted(rooms.astype(int).unique().tolist()),
                                        key=k + "rooms") if not rooms.empty else []
            floors = df["Floor"].dropna() if "Floor" in df else pd.Series(dtype=float)
            fb = (int(floors.min()), int(floors.max())) if not floors.empty else None
            f["floor_bounds"] = fb
            f["floor"] = st.slider("Floor", fb[0], fb[1], fb, key=k + "floor") if fb and fb[0] < fb[1] else fb
            area = df["In/Area"].dropna() if "In/Area" in df else pd.Series(dtype=float)
            ab = (float(area.min()), float(area.max())) if not area.empty else None
            f["area_bounds"] = ab
            f["area"] = st.slider("Area (m²)", ab[0], ab[1], ab, step=1.0, key=k + "area") if ab and ab[0] < ab[1] else ab
        st.button("↩️ Reset filters", on_click=_reset_filters)
    return f


def _text_match(series: pd.Series, query: str) -> pd.Series:
    return series.astype("string").fillna("").str.lower().str.contains(query, regex=False)


def apply_filters(df: pd.DataFrame, f: dict) -> pd.DataFrame:
    out = df
    if f["phase"]:
        out = out[out["Phase"].isin(f["phase"])]
    if f["status"]:
        out = out[out["status_group"].isin(f["status"])]
    for col, key in (("Type", "type"), ("Building", "building"), ("Unit Type", "unit_type")):
        if f[key] and col in out:
            out = out[out[col].astype(str).isin(f[key])]
    if f["rooms"]:
        out = out[out["Rooms Num"].isin(f["rooms"])]
    if f["floor"] and f["floor"] != f["floor_bounds"]:
        out = out[out["Floor"].between(*f["floor"])]
    if f["area"] and f["area"] != f["area_bounds"]:
        out = out[out["In/Area"].between(*f["area"])]
    if f["search"]:
        mask = _text_match(out["Unit Code"], f["search"])
        if "Name of client" in out:
            mask |= _text_match(out["Name of client"], f["search"])
        out = out[mask.astype(bool)]
    return out

# KPIs
def calculate_kpis(d: pd.DataFrame) -> dict:
    total = len(d)
    sold, avail = d[d["is_sold"]], d[d["is_available"]]
    has_value = _has(d, "Total Unit")
    return {
        "total": total, "sold": len(sold), "available": len(avail), "hold": int(d["is_hold"].sum()),
        "reserved": int(d["is_reserved"].sum()), "conversion": len(sold) / total if total else 0.0,
    }


def render_kpis(k: dict) -> None:
    c = st.columns(5)
    c[0].metric("Total units", fmt_int(k["total"]))
    c[1].metric("Sold", fmt_int(k["sold"]),
                help="Sold units and the sum of 'Total Unit' for them")
    c[2].metric("Available", fmt_int(k["available"]), help="Available units and the sum of 'Total Unit' for them")
    c[3].metric("Hold / Reserved", f"{fmt_int(k['hold'])} / {fmt_int(k['reserved'])}",
                help="Units in the pipeline that are not sold yet")
    c[4].metric("Conversion", f"{k['conversion']:.1%}", help="Sold units ÷ all units")

# CHARTS (Plotly). 
def _present(d: pd.DataFrame) -> list[str]:
    seen = set(d["status_group"])
    return [s for s in STATUS_ORDER if s in seen] + sorted(s for s in seen if s not in STATUS_ORDER)


def _finish(fig, height: int = 390):
    fig.update_layout(height=height, margin=dict(l=20, r=20, t=55, b=45), hovermode="x unified", legend_title_text="")
    return fig


def _phases(d: pd.DataFrame) -> list[str]:
    """Every phase in the data, so a newly added phase / sub-project gets its own bars automatically."""
    return sorted(d["Phase"].dropna().astype(str).unique().tolist(), key=_nat_key)


def _plot(fig, empty: str = "Not enough data for this chart.") -> None:
    if fig is None:
        st.info(empty)
    else:
        st.plotly_chart(fig, use_container_width=True)


def inventory_status_fig(d: pd.DataFrame):
    """Units per status, one bar per phase."""
    phases = _phases(d)
    if not phases:
        return None
    base = ["Sold", "Available", "Hold", "Reserved"]
    statuses = base + sorted(x for x in set(d["status_group"]) if x not in base)
    counts = d.assign(Phase=d["Phase"].astype(str)).groupby(["Phase", "status_group"]).size()
    g = counts.reindex(pd.MultiIndex.from_product([phases, statuses]), fill_value=0).reset_index()
    g.columns = ["Phase", "Status", "Units"]
    fig = px.bar(g, x="Status", y="Units", color="Phase", barmode="group", text="Units",
                 category_orders={"Status": statuses, "Phase": phases}, labels={"Status": "Status", "Units": "Units"})
    fig.update_traces(textposition="outside", cliponaxis=False, hovertemplate="%{x}<br>Units: %{y}<extra></extra>")
    fig.update_yaxes(rangemode="tozero")
    return _finish(fig)


def conversion_fig(d: pd.DataFrame):
    """Sales conversion (% sold) per phase."""
    phases = _phases(d)
    if not phases:
        return None
    ph = d["Phase"].astype(str)
    rows = []
    for phase in phases:
        x = d[ph == phase]
        rows.append({"Phase": phase, "Conversion": round(x["is_sold"].mean() * 100, 1) if len(x) else 0})
    conv = pd.DataFrame(rows)
    fig = px.bar(conv, x="Phase", y="Conversion", text="Conversion",
                 labels={"Conversion": "Sales Conversion (%)", "Phase": "Phase"})
    fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside", cliponaxis=False,
                      hovertemplate="%{x}<br>Conversion: %{y:.1f}%<extra></extra>")
    fig.update_yaxes(range=[0, max(100, float(conv["Conversion"].max()) + 10)])
    return _finish(fig, 350)


def available_building_fig(d: pd.DataFrame):
    """Available units per building, grouped by phase."""
    if "Building" not in d:
        return None
    a = d[d["is_available"]].dropna(subset=["Building", "Phase"])
    if a.empty:
        return None
    g = a.assign(Building=a["Building"].astype(str), Phase=a["Phase"].astype(str)) \
         .groupby(["Building", "Phase"]).size().reset_index(name="Units")
    fig = px.bar(g, x="Building", y="Units", color="Phase", barmode="group", text="Units",
                 category_orders={"Building": sorted(g["Building"].unique(), key=_nat_key), "Phase": _phases(d)},
                 labels={"Building": "Building", "Units": "Available Units"})
    fig.update_traces(textposition="outside", cliponaxis=False, hovertemplate="%{x}<br>Available: %{y}<extra></extra>")
    fig.update_xaxes(tickangle=-45)
    return _finish(fig, 350)


def unit_type_fig(d: pd.DataFrame):
    """Units per unit type, grouped by phase."""
    if "Unit Type" not in d:
        return None
    s = d.dropna(subset=["Unit Type", "Phase"])
    if s.empty:
        return None
    g = s.assign(**{"Unit Type": s["Unit Type"].astype(str), "Phase": s["Phase"].astype(str)}) \
         .groupby(["Unit Type", "Phase"]).size().reset_index(name="Units")
    fig = px.bar(g, x="Unit Type", y="Units", color="Phase", barmode="group", text="Units",
                 category_orders={"Phase": _phases(d)}, labels={"Unit Type": "Unit Type", "Units": "Units"})
    fig.update_traces(textposition="outside", cliponaxis=False)
    fig.update_xaxes(tickangle=-35)
    return _finish(fig, 390)


def phase_summary(d: pd.DataFrame) -> pd.DataFrame:
    ph = d["Phase"].astype(str)
    rows = []
    for phase in _phases(d):
        x = d[ph == phase]
        rows.append({"Phase": phase, "Total": len(x), "Sold": int(x["is_sold"].sum()),
                     "Available": int(x["is_available"].sum()),
                     "Conversion": round(x["is_sold"].mean() * 100, 1) if len(x) else 0})
    return pd.DataFrame(rows, columns=["Phase", "Total", "Sold", "Available", "Conversion"])


def render_charts(d: pd.DataFrame) -> None:
    st.subheader("Inventory Status")
    _plot(inventory_status_fig(d))

    left, right = st.columns(2)
    with left:
        st.subheader("Sales Conversion")
        _plot(conversion_fig(d))
    with right:
        st.subheader("Available Units by Building")
        _plot(available_building_fig(d), "No available units.")

    left, right = st.columns(2)
    with left:
        st.subheader("Unit Type Mix")
        _plot(unit_type_fig(d))
    with right:
        st.subheader("Phase Summary")
        st.dataframe(phase_summary(d), use_container_width=True, hide_index=True, column_config={
            "Conversion": st.column_config.NumberColumn("Conversion", format="%.1f%%")})


# UNIT DIRECTORY: inline editing + add unit / new phase
def _missing_required(values: dict) -> list[str]:
    return [c for c in REQUIRED_ON_SAVE if c in values and not str(_clean(values[c]) or "").strip()]


def _code_taken(df: pd.DataFrame, code, ignore_index=None) -> bool:
    codes = df["Unit Code"].dropna().astype(str).str.strip().str.lower()
    if ignore_index is not None:
        codes = codes.drop(index=ignore_index, errors="ignore")
    return str(code).strip().lower() in set(codes)


def _excel_bytes(frame: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        frame.to_excel(writer, index=False, sheet_name=SHEET)
    return buffer.getvalue()


def _data_cols(df: pd.DataFrame) -> list[str]:
    """The sheet's own columns (without the analysis-only columns added by add_derived)."""
    return [c for c in df.columns if c != "status_group" and not c.startswith("is_")]


def _ordered_cols(df: pd.DataFrame) -> list[str]:
    cols = _data_cols(df)
    return [c for c in DISPLAY_COLS if c in cols] + [c for c in cols if c not in DISPLAY_COLS]


def diff_edits(edited: pd.DataFrame, df: pd.DataFrame, editable: list[str]) -> tuple[dict, list]:
    """Compare the edited table with the loaded data. Returns ({row: {column: new value}}, [rows to delete])."""
    changes, deletes = {}, []
    for i in edited.index:
        if bool(edited.at[i, DEL]):
            deletes.append(i)
            continue
        cells = {c: edited.at[i, c] for c in editable if not _same(edited.at[i, c], df.at[i, c])}
        if cells:
            changes[i] = cells
    return changes, deletes


def validate_edits(df: pd.DataFrame, changes: dict, deletes: list) -> list[str]:
    problems = []
    for i, cells in changes.items():
        for col in REQUIRED_ON_SAVE:
            if col in cells and not str(_clean(cells[col]) or "").strip():
                problems.append(f"Unit {df.at[i, 'Unit Code']}: {col} cannot be empty.")
    codes = df["Unit Code"].astype("string").str.strip().str.lower().drop(index=deletes, errors="ignore")
    for i, cells in changes.items():
        if "Unit Code" in cells and i in codes.index:
            codes.at[i] = str(_clean(cells["Unit Code"]) or "").strip().lower()
    counts = codes.value_counts()
    for i, cells in changes.items():
        if "Unit Code" in cells and i in codes.index and codes.at[i] and counts.get(codes.at[i], 0) > 1:
            problems.append(f"Unit Code '{cells['Unit Code']}' already exists.")
    return problems


def _choices(df: pd.DataFrame) -> dict[str, list[str]]:
    out = {c: _options(df, c) for c in CHOICE_COLS}
    out["STATUS"] = list(CANONICAL_STATUS)  # fixed list
    return out


def _choice_input(column: str, options: list[str], key: str) -> str:
    """Pick an existing value or create a new one (a new phase / sub-project, building, unit type...)."""
    label = "Phase / sub-project" if column == "Phase" else column
    picked = st.selectbox(label, [""] + options + [NEW_OPTION], key=key)
    if picked == NEW_OPTION:
        return st.text_input(f"New {label.lower()} name", key=f"{key}_new").strip()
    return picked


def _add_form(columns: list[str], choices: dict, key_prefix: str) -> dict:
    values = {}
    half = (len(columns) + 1) // 2
    for ui_col, group in zip(st.columns(2), [columns[:half], columns[half:]]):
        with ui_col:
            for column in group:
                key = f"{key_prefix}_{column}"
                if column in CHOICE_COLS:
                    values[column] = _choice_input(column, choices.get(column, []), key)
                elif column in NUMERIC_COLS:
                    values[column] = st.number_input(column, value=None, key=key)
                else:
                    values[column] = st.text_input(column, key=key)
    return values


def _missing_required(values: dict) -> list[str]:
    return [c for c in REQUIRED_ON_SAVE if c in values and not str(_clean(values[c]) or "").strip()]


def _code_taken(df: pd.DataFrame, code) -> bool:
    return str(code).strip().lower() in set(df["Unit Code"].dropna().astype(str).str.strip().str.lower())


@st.dialog("Add New Unit")
def add_dialog(df: pd.DataFrame, path: Path, fcols: list[str]) -> None:
    st.caption("To create a NEW phase / sub-project (or building, unit type...), choose “➕ Add new…” in its list. "
               "Calculated columns are filled in by the sheet's formulas.")
    columns = [c for c in _ordered_cols(df) if c not in fcols]
    values = _add_form(columns, _choices(df), "ops_add")
    save, cancel = st.columns(2)
    if save.button("Add unit", type="primary", use_container_width=True):
        for col in ("Phase", "STATUS"):  # same upper-case convention as the sheet
            if col in values and values[col]:
                values[col] = values[col].strip().upper()
        missing = _missing_required(values)
        if missing:
            st.error("Please complete: " + ", ".join(missing))
        elif _code_taken(df, values["Unit Code"]):
            st.error("This Unit Code already exists.")
        else:
            try:
                append_unit(path, {c: v for c, v in values.items() if _clean(v) not in (None, "")})
                st.session_state["ops_ver"] = st.session_state.get("ops_ver", 0) + 1
                st.toast("Unit added", icon="✅")
                st.rerun()
            except Exception as exc:
                st.error(f"Could not add the unit: {exc}")
    if cancel.button("Cancel", use_container_width=True):
        st.rerun()


def _excel_bytes(frame: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        frame.to_excel(writer, index=False, sheet_name=SHEET)
    return buffer.getvalue()


def _column_config(df: pd.DataFrame, cols: list[str]) -> dict:
    config = {DEL: st.column_config.CheckboxColumn(DEL, help="Tick to delete this unit when you press Save", default=False)}
    statuses = CANONICAL_STATUS + sorted(v for v in df["STATUS"].dropna().astype(str).unique() if v not in CANONICAL_STATUS)
    for c in cols:
        if c == "STATUS":
            config[c] = st.column_config.SelectboxColumn(c, options=statuses)
        elif c in NUMERIC_COLS:
            config[c] = st.column_config.NumberColumn(c, format="%.1f" if c == "In/Area" else "%.0f")
    return config


def render_unit_table(filtered: pd.DataFrame, df: pd.DataFrame, edit_path: Path | None, filter_sig: str,
                      fcols: list[str]) -> None:
    st.subheader("Unit directory")
    cols = _ordered_cols(df)
    ver = st.session_state.setdefault("ops_ver", 0)

    if edit_path is None:
        st.caption(f"{len(filtered):,} of {len(df):,} units")
        st.dataframe(filtered[cols], use_container_width=True, hide_index=True, height=460)
        st.warning(f"✏️ Editing is off: the workbook was not found on the server, so this is an uploaded copy. "
                   f"Put the file at {DATA_FILE} to enable editing.")
        st.download_button("Export", _excel_bytes(filtered[cols]), icon=":material/download:",
                           file_name="Gardenia_Town_Filtered.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        return

    st.caption(f"{len(filtered):,} of {len(df):,} units · ✏️ Double-click a cell to edit · 🗑️ Tick Delete on a row · "
               "💾 Save writes to the Excel file (a backup is made first) · ➕ Add unit for new units or a new phase.")
    if fcols:
        st.caption("🔒 Calculated by Excel formulas (read-only): " + ", ".join(fcols))

    shown = filtered[cols].copy()
    shown.insert(0, DEL, False)
    editable = [c for c in cols if c not in fcols]
    edited = st.data_editor(shown, hide_index=True, height=460, key=f"ops_dir_{ver}_{abs(hash(filter_sig))}",
                            disabled=list(fcols), column_config=_column_config(df, cols))

    changes, deletes = diff_edits(edited, df, editable)
    problems = validate_edits(df, changes, deletes)
    for problem in problems:
        st.error(problem)

    b_save, b_add, b_export, _ = st.columns([1.6, 1, 1, 1.4])
    if b_save.button(f"💾 Save changes ({len(changes)} edited · {len(deletes)} to delete)", type="primary",
                     disabled=not (changes or deletes) or bool(problems), use_container_width=True):
        try:
            save_edits(edit_path, changes, deletes)
            st.session_state["ops_ver"] = ver + 1
            st.toast("Saved to Excel", icon="✅")
            st.rerun()
        except Exception as exc:
            st.error(f"Could not save: {exc}")
    if b_add.button("Add unit", icon=":material/add:", use_container_width=True):
        add_dialog(df, edit_path, fcols)
    b_export.download_button("Export", _excel_bytes(filtered[cols]), icon=":material/download:",
                             file_name="Gardenia_Town_Filtered.xlsx",
                             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                             use_container_width=True)

# PAGE
def main() -> None:
    render_header(DEPARTMENT)

    loaded = get_inventory()
    if loaded is None:
        return
    raw, edit_path = loaded
    missing = [c for c in REQUIRED_COLS if c not in raw.columns]
    if missing:
        st.error("Missing columns: " + ", ".join(missing))
        return
    df = add_derived(raw)

    fcols = formula_columns(str(edit_path), edit_path.stat().st_mtime) if edit_path else []

    f = render_filters(df)
    filtered = apply_filters(df, f)
    if filtered.empty:
        st.info("No units match your filters.") 
    else:
        render_kpis(calculate_kpis(filtered))
        render_charts(filtered)
    render_unit_table(filtered, df, edit_path, repr(sorted(f.items())), fcols)

main() 