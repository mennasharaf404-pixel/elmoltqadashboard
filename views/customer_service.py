from __future__ import annotations
import hashlib
import io
import os
import re
from pathlib import Path
import altair as alt
import numpy as np
import pandas as pd
import streamlit as st
from theme import NAVY, NAVY_LIGHT, render_header

# CONFIGURATION
DATA_DIR = Path(os.environ.get("CUSTOMER_SERVICE_DATA_DIR", Path(__file__).resolve().parent.parent / "data"))

COMBINED_FILE = DATA_DIR / "Customer_Service_All.xlsx"
COMBINED_SHEET = 0
PROJECT_SOURCES = {
    "Gardenia 3": {
        "contracts": DATA_DIR / "G3-DB-Customers.xlsx",
        "installments": DATA_DIR / "dataexcel.xlsx",
    },
}
G3_PROJECT = "Gardenia 3"
TOWN_PROJECT = "Gardenia Town"
TOWN_GLOBS = ("gtown_customers*.xlsx", "g_town-installments-delayed*.xlsx")
CURRENCY = "EGP"
ENDING_SOON_DAYS = 90
ARABIC_DIGITS = False
DEPARTMENT = "Customer Service"
DEL = "🗑️ Delete"

# CONSTANTS
REQUIRED_COLUMNS = ["project", "client_name", "building_no", "unit_no", "area_sqm",
                    "contract_date", "handover_date", "remaining_installments", "notes"]
OPTIONAL_COLUMNS = ["floor", "unit_value", "delivery_term", "contract_notes"]
CONTRACT_COLS = {
    "أسم العميل": "client_name", "اسم العميل": "client_name",
    "تاريخ التعاقد": "contract_date", "رقم الوحدة": "unit_no", "رقم العمارة": "building_no",
    "الدور": "floor", "المساحة": "area_sqm", "قيمة الوحدة": "unit_value",
    "تاريخ الإستلام": "handover_date", "مهلة التسليم": "delivery_term", "ملاحظات": "contract_notes",
}
INSTALLMENT_COLS = {"العمارة": "building_no", "الوحدة": "unit_no",
                    "عدد الأقساط المتبقية": "remaining_installments", "ملاحظات": "notes"}
PROJECT_ALIASES = {
    "gardenia 3": "Gardenia 3", "gardenia3": "Gardenia 3", "g3": "Gardenia 3", "جاردينيا 3": "Gardenia 3",
    "gardenia town": "Gardenia Town", "gardenia-town": "Gardenia Town", "town": "Gardenia Town",
    "جاردينيا تاون": "Gardenia Town",
}
FLOOR_FIXES = {"الارضي": "الأرضي", "الاول": "الأول", "الاولى": "الأولى"}

STATUS_ORDER = ["منتظم", "ملتزم", "إلى حد ما منتظم", "غير منتظم", "غير ملتزم", "متأخر", "أخرى"]
STATUS_COLORS = {
    "منتظم": "#27a85a", "ملتزم": "#17a398", "إلى حد ما منتظم": "#ee9716",
    "غير منتظم": "#d94343", "غير ملتزم": "#9b2c2c", "متأخر": "#7141c7", "أخرى": "#9aa1ae",
}
RISK_STATUSES = ("غير منتظم", "غير ملتزم", "متأخر")
BUCKET_ORDER = ["0 (fully paid)", "1–6", "7–12", "13–24", "More than 24"]
DANGER = "#C0563B"

TOWN_COLS = {
    "الاسم": "client_name", "المشروع": "sub_project", "العمارة": "building_no", "الوحدة": "unit_no",
    "المتأخرات": "arrears", "الشهور": "delay_period", "عدد الشهور": "months_delayed", "ملاحظات": "notes",
}
ARABIC_MONTHS = {"يناير": 1, "فبراير": 2, "مارس": 3, "ابريل": 4, "مايو": 5, "يونيو": 6,
                 "يوليو": 7, "اغسطس": 8, "سبتمبر": 9, "اكتوبر": 10, "نوفمبر": 11, "ديسمبر": 12}
FOLLOWUP_RULES = [
    ("Resale / termination", ("انذار فسخ", "ريسيل", "resale")),
    ("Contact problem", ("الرقم غلط", "no whatsapp")),
    ("Promised to pay", ("اكسبشن", "هيتم السداد", "هننزل الشيك", "بيدفع بشيك")),  # before "paid"
    ("Marked paid", ("paid", "تم السداد")),
    ("Message sent", ("whatsapp",)),
]
NO_NOTE = "No follow-up note"
FOLLOWUP_ORDER = ["Resale / termination", "Contact problem", "Promised to pay", "Message sent",
                  "Marked paid", NO_NOTE, "Other"]
AMOUNT_BANDS = [(0, 50_000, "Up to 50K"), (50_000, 100_000, "50K–100K"), (100_000, 150_000, "100K–150K"),
                (150_000, 250_000, "150K–250K"), (250_000, float("inf"), "Over 250K")]
MONTH_BUCKETS = ["Not recorded (0)", "1 month", "2 months", "3+ months"]

# =====================================================================
# HELPERS
# =====================================================================
def parse_number(value) -> float:
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


def fmt_money(x) -> str:
    if pd.isna(x):
        return "—"
    return f"{x / 1e6:,.2f}M" if abs(x) >= 1e6 else f"{x / 1e3:,.0f}K" if abs(x) >= 1e3 else f"{x:,.0f}"


def _key_part(value) -> str:
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return ""
    text = re.sub(r"\s+", " ", str(value)).strip()
    return text.split(".")[0] if re.fullmatch(r"\d+\.0+", text) else text


def _ending_soon_mask(dates: pd.Series, today: pd.Timestamp) -> pd.Series:
    return dates.between(today, today + pd.Timedelta(days=ENDING_SOON_DAYS))


def _sorted_buildings(values) -> list[str]:
    return sorted({v for v in values if v}, key=lambda v: (0, int(v)) if v.isdigit() else (1, v))


# =====================================================================
# DATA LOADING + CLEANING - GARDENIA 3
# =====================================================================
def _source_name(source) -> str:
    return getattr(source, "name", None) or "uploaded file"


def _available(source) -> bool:
    return source is not None and (hasattr(source, "read") or Path(source).exists())


def _read_table(path, anchor: str) -> pd.DataFrame:
    """Read a sheet whose header row is not necessarily the first row (finds the row containing `anchor`)."""
    raw = pd.read_excel(path, header=None, dtype=object)
    header_idx = next((i for i, row in raw.iterrows() if anchor in [str(v).strip() for v in row.values]), None)
    if header_idx is None:
        raise ValueError(f"Header column '{anchor}' not found in {_source_name(path)}")
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

    if _available(cfg.get("installments")):
        inst = _select_renamed(_read_table(cfg["installments"], "عدد الأقساط المتبقية"), INSTALLMENT_COLS)
        inst = inst.dropna(subset=["building_no", "unit_no"], how="all")
        inst["_key"] = inst["building_no"].map(_key_part) + "|" + inst["unit_no"].map(_key_part)
        dup = int(inst["_key"].duplicated().sum())
        inst_unique = inst.drop_duplicates("_key", keep="first")
        merged = contracts.merge(inst_unique[["_key", "remaining_installments", "notes"]], on="_key", how="left")
        unmatched = int(merged["remaining_installments"].isna().sum())
        orphan = len(set(inst_unique["_key"]) - set(contracts["_key"]))
        notes.append(f"{project}: {len(contracts)} contracts; {unmatched} without matching installment data; "
                     f"{orphan} installment units without a contract; {dup} duplicate units in the installment file "
                     f"(first record kept).")
        contracts = merged
    else:
        contracts["remaining_installments"] = pd.NA
        contracts["notes"] = pd.NA
        notes.append(f"{project}: no installment file — installments and status are unavailable.")

    contracts["project"] = project
    return contracts.drop(columns="_key"), notes


def _load_combined(path: Path) -> tuple[pd.DataFrame, list[str]]:
    df = pd.read_excel(path, sheet_name=COMBINED_SHEET)
    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"{path.name} is missing required columns: {', '.join(missing)}")
    return df, [f"Source: {path.name}"]


def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """Meaning-preserving cleaning (whitespace, spelling variants, types). Safe to run again after edits."""
    df = df.copy()
    for col in OPTIONAL_COLUMNS:
        if col not in df.columns:
            df[col] = pd.NA
    for col in ["project", "client_name", "building_no", "unit_no", "floor", "notes", "contract_notes"]:
        df[col] = df[col].map(_key_part)
    df["project"] = df["project"].map(lambda p: PROJECT_ALIASES.get(p.lower(), p))
    df["floor"] = df["floor"].replace(FLOOR_FIXES)
    for col in ["contract_date", "handover_date"]:
        df[col] = pd.to_datetime(df[col], errors="coerce")
    df["area_sqm"] = df["area_sqm"].map(parse_number)
    df["unit_value"] = pd.to_numeric(df["unit_value"], errors="coerce")
    df["remaining_installments"] = df["remaining_installments"].map(parse_number)
    df["status_category"] = df["notes"].map(get_status)
    df["unit_label"] = "عمارة " + df["building_no"] + " – وحدة " + df["unit_no"]
    df["search_text"] = (df["client_name"] + " " + df["notes"] + " " + df["contract_notes"]).str.lower()
    return df


def _source_signature() -> tuple:
    paths = [COMBINED_FILE] + [Path(p) for cfg in PROJECT_SOURCES.values() for p in cfg.values()]
    return tuple((str(p), p.stat().st_mtime if p.exists() else None) for p in paths)


@st.cache_data(show_spinner="Loading Gardenia 3 data...")
def load_data(signature: tuple) -> tuple[pd.DataFrame, list[str]]:
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
            raise FileNotFoundError("No Gardenia 3 data found")
        df = pd.concat(frames, ignore_index=True)
    return clean_data(df), notes


@st.cache_data(show_spinner="Reading uploaded files...")
def load_g3_uploaded(contracts: bytes, installments: bytes | None) -> tuple[pd.DataFrame, list[str]]:
    cfg = {"contracts": io.BytesIO(contracts), "installments": io.BytesIO(installments) if installments else None}
    frame, notes = _load_project_sources(G3_PROJECT, cfg)
    return clean_data(frame), notes + ["Data came from files uploaded in this session (nothing is stored on the server)."]


# =====================================================================
# DATA LOADING + CLEANING - GARDENIA TOWN
# =====================================================================
def _norm_ar(text: str) -> str:
    return text.translate(str.maketrans("أإآ", "ااا")).lower()


def _parse_delay_period(label: str) -> tuple[list[int], float]:
    """'15/اكتوبر' -> ([10], 15); 'مايو - اغسطس' -> ([5, 8], nan). The file has no year."""
    text = _norm_ar(label)
    months = sorted({num for name, num in ARABIC_MONTHS.items() if name in text})
    day = re.search(r"(\d{1,2})\s*/", text)
    return months, float(day.group(1)) if day else float("nan")


def classify_followup(note) -> str:
    text = _norm_ar(note) if isinstance(note, str) else ""
    if not text.strip():
        return NO_NOTE
    for status, keywords in FOLLOWUP_RULES:
        if any(k in text for k in keywords):
            return status
    return "Other"


def _months_bucket(n: float) -> str:
    if pd.isna(n) or n <= 0:
        return MONTH_BUCKETS[0]
    return MONTH_BUCKETS[1] if n == 1 else MONTH_BUCKETS[2] if n == 2 else MONTH_BUCKETS[3]


def _amount_band(x: float) -> str | None:
    if pd.isna(x):
        return None
    return next(label for lo, hi, label in AMOUNT_BANDS if lo <= x <= hi) if x > 0 else AMOUNT_BANDS[0][2]


def clean_town_data(df: pd.DataFrame) -> pd.DataFrame:
    """Whitespace/type cleaning only. Safe to run again after edits."""
    df = df.copy()
    for col in ["client_name", "sub_project", "building_no", "unit_no", "notes", "delay_period"]:
        df[col] = df[col].map(_key_part)
    df["arrears"] = df["arrears"].map(parse_number)
    df["months_delayed"] = df["months_delayed"].map(parse_number)
    parsed = df["delay_period"].map(_parse_delay_period)
    df["oldest_month"] = parsed.map(lambda p: p[0][0] if p[0] else float("nan"))
    df["due_day"] = parsed.map(lambda p: p[1])
    df["followup_status"] = df["notes"].map(classify_followup)
    df["months_bucket"] = df["months_delayed"].map(_months_bucket)
    df["amount_band"] = df["arrears"].map(_amount_band)
    df["unit_label"] = df["sub_project"] + " · Bldg " + df["building_no"] + " · Unit " + df["unit_no"]
    df["search_text"] = (df["client_name"] + " " + df["notes"]).str.lower()
    df["project"] = TOWN_PROJECT
    return df


def _find_town_file() -> Path | None:
    files = sorted({f for pattern in TOWN_GLOBS for f in DATA_DIR.glob(pattern)},
                   key=lambda f: f.stat().st_mtime, reverse=True)
    return files[0] if files else None


def _town_signature() -> tuple:
    path = _find_town_file()
    return (str(path), path.stat().st_mtime) if path else (None, None)


def _build_town(source, filename: str) -> tuple[pd.DataFrame, list[str]]:
    table = _select_renamed(_read_table(source, "المتأخرات"), TOWN_COLS)
    missing = [c for c in TOWN_COLS.values() if c not in table.columns]
    if missing:
        raise ValueError(f"{filename} is missing columns: {', '.join(missing)}")
    table = table.dropna(subset=["client_name", "arrears"], how="all").reset_index(drop=True)
    return clean_town_data(table), []


@st.cache_data(show_spinner="Loading Gardenia Town data...")
def load_town_data(signature: tuple) -> tuple[pd.DataFrame, list[str]]:
    path = _find_town_file()
    if path is None:
        raise FileNotFoundError("No Gardenia Town data found")
    return _build_town(path, path.name)


@st.cache_data(show_spinner="Reading uploaded file...")
def load_town_uploaded(content: bytes, filename: str) -> tuple[pd.DataFrame, list[str]]:
    df, notes = _build_town(io.BytesIO(content), filename)
    return df, notes + ["Data came from a file uploaded in this session (nothing is stored on the server)."]


# =====================================================================
# SESSION STATE (one editable copy of the data per project)
# =====================================================================
def _prepare(prefix: str, key: tuple, loader) -> str | None:
    """(Re)load a project's data into session state when its source changed. Returns an error text or None."""
    if st.session_state.get(f"{prefix}_key") == key and f"{prefix}_data" in st.session_state:
        return None
    try:
        data, notes = loader()
    except (FileNotFoundError, ValueError) as exc:
        st.session_state.pop(f"{prefix}_data", None)
        st.session_state.pop(f"{prefix}_key", None)
        return str(exc)
    st.session_state[f"{prefix}_data"] = data.copy()
    st.session_state[f"{prefix}_notes"] = notes
    st.session_state[f"{prefix}_ver"] = 0
    st.session_state[f"{prefix}_key"] = key
    return None


# =====================================================================
# KPIs
# =====================================================================
def calculate_kpis(df: pd.DataFrame, today: pd.Timestamp) -> dict:
    total = len(df)
    at_risk = int(df["status_category"].isin(RISK_STATUSES).sum())
    paid_off = int((df["remaining_installments"] == 0).sum())
    return {
        "total": total, "area": df["area_sqm"].sum(),
        "completion": paid_off / total * 100 if total else 0.0,
        "at_risk": at_risk, "at_risk_pct": at_risk / total * 100 if total else 0.0,
        "ending_soon": int(_ending_soon_mask(df["handover_date"], today).sum()),
    }


def calculate_town_kpis(df: pd.DataFrame) -> dict:
    total = df["arrears"].sum()
    return {
        "accounts": len(df), "total": total, "median": df["arrears"].median(),
        "multi": int((df["months_delayed"] >= 2).sum()),
        "top10_share": df["arrears"].nlargest(10).sum() / total * 100 if total else 0.0,
    }


# =====================================================================
# DIRECTORY (search / edit / delete / add / save / download) - shared by both projects
# =====================================================================
def _num(v):
    return np.nan if v is None else float(v)


def g3_add_form(df: pd.DataFrame, today: pd.Timestamp):
    with st.form("g3_add", clear_on_submit=True):
        name = st.text_input("Client name *")
        building = st.text_input("Building *")
        unit = st.text_input("Unit *")
        floor = st.text_input("Floor")
        area = st.number_input("Area (m²)", min_value=0.0, value=None, step=1.0)
        value = st.number_input("Unit value", min_value=0.0, value=None, step=1000.0)
        remaining = st.number_input("Remaining installments", min_value=0, value=None, step=1)
        contract = st.date_input("Contract date", value=None)
        handover = st.date_input("Handover date", value=None)
        notes = st.text_area("Notes")
        submitted = st.form_submit_button("Add contract", type="primary")
    if not submitted:
        return None
    if not (name.strip() and building.strip() and unit.strip()):
        st.error("Client name, building and unit are required.")
        return None
    return {"project": G3_PROJECT, "client_name": name, "building_no": building, "unit_no": unit, "floor": floor,
            "area_sqm": _num(area), "unit_value": _num(value), "remaining_installments": _num(remaining),
            "contract_date": pd.Timestamp(contract) if contract else pd.NaT,
            "handover_date": pd.Timestamp(handover) if handover else pd.NaT, "notes": notes}


def town_add_form(df: pd.DataFrame, today: pd.Timestamp):
    with st.form("gt_add", clear_on_submit=True):
        name = st.text_input("Client name *")
        sub = st.selectbox("Sub-project", [""] + sorted(df["sub_project"].unique()))
        building = st.text_input("Building")
        unit = st.text_input("Unit")
        arrears = st.number_input("Arrears amount *", min_value=0.0, value=None, step=1000.0)
        months = st.number_input("Months delayed", min_value=0, value=None, step=1)
        period = st.text_input("Delay period (e.g. 15/اكتوبر)")
        notes = st.text_area("Notes")
        submitted = st.form_submit_button("Add account", type="primary")
    if not submitted:
        return None
    if not name.strip() or not arrears:
        st.error("Client name and a positive arrears amount are required.")
        return None
    return {"project": TOWN_PROJECT, "client_name": name, "sub_project": sub, "building_no": building,
            "unit_no": unit, "arrears": float(arrears), "months_delayed": _num(months),
            "delay_period": period, "notes": notes}


G3_TABLE = dict(
    prefix="g3", title="Contract directory", search_label="Search client name or notes",
    add_label="➕ Add contract", clean=clean_data, add_form=g3_add_form,
    cols=["client_name", "building_no", "unit_no", "floor", "area_sqm", "unit_value", "remaining_installments",
          "contract_date", "handover_date", "status_category", "notes"],
    labels={"client_name": "Client", "building_no": "Building", "unit_no": "Unit", "floor": "Floor",
            "area_sqm": "Area (m²)", "unit_value": "Unit value", "remaining_installments": "Remaining installments",
            "contract_date": "Contract date", "handover_date": "Handover date", "status_category": "Status",
            "notes": "Notes"},
    date_cols=["contract_date", "handover_date"],
    num_cols=["area_sqm", "unit_value", "remaining_installments"],
    num_fmt={"area_sqm": "%.1f"},
    disabled=["status_category"],  # derived from the notes
    sort=(["building_no", "unit_no"], True),
    scope=lambda d: d[d["project"] == G3_PROJECT],
    file="customer_service_gardenia3.xlsx", sheet="Contracts",
)
TOWN_TABLE = dict(
    prefix="gt", title="Delayed accounts directory", search_label="Search client name or notes",
    add_label="➕ Add account", clean=clean_town_data, add_form=town_add_form,
    cols=["client_name", "sub_project", "building_no", "unit_no", "arrears", "months_delayed", "delay_period",
          "followup_status", "notes"],
    labels={"client_name": "Client", "sub_project": "Sub-project", "building_no": "Building", "unit_no": "Unit",
            "arrears": "Arrears", "months_delayed": "Months delayed", "delay_period": "Delay period",
            "followup_status": "Follow-up status", "notes": "Notes"},
    date_cols=[], num_cols=["arrears", "months_delayed"], num_fmt={},
    disabled=["followup_status"],  # derived from the notes
    sort=(["arrears"], False),
    scope=lambda d: d,
    file="customer_service_gardenia_town.xlsx", sheet="Delayed accounts",
)


def directory(cfg: dict, view: pd.DataFrame, filter_sig, today: pd.Timestamp) -> None:
    p, cols = cfg["prefix"], cfg["cols"]
    data_key, ver_key = f"{p}_data", f"{p}_ver"

    st.subheader(cfg["title"])
    st.caption("✏️ Double-click to edit · 🗑️ Tick a row to delete · ➕ Add new records · 💾 Save to update charts. "
               "📥 Download before you leave.")
    q = st.text_input(cfg["search_label"], key=f"{p}_q").strip().lower()
    tbl = view
    if q:
        tbl = view[view["search_text"].str.contains(q, regex=False, na=False)]
    tbl = tbl.sort_values(cfg["sort"][0], ascending=cfg["sort"][1])

    shown = tbl[cols].copy()
    shown.insert(0, DEL, False)
    config = {DEL: st.column_config.CheckboxColumn(DEL, help="Tick to delete this record", default=False)}
    for c in cols:
        label = cfg["labels"][c]
        if c in cfg["date_cols"]:
            config[c] = st.column_config.DateColumn(label, format="YYYY-MM-DD")
        elif c in cfg["num_cols"]:
            config[c] = st.column_config.NumberColumn(label, format=cfg["num_fmt"].get(c, "%.0f"))
        else:
            config[c] = st.column_config.TextColumn(label)

    sig = abs(hash((q, filter_sig)))
    edited = st.data_editor(shown, hide_index=True, key=f"{p}_dir_{st.session_state[ver_key]}_{sig}",
                            disabled=cfg["disabled"], column_config=config)

    n_del = int(edited[DEL].sum())
    b1, b2, b3 = st.columns([1, 1, 1])
    if b1.button(f"💾 Save changes ({n_del} to delete)", type="primary", key=f"{p}_save"):
        base = st.session_state[data_key].copy()
        for c in cols:
            if c in cfg["disabled"]:
                continue
            if c in cfg["date_cols"]:
                base.loc[edited.index, c] = pd.to_datetime(edited[c], errors="coerce")
            elif c in cfg["num_cols"]:
                base.loc[edited.index, c] = pd.to_numeric(edited[c], errors="coerce")
            else:
                base.loc[edited.index, c] = edited[c]
        base = base.drop(index=edited.index[edited[DEL]])
        st.session_state[data_key] = cfg["clean"](base)  # recompute status / buckets from the edited values
        st.session_state[ver_key] += 1
        st.rerun()

    with b2.popover(cfg["add_label"]):
        row = cfg["add_form"](st.session_state[data_key], today)
    if row is not None:
        base = st.session_state[data_key]
        new = {c: None for c in base.columns}
        new.update(row)
        new_idx = int(base.index.max()) + 1 if len(base) else 0
        st.session_state[data_key] = cfg["clean"](pd.concat([base, pd.DataFrame([new], index=[new_idx])]))
        st.session_state[ver_key] += 1
        st.rerun()

    export = cfg["scope"](st.session_state[data_key])[cols].copy().rename(columns=cfg["labels"])
    for c in cfg["date_cols"]:
        export[cfg["labels"][c]] = export[cfg["labels"][c]].dt.date
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        export.to_excel(xw, index=False, sheet_name=cfg["sheet"])
    b3.download_button("⬇️ Download Excel", buf.getvalue(), file_name=cfg["file"],
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"{p}_dl")


# =====================================================================
# CHARTS (4 per project)
# =====================================================================
FOLLOWUP_COLORS = {
    "Resale / termination": "#6b2d5c", "Contact problem": "#7141c7", "Promised to pay": "#ee9716",
    "Message sent": "#17a398", "Marked paid": "#27a85a", NO_NOTE: "#d94343", "Other": "#9aa1ae",
}
SEVERITY_COLORS = ["#27a85a", "#9bd17a", "#ee9716", "#d94343", "#9b2c2c"]


def _t(title: str, subtitle: str) -> alt.TitleParams:
    """Chart heading: a plain-language question/title plus a one-line explanation of what is shown."""
    return alt.TitleParams(text=title, subtitle=subtitle, anchor="start", fontSize=15, subtitleFontSize=11,
                           subtitleColor="gray")


def show(chart):
    if chart is not None:
        st.altair_chart(chart)


def _pct_text(counts: pd.Series) -> pd.Series:
    return counts.astype(int).astype(str) + " · " + (counts / counts.sum() * 100).round(0).astype(int).astype(str) + "%"


def _hbar(d: pd.DataFrame, cat: str, val: str, text: str, title: str, tooltip: list,
          color: str = NAVY, color_enc=None, order=None, xfmt: str | None = None):
    """Horizontal bars with a text label at the end of each bar."""
    axis = alt.Axis(format=xfmt) if xfmt else alt.Axis()
    base = alt.Chart(d).encode(
        y=alt.Y(f"{cat}:N", sort=order or "-x", title=None, axis=alt.Axis(labelLimit=280)),
        x=alt.X(f"{val}:Q", title=None, axis=axis),
    )
    bars = base.mark_bar() if color_enc is not None else base.mark_bar(color=color)
    bars = bars.encode(color=color_enc, tooltip=tooltip) if color_enc is not None else bars.encode(tooltip=tooltip)
    labels = base.mark_text(align="left", dx=3, fontSize=10).encode(text=f"{text}:N")
    return (bars + labels).properties(title=title, height=max(160, 32 * len(d)), width="container")


def _vbar(d: pd.DataFrame, cat: str, val: str, text: str, title: str, tooltip: list, order: list,
          color: str = NAVY, color_enc=None, angle: int = 0):
    base = alt.Chart(d).encode(x=alt.X(f"{cat}:N", sort=order, title=None, axis=alt.Axis(labelAngle=angle)),
                               y=alt.Y(f"{val}:Q", title=None))
    bars = base.mark_bar() if color_enc is not None else base.mark_bar(color=color)
    bars = bars.encode(color=color_enc, tooltip=tooltip) if color_enc is not None else bars.encode(tooltip=tooltip)
    labels = base.mark_text(dy=-7, fontSize=10).encode(text=f"{text}:N")
    return (bars + labels).properties(title=title, height=300, width="container")


# ---- Gardenia 3 -------------------------------------------------------
def g3_status_chart(view: pd.DataFrame):
    """How reliably are customers paying?"""
    c = view["status_category"].value_counts()
    present = [s for s in STATUS_ORDER if s in c.index]
    if not present:
        return None
    counts = pd.Series([c[s] for s in present], index=present)
    d = pd.DataFrame({"status": present, "contracts": counts.values, "txt": _pct_text(counts).values})
    color = alt.Color("status:N", legend=None,
                      scale=alt.Scale(domain=present, range=[STATUS_COLORS[s] for s in present]))
    return _hbar(d, "status", "contracts", "txt", _t("Payment Status", "Contracts by how customers pay"),
                 ["status:N", "contracts:Q"], color_enc=color, order=present)


def g3_risk_chart(view: pd.DataFrame):
    """Which buildings have a bigger payment-risk problem than the portfolio average?"""
    d = view[view["building_no"] != ""].copy()
    if d.empty:
        return None
    d["risk"] = d["status_category"].isin(RISK_STATUSES)
    g = d.groupby("building_no").agg(total=("risk", "size"), at_risk=("risk", "sum")).reset_index()
    if g["at_risk"].sum() == 0:
        return None
    overall = float(view["status_category"].isin(RISK_STATUSES).mean())
    g["share"] = g["at_risk"] / g["total"]
    g = g.sort_values(["share", "total"], ascending=False).head(15)
    g["building"] = "Building " + g["building_no"]
    g["txt"] = g["at_risk"].astype(int).astype(str) + " of " + g["total"].astype(str)
    base = alt.Chart(g).encode(y=alt.Y("building:N", sort=g["building"].tolist(), title=None),
                               x=alt.X("share:Q", title=None, axis=alt.Axis(format="%")))
    bars = base.mark_bar().encode(
        color=alt.condition(f"datum.share > {overall}", alt.value(DANGER), alt.value(NAVY_LIGHT)),
        tooltip=["building:N", alt.Tooltip("share:Q", title="At-risk share", format=".0%"), "txt:N"])
    labels = base.mark_text(align="left", dx=3, fontSize=10).encode(text="txt:N")
    rule = alt.Chart(pd.DataFrame({"avg": [overall]})).mark_rule(strokeDash=[5, 4], color="gray", size=2).encode(x="avg:Q")
    return (bars + labels + rule).properties(
        title=_t("At-Risk Contracts by Building", f"Red = above the {overall:.0%} average"),
        height=max(160, 32 * len(g)), width="container")


def _bucket(n: float) -> str | None:
    if pd.isna(n):
        return None
    if n == 0:
        return "0 (fully paid)"
    return "1–6" if n <= 6 else "7–12" if n <= 12 else "13–24" if n <= 24 else "More than 24"


def g3_installments_chart(view: pd.DataFrame):
    b = view["remaining_installments"].map(_bucket).dropna()
    if b.empty:
        return None
    c = b.value_counts().reindex(BUCKET_ORDER, fill_value=0)
    d = pd.DataFrame({"bucket": BUCKET_ORDER, "contracts": c.values, "txt": _pct_text(c).values})
    color = alt.Color("bucket:N", legend=None, scale=alt.Scale(domain=BUCKET_ORDER, range=SEVERITY_COLORS))
    return _vbar(d, "bucket", "contracts", "txt", _t("Installments Left", "Contracts by remaining installments"),
                 ["bucket:N", "contracts:Q"], BUCKET_ORDER, color_enc=color)


def g3_handover_chart(view: pd.DataFrame, today: pd.Timestamp):
    """What is being handed over each quarter, and how much of it is still unpaid?"""
    d = view.dropna(subset=["handover_date"]).copy()
    if d.empty:
        return None
    d["q"] = d["handover_date"].dt.to_period("Q")
    rem = d["remaining_installments"]
    d["group"] = np.where(rem.isna(), "No installment data", np.where(rem > 0, "Still paying", "Fully paid"))
    qs = pd.period_range(d["q"].min(), d["q"].max(), freq="Q")
    labels = [f"{p.year} Q{p.quarter}" for p in qs]
    g = d.groupby(["q", "group"]).size().reset_index(name="contracts")
    g["quarter"] = g["q"].map(lambda p: f"{p.year} Q{p.quarter}")
    g = g.drop(columns="q")  # Period dtype cannot be sent to the chart
    colors = {"Fully paid": "#27a85a", "Still paying": "#ee9716", "No installment data": "#9aa1ae"}
    present = [x for x in colors if x in set(g["group"])]
    bars = alt.Chart(g).mark_bar().encode(
        x=alt.X("quarter:N", sort=labels, scale=alt.Scale(domain=labels), title=None, axis=alt.Axis(labelAngle=-45)),
        y=alt.Y("contracts:Q", title=None),
        color=alt.Color("group:N", title=None, legend=alt.Legend(orient="top"),
                        scale=alt.Scale(domain=present, range=[colors[x] for x in present])),
        tooltip=["quarter:N", "group:N", "contracts:Q"])
    chart = bars
    cur = today.to_period("Q")
    if cur in qs:
        rule = alt.Chart(pd.DataFrame({"quarter": [f"{cur.year} Q{cur.quarter}"]})).mark_rule(
            strokeDash=[5, 4], color="gray", size=2).encode(x=alt.X("quarter:N", sort=labels))
        chart = bars + rule
    return chart.properties(title=_t("Handovers by Quarter", "Dashed line = current quarter"),
                            height=300, width="container")


# ---- Gardenia Town ----------------------------------------------------
def town_buildings_chart(view: pd.DataFrame):
    g = view.groupby(["sub_project", "building_no"]).agg(amt=("arrears", "sum"), n=("arrears", "size"))
    g = g.sort_values("amt", ascending=False).head(10).reset_index()
    if g.empty:
        return None
    d = pd.DataFrame({"building": g["sub_project"] + " · " + g["building_no"], "sub": g["sub_project"],
                      "amt": g["amt"], "accounts": g["n"],
                      "txt": g["amt"].map(fmt_money) + " · " + g["n"].astype(str) + " accts"})
    color = alt.Color("sub:N", title=None, legend=alt.Legend(orient="top"))
    return _hbar(d, "building", "amt", "txt", _t("Top 10 Buildings", "Buildings with the most arrears"),
                 ["building:N", alt.Tooltip("amt:Q", title="Arrears", format=",.0f"), "accounts:Q"],
                 color_enc=color, order=d["building"].tolist())


def town_followup_chart(view: pd.DataFrame):
    """How much money sits in each follow-up stage? Red = nobody has followed up yet."""
    g = view.groupby("followup_status").agg(amt=("arrears", "sum"), n=("arrears", "size"))
    order = [s for s in FOLLOWUP_ORDER if s in g.index]
    if not order:
        return None
    g = g.loc[order].reset_index()
    d = pd.DataFrame({"status": g["followup_status"], "amt": g["amt"], "accounts": g["n"],
                      "txt": g["amt"].map(fmt_money) + " · " + g["n"].astype(str) + " accts"})
    color = alt.Color("status:N", legend=None,
                      scale=alt.Scale(domain=order, range=[FOLLOWUP_COLORS[s] for s in order]))
    return _hbar(d, "status", "amt", "txt", _t("Follow-up Status", "Arrears by follow-up stage · red = not followed up"),
                 ["status:N", alt.Tooltip("amt:Q", title="Arrears", format=",.0f"), "accounts:Q"],
                 color_enc=color, order=order)


def town_months_chart(view: pd.DataFrame):
    """The longer an account is late, the harder it is to collect."""
    g = view.groupby("months_bucket").agg(amt=("arrears", "sum"), n=("arrears", "size"))
    g = g.reindex(MONTH_BUCKETS, fill_value=0).reset_index()
    total = g["amt"].sum()
    if total <= 0:
        return None
    d = pd.DataFrame({"months": g["months_bucket"], "amt": g["amt"], "accounts": g["n"],
                      "txt": g["amt"].map(fmt_money) + " · " + (g["amt"] / total * 100).round(0).astype(int).astype(str)
                             + "% · " + g["n"].astype(str)})
    color = alt.Color("months:N", legend=None,
                      scale=alt.Scale(domain=MONTH_BUCKETS, range=["#9aa1ae", "#ee9716", "#d94343", "#9b2c2c"]))
    return _vbar(d, "months", "amt", "txt", _t("Months Delayed", "Arrears by how long accounts are late"),
                 ["months:N", alt.Tooltip("amt:Q", title="Arrears", format=",.0f"), "accounts:Q"],
                 MONTH_BUCKETS, color_enc=color)


def town_top_accounts_chart(view: pd.DataFrame):
    """Biggest balances, colored by follow-up stage - big red bars are the ones nobody has chased."""
    t = view.nlargest(10, "arrears").reset_index(drop=True)
    if t.empty:
        return None
    months = t["months_delayed"].map(lambda m: "—" if pd.isna(m) else f"{int(m)}m")
    d = pd.DataFrame({
        "account": (t.index + 1).astype(str) + ". " + t["client_name"] + " · " + t["building_no"] + "/" + t["unit_no"],
        "amt": t["arrears"], "status": t["followup_status"],
        "txt": t["arrears"].map(fmt_money) + " · " + months})
    present = [s for s in FOLLOWUP_ORDER if s in set(d["status"])]
    color = alt.Color("status:N", title=None, legend=alt.Legend(orient="top"),
                      scale=alt.Scale(domain=present, range=[FOLLOWUP_COLORS[s] for s in present]))
    return _hbar(d, "account", "amt", "txt", _t("Top 10 Accounts", "Largest balances · color = follow-up stage"),
                 ["account:N", "status:N", alt.Tooltip("amt:Q", title="Arrears", format=",.0f")],
                 color_enc=color, order=d["account"].tolist())


# =====================================================================
# SIDEBAR (project selector, then that project's data + filters)
# =====================================================================
PROJECTS = [G3_PROJECT, TOWN_PROJECT]


def sidebar_project() -> str:
    with st.sidebar:
        st.subheader("Project")
        return st.radio("Project", PROJECTS, key="cs_project", label_visibility="collapsed")


def sidebar_data(project: str) -> dict:
    with st.sidebar:
        st.divider()
        st.subheader("Data")
        if project == G3_PROJECT:
            contracts = st.file_uploader("Contracts register (.xlsx)", type=["xlsx"], key="up_g3_contracts")
            installments = st.file_uploader("Installments follow-up (.xlsx, optional)", type=["xlsx"],
                                            key="up_g3_installments")
            return {"contracts": contracts, "installments": installments}
        return {"town": st.file_uploader("Delayed installments (.xlsx)", type=["xlsx"], key="up_town")}


def g3_filters(df: pd.DataFrame) -> dict:
    scope = df[df["project"] == G3_PROJECT]
    with st.sidebar:
        st.divider()
        st.subheader("Filters")
        statuses = st.multiselect("Status", [s for s in STATUS_ORDER if s in set(scope["status_category"])],
                                  key="g3_status")
        buildings = st.multiselect("Building", _sorted_buildings(scope["building_no"]), key="g3_building")
        floors = st.multiselect("Floor", sorted({f for f in scope["floor"] if f}), key="g3_floor")
        years = sorted(scope["handover_date"].dropna().dt.year.unique())
        year = st.selectbox("Handover year", ["All"] + [int(y) for y in years], key="g3_year")
        soon = st.checkbox(f"Handover within {ENDING_SOON_DAYS} days only", key="g3_soon")
        if st.button("↩️ Reset all edits", key="g3_reset"):
            st.session_state["g3_key"] = None
            st.rerun()
    return {"statuses": statuses, "buildings": buildings, "floors": floors, "year": year, "soon": soon}


def filter_g3(df: pd.DataFrame, f: dict, today: pd.Timestamp) -> pd.DataFrame:
    out = df[df["project"] == G3_PROJECT]
    if f["statuses"]:
        out = out[out["status_category"].isin(f["statuses"])]
    if f["buildings"]:
        out = out[out["building_no"].isin(f["buildings"])]
    if f["floors"]:
        out = out[out["floor"].isin(f["floors"])]
    if f["year"] != "All":
        out = out[out["handover_date"].dt.year == f["year"]]
    if f["soon"]:
        out = out[_ending_soon_mask(out["handover_date"], today)]
    return out


def town_filters(df: pd.DataFrame) -> dict:
    with st.sidebar:
        st.divider()
        st.subheader("Filters")
        sub = st.multiselect("Sub-project", sorted(df["sub_project"].unique()), key="gt_sub")
        bld = st.multiselect("Building", _sorted_buildings(df["building_no"]), key="gt_bld")
        months = st.multiselect("Months delayed", [b for b in MONTH_BUCKETS if b in set(df["months_bucket"])],
                                key="gt_months")
        status = st.multiselect("Follow-up status", [x for x in FOLLOWUP_ORDER if x in set(df["followup_status"])],
                                key="gt_status")
        lo, hi = float(df["arrears"].min()), float(df["arrears"].max())
        amount = (st.slider("Arrears amount", lo, hi, (lo, hi), step=1000.0, key=f"gt_amount_{int(lo)}_{int(hi)}")
                  if lo < hi else (lo, hi))
        if st.button("↩️ Reset all edits", key="gt_reset"):
            st.session_state["gt_key"] = None
            st.rerun()
    return {"sub": sub, "bld": bld, "months": months, "status": status, "amount": amount, "bounds": (lo, hi)}


def filter_town(df: pd.DataFrame, f: dict) -> pd.DataFrame:
    out = df
    if f["sub"]:
        out = out[out["sub_project"].isin(f["sub"])]
    if f["bld"]:
        out = out[out["building_no"].isin(f["bld"])]
    if f["months"]:
        out = out[out["months_bucket"].isin(f["months"])]
    if f["status"]:
        out = out[out["followup_status"].isin(f["status"])]
    if f["amount"] != f["bounds"]:
        out = out[out["arrears"].between(*f["amount"])]
    return out


# =====================================================================
# PROJECT PAGES
# =====================================================================
def render_g3(error, filters, today) -> None:
    if error:
        st.warning(f"No Gardenia 3 data found. Upload the files in the sidebar. ({error})")
        return
    view = filter_g3(st.session_state["g3_data"], filters, today)
    if view.empty:
        st.info("No contracts match the current filters.")
        return

    k = calculate_kpis(view, today)
    cols = st.columns(5)
    cols[0].metric("Contracts", fmt_int(k["total"]), help="All registered contracts")
    cols[1].metric("Total area (m²)", fmt_int(k["area"]), help="Combined unit area")
    cols[2].metric("Fully paid", f"{k['completion']:.0f}%", help="Contracts with no remaining installments")
    cols[3].metric("At-risk contracts", f"{fmt_int(k['at_risk'])} ({k['at_risk_pct']:.0f}%)",
                   help="Status recorded as irregular, non-committed or late")
    cols[4].metric(f"Handover in {ENDING_SOON_DAYS} days", fmt_int(k["ending_soon"]),
                   help="Contracts with an upcoming handover date")

    c1, c2 = st.columns(2)
    with c1:
        show(g3_status_chart(view))
    with c2:
        show(g3_risk_chart(view))
    c3, c4 = st.columns(2)
    with c3:
        show(g3_installments_chart(view))
    with c4:
        show(g3_handover_chart(view, today))

    directory(G3_TABLE, view, repr(sorted(filters.items())), today)


def render_town(error, filters, today) -> None:
    if error:
        st.warning(f"No Gardenia Town data found. Upload the file in the sidebar. ({error})")
        return
    view = filter_town(st.session_state["gt_data"], filters)
    if view.empty:
        st.info("No accounts match the current filters.")
        return

    k = calculate_town_kpis(view)
    cols = st.columns(5)
    cols[0].metric("Delayed accounts", fmt_int(k["accounts"]), help="Accounts with overdue installments")
    cols[1].metric(f"Total arrears ({CURRENCY})", fmt_money(k["total"]), help="Sum of all overdue amounts")
    cols[2].metric("Median arrears", fmt_money(k["median"]), help="Typical account: half owe less, half owe more")
    cols[3].metric("Delayed 2+ months", fmt_int(k["multi"]), help="Accounts that missed more than one month")
    cols[4].metric("Top 10 share", f"{k['top10_share']:.0f}%",
                   help="Share of arrears held by the 10 largest balances")

    c1, c2 = st.columns(2)
    with c1:
        show(town_followup_chart(view))
    with c2:
        show(town_months_chart(view))
    c3, c4 = st.columns(2)
    with c3:
        show(town_top_accounts_chart(view))
    with c4:
        show(town_buildings_chart(view))

    directory(TOWN_TABLE, view, repr(sorted(filters.items())), today)


# =====================================================================
# PAGE
# =====================================================================
def main() -> None:
    render_header(DEPARTMENT)
    today = pd.Timestamp.today().normalize()
    project = sidebar_project()      # 1. choose the project
    up = sidebar_data(project)       # 2. that project's data
    st.markdown(f"#### {project}")

    if project == G3_PROJECT:
        if up["contracts"] is not None:
            c = up["contracts"].getvalue()
            i = up["installments"].getvalue() if up["installments"] else None
            key = ("up", hashlib.md5(c).hexdigest(), hashlib.md5(i).hexdigest() if i else None)
            err = _prepare("g3", key, lambda: load_g3_uploaded(c, i))
        else:
            sig = _source_signature()
            err = _prepare("g3", ("srv", sig), lambda: load_data(sig))
        filters = None if err else g3_filters(st.session_state["g3_data"])   # 3. that project's filters
        render_g3(err, filters, today)
    else:
        if up["town"] is not None:
            t = up["town"].getvalue()
            err = _prepare("gt", ("up", hashlib.md5(t).hexdigest()), lambda: load_town_uploaded(t, up["town"].name))
        else:
            tsig = _town_signature()
            err = _prepare("gt", ("srv", tsig), lambda: load_town_data(tsig))
        filters = None if err else town_filters(st.session_state["gt_data"])
        render_town(err, filters, today)


main()