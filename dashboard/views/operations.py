# -*- coding: utf-8 -*-
"""
EL MOLTQA — Operations & Sales Dashboard
Reads the existing Google Sheet without changing the source.
The workbook contains monthly tabs with three separate tables:
الحجوزات / التعاقدات / الالغاءات.

This version is intentionally built around the ACTUAL worksheet structure:
- section/banner row
- one main header row beginning with "م"
- optional second header row for unit details
- records until the next section

NOTE: this file is loaded as a page inside the multipage app (see app.py),
so it does NOT call st.set_page_config() itself — that's set once, globally,
in app.py.
"""
from __future__ import annotations

import re
import json
import uuid
from io import BytesIO
from pathlib import Path
from urllib.parse import quote

import numpy as np
import pandas as pd
import requests
import streamlit as st

from theme import (
    inject_css,
    render_header,
    sidebar_brand,
    NAVY,
    NAVY_DARK,
    NAVY_LIGHT,
    ACCENT_GOLD,
    ACCENT_TEAL,
    BG,
    CARD_BG,
    TEXT_MUTED,
)

# =========================================================
# CONFIG
# =========================================================
SHEET_ID = "1lHloVtHag6yZs02q8XLcTu-XOF5Pn87rAQNOStkPmc4"
SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit"

MONTH_TAB_MAP = {
    "يناير": "JAN",
    "فبراير": "FEB",
    "مارس": "MAR",
    "أبريل": "April",
    "مايو": "May",
    "يونيو": "june",
    "يوليو": "July",
    "أغسطس": "Aug",
    "سبتمبر": "SEP",
}
BASE_MONTH_ORDER = [
    "يناير", "فبراير", "مارس", "أبريل", "مايو",
    "يونيو", "يوليو", "أغسطس", "سبتمبر",
]
MONTHS = BASE_MONTH_ORDER.copy()
CATEGORIES = ["الحجوزات", "التعاقدات", "الالغاءات"]
ICON = {"الحجوزات": "📥", "التعاقدات": "📝", "الالغاءات": "↩️"}


def ordered_months(values):
    """Return months in calendar order; custom months come after the standard year."""
    values = [clean(v) for v in values if clean(v)]
    standard = [m for m in BASE_MONTH_ORDER if m in values]
    custom = [m for m in values if m not in BASE_MONTH_ORDER]
    return standard + custom

# =========================================================
# STYLE — pulled from theme.py so this page matches the rest of the app
# (same navy/gold/teal palette, same sidebar, same card look).
# =========================================================
inject_css()  # shared base styling: sidebar gradient, metric cards, app background

# Category accent colors reuse the shared palette instead of one-off hex codes:
CAT_COLOR = {
    "الحجوزات": ACCENT_TEAL,   # bookings
    "التعاقدات": ACCENT_GOLD,  # contracts
    "الالغاءات": "#C0563B",    # cancellations — same rust used for HR elsewhere
}

st.markdown(
    f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;500;600;700;800&display=swap');
    html, body, [class*="css"] {{ font-family: "Cairo", sans-serif; }}

    /* Page-specific pieces not already covered by the shared theme */
    .section-title {{ font-size:1.22rem; font-weight:800; color:{NAVY}; margin:1.1rem 0 .65rem; }}
    .kpi {{
      background:{CARD_BG}; border:1px solid #E4E8EA; border-left:5px solid {NAVY};
      border-radius:14px; padding:18px 20px; min-height:118px;
      box-shadow:0 1px 3px rgba(20,40,50,0.06);
    }}
    .kpi-label {{ font-size:.83rem; color:{TEXT_MUTED}; font-weight:700; }}
    .kpi-value {{ font-size:1.72rem; color:{NAVY}; font-weight:800; margin-top:5px; }}
    .kpi-note {{ font-size:.72rem; color:{TEXT_MUTED}; margin-top:3px; }}
    .card {{ background:{CARD_BG}; border:1px solid #E4E8EA; border-radius:14px; padding:18px; box-shadow:0 1px 3px rgba(20,40,50,0.06); }}
    .cat-booking {{ border-right:5px solid {CAT_COLOR['الحجوزات']}; }}
    .cat-contract {{ border-right:5px solid {CAT_COLOR['التعاقدات']}; }}
    .cat-cancel {{ border-right:5px solid {CAT_COLOR['الالغاءات']}; }}
    .small-tag {{ display:inline-block; padding:5px 10px; border-radius:999px; background:{BG}; color:{NAVY}; font-size:.75rem; font-weight:700; }}
    .notice {{ background:#fff8e6; border:1px solid #f1dfad; color:#755d19; border-radius:14px; padding:12px 14px; }}
    button[kind="primary"] {{ border-radius:10px; background:{NAVY} !important; border-color:{NAVY} !important; }}
    </style>
    """,
    unsafe_allow_html=True,
)

# =========================================================
# TEXT / STRUCTURE HELPERS
# =========================================================
def clean(v) -> str:
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return ""
    return str(v).strip()


def norm(v) -> str:
    s = clean(v)
    s = re.sub(r"[\u064B-\u065F\u0670\u0640\u200f\u200e\s]+", "", s)
    s = s.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    s = s.replace("ى", "ي").lower()
    return s


def section_from_row(row) -> str | None:
    """A section banner is short and mostly empty; data notes may contain the same words."""
    vals = [clean(v) for v in row.tolist() if clean(v)]
    if not vals or len(vals) > 4:
        return None
    text = norm(" ".join(vals))
    if "حجوز" in text or text in {"حجز", "الحجز"}:
        return "الحجوزات"
    if "تعاقد" in text or text in {"عقد", "عقود", "العقد"}:
        return "التعاقدات"
    if "الغاء" in text or "الغاءات" in text:
        return "الالغاءات"
    return None


def find_section_rows(raw: pd.DataFrame):
    found = []
    for i in range(len(raw)):
        kind = section_from_row(raw.iloc[i])
        if kind:
            found.append((i, kind))
    # keep the real section sequence; ignore duplicate accidental matches close together
    result = []
    last_row = -100
    for row, kind in sorted(found, key=lambda x: x[0]):
        if row - last_row >= 3:
            result.append((row, kind))
            last_row = row
    return result


def looks_like_header(row) -> bool:
    vals = [norm(v) for v in row.tolist()]
    nonempty = [v for v in vals if v]
    if len(nonempty) < 4:
        return False
    has_serial = any(v == "م" for v in nonempty)
    has_client = any("اسمالعميل" in v or "اسم" == v for v in nonempty)
    has_code = any("كودالعميل" in v or "كود" == v for v in nonempty)
    has_unit = any("بياناتالوحده" in v or "رقمالوحده" in v or "رقمالعمارة" in v for v in nonempty)
    return (has_serial and (has_client or has_code) and has_unit)


def find_header_row(raw, section_row, end_row):
    for r in range(section_row + 1, min(end_row, section_row + 10)):
        if looks_like_header(raw.iloc[r]):
            return r
    return None


def is_subheader(row) -> bool:
    # A genuine subheader row never carries a serial number in the first
    # column (م) — only real data rows do. Requiring that column to be
    # empty stops a real data row from ever being misread as a subheader,
    # which previously could swallow a record and shift every row after
    # it by one, corrupting the whole table.
    first_cell_empty = norm(clean(row.iloc[0])) == ""
    if not first_cell_empty:
        return False
    vals = [norm(v) for v in row.tolist() if norm(v)]
    markers = ("رقمالعمارة", "رقمالوحده", "المساحه", "الحديقه", "اسمالمرحله", "المشروع")
    return sum(any(m in v for m in markers) for v in vals) >= 2


def build_columns(raw, header_row, subheader_row):
    width = raw.shape[1]
    top = [clean(x) for x in raw.iloc[header_row, :width]]
    sub = [clean(x) for x in raw.iloc[subheader_row, :width]] if subheader_row is not None else [""] * width
    names = []
    for i in range(width):
        # Subheader is the actual name for unit-detail columns; otherwise use the main header.
        name = sub[i] or top[i] or f"عمود {i+1}"
        names.append(name)
    seen = {}
    out = []
    for n in names:
        seen[n] = seen.get(n, 0) + 1
        out.append(n if seen[n] == 1 else f"{n} ({seen[n]})")
    return out


def extract_section(raw, section_row, next_section_row):
    end = next_section_row if next_section_row is not None else len(raw)
    h = find_header_row(raw, section_row, end)
    if h is None:
        return pd.DataFrame(), {"section_row": section_row + 1, "header_row": None, "subheader_row": None}

    sub = h + 1 if h + 1 < end and is_subheader(raw.iloc[h + 1]) else None
    cols = build_columns(raw, h, sub)
    start = sub + 1 if sub is not None else h + 1
    data = raw.iloc[start:end, :len(cols)].copy()
    data.columns = cols

    # Valid source records have a populated serial in column A. This removes blank/total rows
    # without requiring the serial to be numeric.
    first = data.iloc[:, 0].map(clean) if not data.empty else pd.Series(dtype=str)
    nfirst = first.map(norm)
    keep = first.ne("") & ~nfirst.isin({"م", "total", "totals"})
    keep &= ~nfirst.str.contains("اجمال|مجموع", regex=True, na=False)
    data = data.loc[keep].copy()

    for c in data.columns:
        data[c] = data[c].map(clean)

    return data.reset_index(drop=True), {
        "section_row": section_row + 1,
        "header_row": h + 1,
        "subheader_row": sub + 1 if sub is not None else None,
    }


def parse_month(raw):
    result = {c: pd.DataFrame() for c in CATEGORIES}
    diagnostics = {c: [] for c in CATEGORIES}
    sections = find_section_rows(raw)
    for i, (section_row, category) in enumerate(sections):
        # Every matching section is kept and appended, instead of only the
        # first — a month's tab that lists a category more than once
        # (e.g. a second batch of records) no longer has the later records
        # silently dropped.
        nxt = sections[i + 1][0] if i + 1 < len(sections) else None
        frame, diag = extract_section(raw, section_row, nxt)
        diagnostics[category].append(diag)
        if not frame.empty:
            result[category] = (
                pd.concat([result[category], frame], ignore_index=True)
                if not result[category].empty else frame
            )
    return result, diagnostics, sections

# =========================================================
# GOOGLE SHEETS LOADING
# =========================================================

def xlsx_export_url(sheet_id):
    return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"


def csv_export_url(sheet_id, tab):
    return f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={quote(tab)}"


def html_export_url(sheet_id, tab):
    return f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:html&sheet={quote(tab)}"


@st.cache_data(ttl=300, show_spinner=False)
def load_tab(tab_name):
    """Load the original worksheet structure as faithfully as possible.

    IMPORTANT: Never silently prefer CSV over XLSX if an Excel export is
    available. The monthly source uses stacked tables and merged/two-row
    headers; CSV can trim trailing rows and alter the layout.
    """
    errors = []

    # 1) Google XLSX export — the source of truth for stacked worksheet layout.
    try:
        session = requests.Session()
        session.headers.update({
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/140 Safari/537.36",
            "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream;q=0.9,*/*;q=0.8",
        })
        resp = session.get(xlsx_export_url(SHEET_ID), timeout=45, allow_redirects=True)
        resp.raise_for_status()
        content = resp.content
        if not content.startswith(b"PK"):
            raise ValueError(f"Google XLSX export returned {resp.status_code} {resp.headers.get('content-type','')}")
        book = pd.ExcelFile(BytesIO(content), engine="openpyxl")
        actual = next((x for x in book.sheet_names if x == tab_name), None)
        if actual is None:
            actual = next((x for x in book.sheet_names if x.strip().lower() == tab_name.strip().lower()), None)
        if actual is None:
            raise ValueError(f"Tab not found: {tab_name}. Available: {', '.join(book.sheet_names)}")
        df = pd.read_excel(book, sheet_name=actual, header=None, dtype=str, keep_default_na=False)
        df = df.fillna("")
        if df.empty:
            raise ValueError(f"Tab {tab_name} is empty")
        return df, "google-xlsx", None
    except Exception as e:
        errors.append(f"Google XLSX: {type(e).__name__}: {e}")

    # 2) Google HTML export — better structure than CSV when XLSX is blocked.
    try:
        resp = requests.get(
            html_export_url(SHEET_ID, tab_name),
            headers={"User-Agent": "Mozilla/5.0"},
            timeout=35,
        )
        resp.raise_for_status()
        tables = pd.read_html(resp.text)
        if not tables:
            raise ValueError("No HTML table returned")
        # gviz may return helper tables; choose the widest/longest one.
        df0 = max(tables, key=lambda x: (len(x) * max(1, x.shape[1])))
        df = df0.astype(str).replace("nan", "").fillna("")
        if df.empty:
            raise ValueError("HTML table is empty")
        return df, "google-html", None
    except Exception as e:
        errors.append(f"Google HTML: {type(e).__name__}: {e}")

    # 3) Google CSV fallback — only as a last resort.
    try:
        resp = requests.get(
            csv_export_url(SHEET_ID, tab_name),
            headers={"User-Agent": "Mozilla/5.0"},
            timeout=35,
        )
        resp.raise_for_status()
        if not resp.content:
            raise ValueError("Empty CSV response")
        df = pd.read_csv(BytesIO(resp.content), header=None, dtype=str, keep_default_na=False)
        df = df.fillna("")
        if df.empty:
            raise ValueError("CSV table is empty")
        return df, "google-csv", None
    except Exception as e:
        errors.append(f"Google CSV: {type(e).__name__}: {e}")

    return None, None, " | ".join(errors)


@st.cache_data(ttl=300, show_spinner=False)
def load_all_months(cache_version="2026-09-15-v1"):
    """Read every Google Sheet month, then parse the three tables."""
    all_rows = []
    diagnostics = {}
    raw_sizes = {}
    failures = []

    for month, tab in MONTH_TAB_MAP.items():
        raw, source, err = load_tab(tab)
        if raw is None:
            failures.append(f"{month} ({tab})")
            diagnostics[month] = {"tab": tab, "source": None, "error": err}
            continue

        raw_sizes[month] = (len(raw), raw.shape[1], source)
        parsed, diag, sections = parse_month(raw)
        diagnostics[month] = {
            "tab": tab,
            "source": source,
            "diagnostics": diag,
            "sections": [(r + 1, k) for r, k in sections],
        }

        for category, frame in parsed.items():
            if frame.empty:
                continue
            x = frame.copy()
            x.insert(0, "الشهر", month)
            x.insert(1, "النوع", category)
            x.insert(len(x.columns), "__record_id__", [
                f"src:{month}:{category}:{i}" for i in range(len(x))
            ])
            all_rows.append(x)

    union = pd.concat(all_rows, ignore_index=True, sort=False).fillna("") if all_rows else pd.DataFrame()
    return union, diagnostics, raw_sizes, failures

# =========================================================
# DATA DISPLAY HELPERS
# =========================================================
def numeric_value(series):
    s = (
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("،", "", regex=False)
        .str.replace("٬", "", regex=False)
        .str.replace(" ", "", regex=False)
    )
    return pd.to_numeric(s, errors="coerce")


def find_col(df, names):
    for c in df.columns:
        nc = norm(c)
        for n in names:
            if norm(n) in nc:
                return c
    return None


def category_value_col(df, category):
    """Pick the financially meaningful column for each section from the real source layout."""
    if category == "الحجوزات":
        return find_col(df, [
            "مبلع الحجز", "مبلغ الحجز", "قيمه الحجز", "قيمة الحجز"
        ])
    if category == "التعاقدات":
        return find_col(df, [
            "اجمالي الوحده", "اجمالي الوحدة", "اجمالي"
        ])
    if category == "الالغاءات":
        return find_col(df, [
            "مبلع الحجز", "مبلغ الحجز", "قيمه الحجز", "قيمة الحجز"
        ])
    return None


def money_col(df):
    # Backward-compatible generic helper for non-category-specific views.
    return find_col(df, ["اجمالي الوحده", "اجمالي الوحدة", "مبلغ المقدم", "مبلع الحجز", "اجمالي"])


def area_col(df):
    return find_col(df, ["المساحه", "المساحة"])


def prepare_view(df):
    """Readable copy for UI only; source dataframe is not changed."""
    out = df.copy()
    date_words = ["تاريخ"]
    for c in out.columns:
        if any(w in clean(c) for w in date_words):
            # Handle Excel serial dates while preserving ordinary text dates.
            nums = pd.to_numeric(out[c], errors="coerce")
            converted = pd.Series(out[c], index=out.index, dtype="object")
            mask = nums.between(30000, 60000)
            if mask.any():
                dt = pd.to_datetime(nums, unit="D", origin="1899-12-30", errors="coerce")
                converted.loc[mask] = dt.loc[mask].dt.strftime("%Y-%m-%d")
            out[c] = converted
    return out


def reorder_columns(df):
    priority = [
        "الشهر", "النوع", "م", "كود العميل", "كود العميل  / رقم التليفون",
        "اسم العميل", "تيم ليدر", "رقم العمارة", "رقم الوحده", "المساحه",
        "الحديقه", "اسم المرحلة", "المشروع", "مصدر العميل", "اسم شركه البروكر",
        "بروكر", "اجمالي الوحده", "تاريخ الحجز", "تاريخ التعاقد", "تاريخ الالغاء",
        "تاريخ نزول العميل", "المقدم", "مبلع الحجز", " مبلغ المقدم ", "طريقة الدفع",
        "طريقه الدفع", "سعر المتر", "نسبه المقدم", "اسم البائع", "ملحوظات", "سبب الاسترداد",
    ]
    front = []
    for p in priority:
        for c in df.columns:
            if c == p and c not in front:
                front.append(c)
    rest = [c for c in df.columns if c not in front]
    return df[front + rest]

# =========================================================
# LOCAL EDIT / ADD / DELETE LAYER
# =========================================================
EDIT_FILE = Path("el_moltqa_changes.json")

def load_changes():
    try:
        if EDIT_FILE.exists():
            data = json.loads(EDIT_FILE.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                return data
    except Exception:
        pass
    return {"added": [], "deleted": [], "updated": {}, "months": []}

def save_changes(changes):
    EDIT_FILE.write_text(json.dumps(changes, ensure_ascii=False, indent=2), encoding="utf-8")

def row_key(row):
    # Prefer an immutable internal ID so edits remain editable/deletable even
    # after the visible fields change.
    rid = clean(row.get("__record_id__", ""))
    if rid:
        return rid
    parts = [clean(row.get("الشهر","")), clean(row.get("النوع","")), clean(row.get("م","")), clean(row.get("كود العميل","")), clean(row.get("اسم العميل","")), clean(row.get("رقم الوحده",""))]
    return "|".join(parts)

def apply_changes(base):
    changes = load_changes()
    df = base.copy()
    if df.empty:
        return df, changes
    df["__key__"] = df.apply(row_key, axis=1)
    deleted = set(changes.get("deleted", []))
    if deleted:
        df = df[~df["__key__"].isin(deleted)].copy()
    updates = changes.get("updated", {}) or {}
    if updates:
        for key, vals in updates.items():
            mask = df["__key__"].eq(key)
            for col, val in vals.items():
                if col in df.columns:
                    df.loc[mask, col] = val
    added = pd.DataFrame(changes.get("added", []))
    if not added.empty:
        for c in df.columns:
            if c not in added.columns:
                added[c] = ""
        for c in added.columns:
            if c not in df.columns and c != "__key__":
                df[c] = ""
        if "__record_id__" not in added.columns:
            added["__record_id__"] = [f"add:{uuid.uuid4().hex}" for _ in range(len(added))]
        else:
            added["__record_id__"] = added["__record_id__"].map(clean)
            added.loc[added["__record_id__"].eq(""), "__record_id__"] = [f"add:{uuid.uuid4().hex}" for _ in range(int(added["__record_id__"].eq("").sum()))]
        added["__key__"] = added.apply(row_key, axis=1)
        if "__record_id__" not in df.columns:
            df["__record_id__"] = ""
        df = pd.concat([df, added[df.columns]], ignore_index=True, sort=False)
    return df, changes

def register_updated(changes, original_row, new_values):
    key = row_key(original_row)
    changes.setdefault("updated", {})[key] = new_values
    save_changes(changes)

def register_deleted(changes, row):
    key = row_key(row)
    changes.setdefault("deleted", []).append(key)
    changes["deleted"] = list(dict.fromkeys(changes["deleted"]))
    save_changes(changes)

def register_added(changes, row):
    row = dict(row)
    row.setdefault("__record_id__", f"add:{uuid.uuid4().hex}")
    changes.setdefault("added", []).append(row)
    save_changes(changes)

def reset_local_changes():
    save_changes({"added": [], "deleted": [], "updated": {}, "months": []})

def load_custom_months():
    changes = load_changes()
    custom = changes.get("months", []) or []
    for item in custom:
        name = clean(item.get("name", "")) if isinstance(item, dict) else ""
        tab = clean(item.get("tab", "")) if isinstance(item, dict) else ""
        if name and name not in MONTH_TAB_MAP:
            MONTH_TAB_MAP[name] = tab or name
            if name not in MONTHS:
                MONTHS.append(name)

def next_serial(df):
    if "م" not in df.columns or df.empty:
        return 1
    nums = pd.to_numeric(df["م"], errors="coerce")
    return int(nums.max()) + 1 if nums.notna().any() else len(df) + 1

def compact_columns(df):
    preferred = ["الشهر","النوع","م","كود العميل","اسم العميل","تيم ليدر","رقم العمارة","رقم الوحده","المساحه","اسم المرحلة","المشروع","مصدر العميل","تاريخ الحجز","تاريخ التعاقد","تاريخ الالغاء","اجمالي الوحده","المقدم","مبلع الحجز","طريقة الدفع","اسم البائع","ملحوظات"]
    cols = [c for c in preferred if c in df.columns]
    cols += [c for c in df.columns if c not in cols and c != "__key__"]
    return cols

def add_management_ui(all_data, changes):
    st.markdown("<div class='section-title'>إدارة البيانات</div>", unsafe_allow_html=True)
    tabs = st.tabs(["➕ إضافة شهر", "➕ إضافة سجل", "✏️ تعديل", "🗑️ حذف"])

    with tabs[0]:
        with st.form("add_month_form", clear_on_submit=True):
            new_month = st.text_input("اسم الشهر / الفترة", placeholder="مثال: أكتوبر")
            new_tab = st.text_input("اسم الـ Tab في Google Sheet (اختياري)", placeholder="مثال: Oct")
            create = st.form_submit_button("إضافة الشهر", type="primary", use_container_width=True)
        if create and new_month.strip():
            if new_month.strip() in MONTHS:
                st.warning("هذا الشهر موجود بالفعل.")
            else:
                MONTHS.append(new_month.strip())
                MONTHS[:] = ordered_months(MONTHS)
                MONTH_TAB_MAP[new_month.strip()] = new_tab.strip() or new_month.strip()
                changes.setdefault("months", []).append({"name": new_month.strip(), "tab": new_tab.strip() or new_month.strip()})
                save_changes(changes)
                st.success(f"تمت إضافة {new_month.strip()} داخل التطبيق. يمكنك بعد ذلك إضافة سجلات له.")
                st.rerun()
        st.caption("الشهر المضاف هنا يضاف كطبقة داخل التطبيق، ولا يغيّر Google Sheet تلقائيًا.")

    with tabs[1]:
        month_options = MONTHS
        cat_options = CATEGORIES
        with st.form("add_record_form", clear_on_submit=True):
            a1, a2 = st.columns(2)
            with a1:
                month = st.selectbox("الشهر", month_options, key="add_month")
            with a2:
                cat = st.selectbox(
                    "نوع السجل",
                    cat_options,
                    format_func=lambda x: f"{ICON.get(x, '')} {x}",
                    key="add_category",
                )
            source = all_data[all_data["النوع"].eq(cat)] if not all_data.empty else pd.DataFrame()
            cols = [c for c in compact_columns(source) if c not in {"الشهر", "النوع", "__key__"}]
            common = [c for c in ["م","كود العميل","اسم العميل","تيم ليدر","رقم العمارة","رقم الوحده","المساحه","اسم المرحلة","المشروع","مصدر العميل","تاريخ الحجز","تاريخ التعاقد","تاريخ الالغاء","اجمالي الوحده","المقدم","مبلع الحجز","طريقة الدفع","اسم البائع","ملحوظات"] if c in cols]
            fields = {}
            if not common:
                common = ["م", "اسم العميل", "رقم الوحده", "ملحوظات"]
            grid = st.columns(2)
            for i, c in enumerate(common):
                with grid[i % 2]:
                    fields[c] = st.text_input(c, value=str(next_serial(source)) if c == "م" else "", key=f"add_{c}")
            extra = st.text_area("ملاحظات إضافية (اختياري)", key="add_extra")
            add = st.form_submit_button("إضافة السجل", type="primary", use_container_width=True)
        if add:
            row = {c: fields.get(c, "") for c in common}
            row["الشهر"] = month
            row["النوع"] = cat
            if extra:
                row["ملحوظات"] = extra
            register_added(changes, row)
            st.success("تمت إضافة السجل وحفظه محليًا.")
            st.rerun()

    with tabs[2]:
        # FIX: the original code used `return` when a search came back empty,
        # which silently skipped the Delete tab and the "reset changes" button
        # below. Restructured with if/else so the rest of the page still renders.
        candidates = all_data.copy()
        if candidates.empty:
            st.info("لا توجد سجلات للتعديل.")
        else:
            edit_q = st.text_input("بحث سريع عن السجل", placeholder="اسم العميل أو رقم الوحدة أو الكود...", key="edit_q")
            if edit_q.strip():
                q = edit_q.strip().lower()
                mask = candidates.astype(str).apply(lambda col: col.str.lower().str.contains(q, regex=False, na=False)).any(axis=1)
                candidates = candidates.loc[mask].copy()
            candidates["_label"] = candidates.apply(
                lambda r: f"{r.get('الشهر','')} | {r.get('النوع','')} | {r.get('م','')} | {r.get('اسم العميل','')}",
                axis=1,
            )
            if candidates.empty:
                st.info("لا توجد سجلات مطابقة للبحث.")
            else:
                selected_label = st.selectbox("اختار السجل", candidates["_label"].tolist(), key="edit_select")
                pos = candidates.index[candidates["_label"].eq(selected_label)][0]
                original = candidates.loc[pos].to_dict()
                edit_cols = [c for c in compact_columns(candidates) if c not in {"الشهر", "النوع", "__key__"}]
                edit_values = {}
                with st.form("edit_record_form"):
                    ec = st.columns(2)
                    for i, c in enumerate(edit_cols[:20]):
                        with ec[i % 2]:
                            edit_values[c] = st.text_input(c, value=clean(original.get(c, "")), key=f"edit_{c}")
                    save = st.form_submit_button("حفظ التعديل", type="primary", use_container_width=True)
                if save:
                    register_updated(changes, original, edit_values)
                    st.success("تم تعديل السجل وحفظ التغيير.")
                    st.rerun()

    with tabs[3]:
        # Same fix as the Edit tab above.
        candidates = all_data.copy()
        if candidates.empty:
            st.info("لا توجد سجلات للحذف.")
        else:
            delete_q = st.text_input("بحث سريع عن السجل", placeholder="اسم العميل أو رقم الوحدة أو الكود...", key="delete_q")
            if delete_q.strip():
                q = delete_q.strip().lower()
                mask = candidates.astype(str).apply(lambda col: col.str.lower().str.contains(q, regex=False, na=False)).any(axis=1)
                candidates = candidates.loc[mask].copy()
            candidates["_label"] = candidates.apply(
                lambda r: f"{r.get('الشهر','')} | {r.get('النوع','')} | {r.get('م','')} | {r.get('اسم العميل','')}",
                axis=1,
            )
            if candidates.empty:
                st.info("لا توجد سجلات مطابقة للبحث.")
            else:
                selected_label = st.selectbox("اختار السجل للحذف", candidates["_label"].tolist(), key="delete_select")
                pos = candidates.index[candidates["_label"].eq(selected_label)][0]
                row = candidates.loc[pos].to_dict()
                st.warning("الحذف هنا يحذف السجل من طبقة التطبيق المحلية فقط، ولن يحذف صفًا من Google Sheet.")
                if st.button("حذف السجل", type="primary", use_container_width=True):
                    register_deleted(changes, row)
                    st.success("تم حذف السجل من العرض.")
                    st.rerun()

    if st.button("↩️ إزالة كل التعديلات المحلية والعودة للمصدر", use_container_width=True):
        reset_local_changes()
        st.success("تمت إزالة التعديلات المحلية.")
        st.rerun()

# =========================================================
# MAIN
# =========================================================
def main():
    load_custom_months()
    with st.sidebar:
        sidebar_brand()  # shared logo + workspace caption, same as every other page

        if st.button("🔄 تحديث البيانات", use_container_width=True):
            st.cache_data.clear()
            st.rerun()

        st.divider()
        view_mode = st.radio(
            "Select option",
            ["نظرة عامة", "الحجوزات", "التعاقدات", "الالغاءات", "كل البيانات", "إدارة البيانات"],
            label_visibility="visible",
        )

        selected_month = st.selectbox("اختيار الشهر", ["كل الشهور"] + ordered_months(MONTHS))
        search = st.text_input("بحث سريع", placeholder="اسم العميل، الوحدة، العمارة...")

    with st.spinner("جاري تحميل البيانات..."):
        all_data, diagnostics, raw_sizes, failures = load_all_months()

    all_data, changes = apply_changes(all_data)

    if all_data.empty:
        st.error("لم نتمكن من تحميل البيانات الآن.")
        return

    # Base view
    if view_mode == "إدارة البيانات":
        st.markdown("<div class='hero'><div class='hero-title'>إدارة البيانات</div><div class='hero-sub'>إضافة شهر، إضافة سجل، تعديل أو حذف — مع الحفاظ على Google Sheet كما هو.</div></div>", unsafe_allow_html=True)
        add_management_ui(all_data, changes)
        return

    if view_mode == "نظرة عامة" or view_mode == "كل البيانات":
        current = all_data.copy()
    else:
        current = all_data[all_data["النوع"] == view_mode].copy()

    if selected_month != "كل الشهور":
        current = current[current["الشهر"] == selected_month].copy()

    if search.strip():
        q = search.strip().lower()
        mask = pd.Series(False, index=current.index)
        for c in current.columns:
            mask |= current[c].astype(str).str.lower().str.contains(q, regex=False, na=False)
        current = current.loc[mask].copy()

    # Header — shared banner (logo + title + colored department pill), same as every other page
    render_header("Operations")
    st.markdown(
        f"<span class='small-tag'>"
        f"{('كل الشهور' if selected_month == 'كل الشهور' else selected_month)} • {len(current):,} سجل ظاهر"
        f"</span>",
        unsafe_allow_html=True,
    )
    st.markdown("<br>", unsafe_allow_html=True)

    # =====================================================
    # OVERVIEW
    # =====================================================
    if view_mode == "نظرة عامة":
        # The main dashboard follows the selected month. When "كل الشهور" is
        # selected, it naturally falls back to the full dataset.
        overview_data = current.copy()
        counts = overview_data["النوع"].value_counts()
        total_records = len(overview_data)
        booking_rows = overview_data[overview_data["النوع"].eq("الحجوزات")]
        contract_rows = overview_data[overview_data["النوع"].eq("التعاقدات")]
        cancel_rows = overview_data[overview_data["النوع"].eq("الالغاءات")]

        booking_value_col = category_value_col(booking_rows, "الحجوزات")
        contract_value_col = category_value_col(contract_rows, "التعاقدات")
        cancel_value_col = category_value_col(cancel_rows, "الالغاءات")

        booking_value = numeric_value(booking_rows[booking_value_col]).sum() if booking_value_col else np.nan
        contract_value = numeric_value(contract_rows[contract_value_col]).sum() if contract_value_col else np.nan
        cancel_value = numeric_value(cancel_rows[cancel_value_col]).sum() if cancel_value_col else np.nan
        k1, k2, k3, k4 = st.columns(4)
        cards = [
            (k1, "إجمالي السجلات", f"{total_records:,}", "حسب الاختيار الحالي"),
            (k2, "الحجوزات", f"{int(counts.get('الحجوزات',0)):,}", "حسب الشهر المحدد"),
            (k3, "التعاقدات", f"{int(counts.get('التعاقدات',0)):,}", "حسب الشهر المحدد"),
            (k4, "الإلغاءات", f"{int(counts.get('الالغاءات',0)):,}", "حسب الشهر المحدد"),
        ]
        for col, label, val, note in cards:
            with col:
                st.markdown(f"<div class='kpi'><div class='kpi-label'>{label}</div><div class='kpi-value'>{val}</div><div class='kpi-note'>{note}</div></div>", unsafe_allow_html=True)

        st.markdown("### ملخص شهري")
        if selected_month == "كل الشهور":
            month_order = ordered_months(overview_data["الشهر"].unique())
            summary = (
                overview_data.groupby(["الشهر", "النوع"])
                .size()
                .unstack(fill_value=0)
                .reindex(list(reversed(month_order)))
            )
            for c in CATEGORIES:
                if c not in summary.columns:
                    summary[c] = 0
            summary = summary[CATEGORIES]
            st.bar_chart(summary, use_container_width=True)

        c1, c2 = st.columns(2)
        with c1:
            st.markdown("#### توزيع الأقسام")
            st.bar_chart(
                counts.reindex(CATEGORIES).fillna(0),
                use_container_width=True
            )
        with c2:
            st.markdown("#### إجمالي السجلات حسب الشهر")
            month_order = ordered_months(overview_data["الشهر"].unique())
            by_month = overview_data.groupby("الشهر").size().reindex(list(reversed(month_order))).fillna(0)
            st.bar_chart(by_month, use_container_width=True)

        st.markdown("### القيم المالية حسب القسم")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.metric("إجمالي قيم الحجوزات", f"{booking_value:,.0f}" if pd.notna(booking_value) else "—")
            st.caption(f"المصدر: {booking_value_col or 'غير متاح'}")
        with c2:
            st.metric("إجمالي قيمة التعاقدات", f"{contract_value:,.0f}" if pd.notna(contract_value) else "—")
            st.caption(f"المصدر: {contract_value_col or 'غير متاح'}")
        with c3:
            st.metric("إجمالي قيم الإلغاءات", f"{cancel_value:,.0f}" if pd.notna(cancel_value) else "—")
            st.caption(f"المصدر: {cancel_value_col or 'غير متاح'}")
        with c4:
            ac = area_col(all_data)
            area_total = numeric_value(overview_data[ac]).sum() if ac else np.nan
            st.metric("إجمالي المساحة", f"{area_total:,.0f}" if pd.notna(area_total) else "—")

    else:
        # =================================================
        # DETAIL VIEW
        # =================================================
        if current.empty:
            st.warning("لا توجد نتائج مطابقة للاختيار الحالي.")
            return

        # Smart filters based on real columns.
        st.markdown("### تصفية البيانات")
        filter_names = ["تيم ليدر", "المشروع", "مصدر العميل", "اسم المرحلة", "اسم البائع"]
        filters = {}
        available = [c for c in filter_names if c in current.columns]
        cols = st.columns(min(5, max(1, len(available)))) if available else []
        for i, c in enumerate(available[:5]):
            vals = sorted([x for x in current[c].dropna().astype(str).unique().tolist() if x.strip()])
            with cols[i]:
                filters[c] = st.multiselect(c, vals, key=f"f_{view_mode}_{selected_month}_{c}")
        for c, vals in filters.items():
            if vals:
                current = current[current[c].isin(vals)]

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("السجلات", f"{len(current):,}")
        vc = category_value_col(current, view_mode) if view_mode in CATEGORIES else money_col(current)
        v = numeric_value(current[vc]).sum() if vc else np.nan
        value_label = {
            "الحجوزات": "إجمالي قيم الحجوزات",
            "التعاقدات": "إجمالي قيمة التعاقدات",
            "الالغاءات": "إجمالي قيم الإلغاءات",
        }.get(view_mode, "إجمالي القيمة")
        c2.metric(value_label, f"{v:,.0f}" if pd.notna(v) else "—")
        ac = area_col(current)
        a = numeric_value(current[ac]).sum() if ac else np.nan
        c3.metric("إجمالي المساحة", f"{a:,.0f}" if pd.notna(a) else "—")
        client_col = find_col(current, ["اسم العميل", "Name of client"])
        unique_clients = (
            current[client_col].replace("", np.nan).dropna().nunique()
            if client_col else 0
        )
        c4.metric("العملاء", f"{unique_clients:,}" if client_col else "—")

        # Charts based on actual content.
        ch1, ch2 = st.columns(2)
        with ch1:
            by_month = current.groupby("الشهر").size().reindex(ordered_months(current["الشهر"].unique())).fillna(0)
            st.markdown("#### السجلات حسب الشهر")
            st.bar_chart(by_month)
        with ch2:
            # Use a useful operational dimension instead of Team Leader.
            alternative_col = find_col(
                current,
                ["الحالة", "المشروع", "مصدر العميل", "اسم المرحلة", "طريقة الدفع", "STATUS", "Status"]
            )
            if alternative_col:
                dist = (
                    current[alternative_col]
                    .replace("", np.nan)
                    .dropna()
                    .value_counts()
                    .head(10)
                )
                if not dist.empty:
                    st.markdown(f"#### التوزيع حسب {alternative_col}")
                    st.bar_chart(dist, use_container_width=True)
                else:
                    st.info("لا توجد قيم كافية لعرض الرسم.")
            else:
                st.info("لا يوجد بُعد تشغيلي مناسب للرسم في هذا القسم.")

        st.markdown("### البيانات")
        display = prepare_view(reorder_columns(current))
        st.dataframe(display, use_container_width=True, hide_index=True, height=620)

        csv_bytes = display.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "⬇️ تنزيل البيانات الحالية CSV",
            csv_bytes,
            file_name=f"EL_MOLTQA_{view_mode}_{selected_month}.csv",
            mime="text/csv",
            use_container_width=True,
        )

    # =====================================================
    # FULL DATA
    # =====================================================
    if view_mode == "كل البيانات":
        st.markdown("### البيانات الكاملة")
        display = prepare_view(reorder_columns(current))
        st.dataframe(display, use_container_width=True, hide_index=True, height=680)
        st.download_button(
            "⬇️ تنزيل كل البيانات CSV",
            display.to_csv(index=False).encode("utf-8-sig"),
            file_name="EL_MOLTQA_All_Data.csv",
            mime="text/csv",
            use_container_width=True,
        )

    # =====================================================
    # DIAGNOSTICS — useful, but not in the way
    # =====================================================
    with st.expander("فحص مصدر البيانات"):
        for month in MONTHS:
            info = diagnostics.get(month, {})
            sizes = raw_sizes.get(month, (0, 0, "-"))
            st.write(
                f"**{month}** — Tab `{info.get('tab', '-')}` — "
                f"{sizes[0]:,} صف × {sizes[1]:,} عمود — مصدر القراءة: `{sizes[2]}`"
            )
            d = info.get("diagnostics", {})
            for cat in CATEGORIES:
                # d[cat] is now a list — one entry per time this category's
                # banner appeared in the tab (see parse_month above).
                for x in d.get(cat, []):
                    if not x.get("header_row"):
                        continue
                    st.caption(
                        f"{ICON[cat]} {cat}: section {x.get('section_row')}, "
                        f"header {x.get('header_row')}, "
                        f"subheader {x.get('subheader_row') or 'بدون'}"
                    )
        if failures:
            st.warning("تعذر تحميل: " + ", ".join(failures))

        st.markdown("---")
        st.markdown("**معاينة البيانات المستخرجة (للتحقق من صحة القراءة):**")
        preview_month = st.selectbox("اختر شهرًا للمعاينة", ordered_months(MONTHS), key="diag_preview_month")
        preview_cat = st.selectbox("اختر القسم", CATEGORIES, format_func=lambda c: f"{ICON[c]} {c}", key="diag_preview_cat")
        preview_rows = all_data[(all_data["الشهر"] == preview_month) & (all_data["النوع"] == preview_cat)]
        st.caption(f"عدد السجلات: {len(preview_rows):,} — الأعمدة المكتشفة: {', '.join(preview_rows.columns[:8])}")
        st.dataframe(preview_rows.head(5), use_container_width=True, hide_index=True)


main()