from __future__ import annotations

import io
import os
import re
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st
from theme import render_header

# Configuration
DEPARTMENT = "Marketing"
DATA_DIR = Path(os.environ.get("MARKETING_DATA_DIR", Path(__file__).resolve().parent.parent / "data"))
NAME_COL, TOTAL_COL = "Full Name", "Total"
AGENT, LEADS = "Agent", "Leads"
FILTER_PREFIX = "mkt_f_"

# Decision stages (display order). Edit STATUS_META below to move a CRM status to another stage.
WON, HOT, NURTURING, UNREACHABLE, LOST, UNQUALIFIED, OTHER = (
    "Won", "Hot", "Nurturing", "Unreachable", "Lost", "Unqualified", "Other")
STAGES = [WON, HOT, NURTURING, UNREACHABLE, LOST, UNQUALIFIED, OTHER]
STAGE_COLORS = {WON: "#12A150", HOT: "#1F4E9A", NURTURING: "#7B3FBF", UNREACHABLE: "#F28C0F",
                LOST: "#D62728", UNQUALIFIED: "#7F7F7F", OTHER: "#BDBDBD"}

# CRM column (lower-case) -> (display name, stage)
STATUS_META = {
    "fresh leads": ("Fresh Leads", NURTURING),
    "cold calls": ("Cold Calls", NURTURING),
    "follow up": ("Follow Up", NURTURING),
    "rotation": ("Rotation", NURTURING),
    "hold": ("Hold", NURTURING),
    "interested": ("Interested", HOT),
    "immediatley dielvery": ("Immediate Delivery", HOT),
    "scheduled meeting": ("Scheduled Meeting", HOT),
    "follow up after meeting": ("Follow-up After Meeting", HOT),
    "reservation": ("Reservation", HOT),
    "important lead": ("Important Lead", HOT),
    "done deals": ("Done Deals", WON),
    "old done deals": ("Old Done Deals", WON),
    "no answer": ("No Answer", UNREACHABLE),
    "no answer at all": ("No Answer At All", UNREACHABLE),
    "no answer after follow": ("No Answer After Follow-up", UNREACHABLE),
    "not interested": ("Not Interested", LOST),
    "low budget": ("Low Budget", LOST),
    "deal with competitor": ("Deal With Competitor", LOST),
    "not available request..": ("Requested Unit Unavailable", LOST),
    "junk leads": ("Junk Leads", UNQUALIFIED),
    "wrong number": ("Wrong Number", UNQUALIFIED),
    "out of service": ("Out Of Service", UNQUALIFIED),
    "not registered": ("Not Registered", UNQUALIFIED),
    "broker": ("Broker", UNQUALIFIED),
    "old data": ("Old Data", UNQUALIFIED),
    "data bank": ("Data Bank", UNQUALIFIED),
    "archive": ("Archive", UNQUALIFIED),
}
STAGE_BY_LABEL = {label: stage for label, stage in STATUS_META.values()}


# Helpers
def _meta(column: str) -> tuple[str, str]:
    return STATUS_META.get(column.strip().lower(), (column, OTHER))


def _stage_col(stage: str) -> str:
    return f"_{stage}"


def _status_cols(df: pd.DataFrame) -> list[str]:
    return [c for c in df.columns if c not in (AGENT, LEADS) and not c.startswith("_")]


def _stage_scale() -> alt.Scale:
    return alt.Scale(domain=STAGES, range=[STAGE_COLORS[s] for s in STAGES])


# Cleaning functions
def _clean_name(value) -> str | None:
    if not isinstance(value, str):
        return None
    text = re.sub(r"[\u064B-\u0652\u200e\u200f]", "", value.replace("\u00a0", " "))  # stray marks
    text = re.sub(r"\s+", " ", text).strip()
    return text.title() if text else None


def clean_report(df: pd.DataFrame) -> pd.DataFrame:
    """One row per agent; counts as integers. Values are never changed."""
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")].copy()
    df.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in df.columns]
    name_col = NAME_COL if NAME_COL in df.columns else df.columns[0]

    df[AGENT] = df[name_col].map(_clean_name)
    df = df[df[AGENT].notna()]
    counts = [c for c in df.columns if c not in (name_col, AGENT)]
    for col in counts:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    df = df.groupby(AGENT, as_index=False)[counts].sum()
    status = [c for c in counts if c != TOTAL_COL]
    df["_sum"] = df[status].sum(axis=1)
    df[LEADS] = df[TOTAL_COL] if TOTAL_COL in df.columns else df["_sum"]
    df = df.drop(columns=[TOTAL_COL], errors="ignore")
    return df.rename(columns={c: _meta(c)[0] for c in status})


@st.cache_data(show_spinner="Cleaning data...")
def load_data(raw: bytes) -> pd.DataFrame:
    return clean_report(pd.read_excel(io.BytesIO(raw), sheet_name=0, dtype=object))


def add_derived(df: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, str], list[str]]:
    """Analysis-only columns: leads per stage for every agent."""
    out = df.copy()
    status = _status_cols(out)
    stage_of = {c: STAGE_BY_LABEL.get(c, OTHER) for c in status}
    for stage in STAGES:
        cols = [c for c in status if stage_of[c] == stage]
        out[_stage_col(stage)] = out[cols].sum(axis=1) if cols else 0
    unknown = [c for c in status if stage_of[c] == OTHER]
    return out, stage_of, unknown


def agent_table(view: pd.DataFrame) -> pd.DataFrame:
    ag = pd.DataFrame({AGENT: view[AGENT], LEADS: view[LEADS]})
    for stage in STAGES:
        ag[stage] = view[_stage_col(stage)]
    ag["Contact rate"] = 1 - ag[UNREACHABLE] / ag[LEADS]
    ag["Hot + Won %"] = (ag[HOT] + ag[WON]) / ag[LEADS]
    ag["Lost %"] = ag[LOST] / ag[LEADS]
    ag["Unqualified %"] = ag[UNQUALIFIED] / ag[LEADS]
    return ag.sort_values(LEADS, ascending=False).reset_index(drop=True)


def status_totals_table(view: pd.DataFrame, stage_of: dict[str, str]) -> pd.DataFrame:
    t = view[list(stage_of)].sum().rename_axis("Status").reset_index(name="Leads")
    t["Stage"] = t["Status"].map(stage_of)
    return t[t["Leads"] > 0].reset_index(drop=True)


# Charts #
TOP_N = 10  # agent charts show the biggest agents so they stay readable
BAR = 26    # same bar thickness everywhere
H = 300     # same chart height everywhere (keeps the grid aligned)


def _pad(series, factor: float = 1.3) -> alt.Scale:
    """Leave room on the right so the value labels are not clipped."""
    return alt.Scale(domain=[0, max(float(series.max()), 1) * factor], nice=False)


def stage_chart(view: pd.DataFrame, leads: int):
    """Where all leads sit, by decision stage."""
    d = pd.DataFrame({"Stage": STAGES, "Leads": [int(view[_stage_col(s)].sum()) for s in STAGES]})
    d = d[d["Leads"] > 0].copy()
    if d.empty:
        return None
    d["txt"] = [f"{v:,} ({v / leads:.0%})" for v in d["Leads"]]
    base = alt.Chart(d).encode(
        y=alt.Y("Stage:N", sort=STAGES, title=None),
        x=alt.X("Leads:Q", title=None, scale=_pad(d["Leads"]), axis=None),
    )
    bars = base.mark_bar(size=BAR).encode(
        color=alt.Color("Stage:N", scale=_stage_scale(), legend=None),
        tooltip=["Stage:N", alt.Tooltip("Leads:Q", format=",")],
    )
    labels = base.mark_text(align="left", dx=4, fontSize=11).encode(text="txt:N")
    return (bars + labels).properties(title="Leads by stage", height=H, width="container")


def loss_chart(totals: pd.DataFrame, leads: int):
    """Why leads are lost."""
    d = totals[totals["Stage"] == LOST].sort_values("Leads", ascending=False).copy()
    if d.empty:
        return None
    d["txt"] = [f"{v:,} ({v / leads:.0%})" for v in d["Leads"]]
    base = alt.Chart(d).encode(
        y=alt.Y("Status:N", sort=d["Status"].tolist(), title=None, axis=alt.Axis(labelLimit=280)),
        x=alt.X("Leads:Q", title=None, scale=_pad(d["Leads"]), axis=None),
    )
    bars = base.mark_bar(size=BAR, color=STAGE_COLORS[LOST]).encode(
        tooltip=["Status:N", alt.Tooltip("Leads:Q", format=",")])
    labels = base.mark_text(align="left", dx=4, fontSize=11).encode(text="txt:N")
    return (bars + labels).properties(title="Why leads are lost", height=H, width="container")


def agent_mix_chart(view: pd.DataFrame):
    """Share of each agent's leads in every stage (lead quality, not just size)."""
    top = view.sort_values(LEADS, ascending=False).head(TOP_N)
    rows = [{"agent": r[AGENT], "stage": s, "ord": i, "leads": int(r[_stage_col(s)])}
            for _, r in top.iterrows() for i, s in enumerate(STAGES) if int(r[_stage_col(s)])]
    if not rows:
        return None
    title = "Stage mix per agent" + (f" (top {TOP_N} by leads)" if len(view) > TOP_N else "")
    return alt.Chart(pd.DataFrame(rows)).mark_bar(size=BAR - 6).encode(
        y=alt.Y("agent:N", sort=top[AGENT].tolist(), title=None, axis=alt.Axis(labelLimit=200)),
        x=alt.X("leads:Q", stack="normalize", title=None, axis=alt.Axis(format="%")),
        color=alt.Color("stage:N", scale=_stage_scale(), title=None, sort=STAGES,
                        legend=alt.Legend(orient="top", columns=7, symbolLimit=7)),
        order=alt.Order("ord:Q"),
        tooltip=["agent:N", "stage:N", alt.Tooltip("leads:Q", format=",")],
    ).properties(title=title, height=H, width="container")


def contact_chart(ag: pd.DataFrame, team_rate: float):
    """Contact rate per agent against the team average."""
    d = ag.head(TOP_N).copy()
    if d.empty:
        return None
    d = d.sort_values("Contact rate", ascending=False)
    d["txt"] = [f"{v:.0%}" for v in d["Contact rate"]]
    d = d.rename(columns={"Contact rate": "rate"})
    base = alt.Chart(d).encode(
        y=alt.Y("Agent:N", sort=d["Agent"].tolist(), title=None, axis=alt.Axis(labelLimit=200)),
        x=alt.X("rate:Q", title=None, scale=alt.Scale(domain=[0, 1.15], nice=False), axis=None),
    )
    bars = base.mark_bar(size=BAR - 6, color=STAGE_COLORS[HOT]).encode(
        tooltip=["Agent:N", alt.Tooltip("rate:Q", title="Contact rate", format=".0%"),
                 alt.Tooltip("Leads:Q", format=",")])
    labels = base.mark_text(align="left", dx=4, fontSize=11).encode(text="txt:N")
    rule = alt.Chart(pd.DataFrame({"x": [team_rate]})).mark_rule(
        strokeDash=[4, 4], color=STAGE_COLORS[LOST], size=2).encode(x="x:Q")
    title = f"Contact rate by agent (dashed line = team {team_rate:.0%})"
    if len(ag) > TOP_N:
        title = f"Contact rate, top {TOP_N} agents by leads (line = team {team_rate:.0%})"
    return (bars + labels + rule).properties(title=title, height=H, width="container")


def show(chart):
    if chart is not None:
        st.altair_chart(chart)


# Page
def _load_source():
    """Uploaded file wins; otherwise use the newest CRM .xlsx in the data folder."""
    with st.sidebar:
        st.subheader("Data")
        up = st.file_uploader("Upload CRM report (.xlsx file)", type=["xlsx"], key="mkt_upload")
    if up is not None:
        return up.getvalue(), up.name
    if DATA_DIR.exists():
        files = [p for p in DATA_DIR.glob("*.xlsx") if "crm" in p.name.lower() and not p.name.startswith("~$")]
        if files:
            path = max(files, key=lambda p: p.stat().st_mtime)
            return path.read_bytes(), path.name
    return None, None


def _reset_filters() -> None:
    for key in [k for k in st.session_state if k.startswith(FILTER_PREFIX)]:
        del st.session_state[key]


def _excel_bytes(frame: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        frame.to_excel(xw, index=False, sheet_name="Agents")
    return buf.getvalue()


def main():
    render_header(DEPARTMENT)

    raw, source = _load_source()
    if raw is None:
        st.warning("Upload the CRM report")
        st.stop()
    try:
        base = load_data(raw)
    except Exception as exc:
        st.error(f"Could not read the file ({exc}).")
        st.stop()
    if LEADS not in base.columns or base.empty:
        st.error("The file does not look like a CRM agent report.")
        st.stop()
    df, stage_of, unknown = add_derived(base)

    # Filters #
    active = df[df[LEADS] > 0]
    with st.sidebar:
        st.subheader("Filters")
        agents = st.multiselect("Agent", sorted(active[AGENT]), key=FILTER_PREFIX + "agent")
        st.button("↩️ Reset filters", on_click=_reset_filters)
        if len(df) - len(active):
            st.caption(f"{len(df) - len(active)} agent(s) with no leads are hidden.")
    view = active[active[AGENT].isin(agents)] if agents else active
    if view.empty:
        st.info("No leads for the selected agents.")
        st.stop()

    # KPIs #
    leads = int(view[LEADS].sum())
    s = {st_: int(view[_stage_col(st_)].sum()) for st_ in STAGES}
    contacted = leads - s[UNREACHABLE]
    contact_rate = contacted / leads if leads else 0.0
    totals = status_totals_table(view, stage_of)

    k = st.columns(5)
    k[0].metric("Total leads", f"{leads:,}", f"{leads / len(view):,.0f} per agent", delta_color="off",
                help="Leads in the CRM report for the selected agents")
    k[1].metric("Contact rate", f"{contact_rate:.0%}", f"{s[UNREACHABLE]:,} unreachable", delta_color="off",
                help="Share of leads that answered at least once (everything except the 'No Answer' statuses)")
    k[2].metric("Deals won", f"{s[WON]:,}", f"{s[WON] / leads:.1%} of leads", delta_color="off",
                help="Done Deals + Old Done Deals")
    k[3].metric("Hot leads", f"{s[HOT]:,}", f"{s[HOT] / leads:.1%} of leads", delta_color="off",
                help="Interested, meetings, reservations, immediate delivery, important leads")
    k[4].metric("Lost", f"{s[LOST] / leads:.0%}", f"{s[LOST]:,} leads", delta_color="off",
                help="Not interested, low budget, competitor, unit unavailable")

    # Charts #
    c1, c2 = st.columns(2)
    with c1:
        show(stage_chart(view, leads))
    with c2:
        show(loss_chart(totals, leads))

    c3, c4 = st.columns(2)
    with c3:
        show(agent_mix_chart(view))
    with c4:
        show(contact_chart(agent_table(view), contact_rate))

    # Agent scorecard #
    st.subheader("Agent scorecard")
    st.caption("Contact rate = leads minus the 'No Answer' statuses. "
               "Stages are defined in STATUS_META at the top of the file. 📥 Download the table below.")
    q = st.text_input("Search agent")
    ag = agent_table(view)
    if q:
        ag = ag[ag[AGENT].str.contains(q, case=False, regex=False)]
    pct_cols = ["Contact rate", "Hot + Won %", "Lost %", "Unqualified %"]
    shown = ag[[AGENT, LEADS, HOT, WON] + pct_cols].copy()
    for c in pct_cols:
        shown[c] = (shown[c] * 100).round(0)
    st.dataframe(shown, hide_index=True, column_config={
        c: st.column_config.ProgressColumn(c, min_value=0, max_value=100, format="%d%%") for c in pct_cols})
    st.download_button("⬇️ Download Excel", _excel_bytes(shown), file_name="Marketing_Agents.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    mismatch = df[df["_sum"] != df[LEADS]]
    if len(mismatch):
        st.warning("Status counts do not add up to 'Total' for: " + ", ".join(mismatch[AGENT]))
    if unknown:
        st.caption("Statuses not mapped to a stage (add them to STATUS_META): " + ", ".join(unknown))


main()