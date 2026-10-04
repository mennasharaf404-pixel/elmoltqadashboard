import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
from theme import inject_css, render_header, sidebar_brand, NAVY, ACCENT_GOLD, ACCENT_TEAL

inject_css()
sidebar_brand()
render_header("Marketing")


@st.cache_data
def load_data():
    rng = np.random.default_rng(4)
    dates = pd.date_range("2025-09-01", periods=52, freq="W")
    df = pd.DataFrame({
        "week": dates,
        "leads": rng.integers(60, 220, len(dates)),
        "conversions": rng.integers(10, 55, len(dates)),
        "spend": rng.integers(3000, 9000, len(dates)),
    })
    df["conversion_rate"] = df["conversions"] / df["leads"]
    channels = pd.DataFrame({
        "channel": ["Social Ads", "Search (SEM)", "Email", "Referral", "Events"],
        "leads": [520, 410, 260, 190, 90],
    })
    return df, channels


df, channels = load_data()

with st.sidebar:
    st.markdown("#### Filters")
    weeks_back = st.slider("Weeks to show", 4, 52, 26)
filtered = df.tail(weeks_back)

total_leads = filtered["leads"].sum()
total_conv = filtered["conversions"].sum()
total_spend = filtered["spend"].sum()
cpl = total_spend / total_leads

c1, c2, c3, c4 = st.columns(4)
c1.metric("Leads Generated", f"{total_leads:,}")
c2.metric("Conversions", f"{total_conv:,}")
c3.metric("Conversion Rate", f"{(total_conv/total_leads)*100:.1f}%")
c4.metric("Cost per Lead", f"${cpl:,.0f}")

st.markdown("<br>", unsafe_allow_html=True)

left, right = st.columns([2, 1])
with left:
    st.markdown("##### Leads & Conversions Over Time")
    fig = px.bar(filtered, x="week", y=["leads", "conversions"], barmode="overlay",
                 color_discrete_sequence=[NAVY, ACCENT_GOLD], opacity=0.85)
    fig.update_layout(margin=dict(t=10, l=0, r=0, b=0), height=380, legend=dict(orientation="h", y=1.12))
    st.plotly_chart(fig, use_container_width=True)

with right:
    st.markdown("##### Leads by Channel")
    fig2 = px.bar(channels.sort_values("leads"), x="leads", y="channel", orientation="h",
                   color="leads", color_continuous_scale=[NAVY, ACCENT_TEAL])
    fig2.update_layout(margin=dict(t=10, l=0, r=0, b=0), height=380, coloraxis_showscale=False)
    st.plotly_chart(fig2, use_container_width=True)

st.markdown("##### Weekly Ad Spend")
fig3 = px.area(filtered, x="week", y="spend", color_discrete_sequence=[ACCENT_TEAL])
fig3.update_layout(margin=dict(t=10, l=0, r=0, b=0), height=280)
st.plotly_chart(fig3, use_container_width=True)
