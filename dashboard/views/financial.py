import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from theme import inject_css, render_header, sidebar_brand, NAVY, ACCENT_GOLD, ACCENT_TEAL

inject_css()
sidebar_brand()
render_header("Financial")


@st.cache_data
def load_data():
    rng = np.random.default_rng(3)
    months = pd.date_range("2025-01-01", periods=12, freq="MS")
    revenue = rng.integers(900_000, 1_600_000, 12).astype(float)
    expenses = revenue * rng.uniform(0.55, 0.75, 12)
    df = pd.DataFrame({
        "month": months,
        "revenue": revenue,
        "expenses": expenses,
        "profit": revenue - expenses,
    })
    breakdown = pd.DataFrame({
        "category": ["Construction", "Sales & Marketing", "Payroll", "Admin", "Other"],
        "amount": [520_000, 180_000, 260_000, 95_000, 60_000],
    })
    return df, breakdown


df, breakdown = load_data()

with st.sidebar:
    st.markdown("#### Filters")
    months_back = st.slider("Months to show", 3, 12, 12)
filtered = df.tail(months_back)

total_rev = filtered["revenue"].sum()
total_exp = filtered["expenses"].sum()
total_profit = filtered["profit"].sum()
margin = total_profit / total_rev

c1, c2, c3, c4 = st.columns(4)
c1.metric("Revenue", f"${total_rev/1e6:.2f}M")
c2.metric("Expenses", f"${total_exp/1e6:.2f}M")
c3.metric("Net Profit", f"${total_profit/1e6:.2f}M")
c4.metric("Profit Margin", f"{margin*100:.1f}%")

st.markdown("<br>", unsafe_allow_html=True)

left, right = st.columns([2, 1])
with left:
    st.markdown("##### Revenue vs Expenses")
    fig = go.Figure()
    fig.add_bar(x=filtered["month"], y=filtered["revenue"], name="Revenue", marker_color=NAVY)
    fig.add_bar(x=filtered["month"], y=filtered["expenses"], name="Expenses", marker_color=ACCENT_GOLD)
    fig.update_layout(barmode="group", margin=dict(t=10, l=0, r=0, b=0), height=380,
                       legend=dict(orientation="h", y=1.12))
    st.plotly_chart(fig, use_container_width=True)

with right:
    st.markdown("##### Expense Breakdown")
    fig2 = px.pie(breakdown, names="category", values="amount", hole=0.55,
                   color_discrete_sequence=[NAVY, ACCENT_TEAL, ACCENT_GOLD, "#8A5A9E", "#C0563B"])
    fig2.update_layout(margin=dict(t=10, l=0, r=0, b=0), height=380)
    st.plotly_chart(fig2, use_container_width=True)

st.markdown("##### Net Profit Trend")
fig3 = px.area(filtered, x="month", y="profit", color_discrete_sequence=[ACCENT_TEAL])
fig3.update_layout(margin=dict(t=10, l=0, r=0, b=0), height=300)
st.plotly_chart(fig3, use_container_width=True)
