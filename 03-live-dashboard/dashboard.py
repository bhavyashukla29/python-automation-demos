"""Live sales dashboard over the cleaned data from the report-automation demo.

Filters by date, channel and product; every number and chart updates instantly.
KPIs compare against the previous period of the same length.

Usage:
    py -3.10 -m streamlit run dashboard.py
"""

from datetime import date
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

DATA = Path(__file__).resolve().parent.parent / "02-sales-report-automation" / "output" / "clean_sales.csv"
CHANNEL_COLORS = {"Website": "#1F3D2B", "Amazon": "#E8A33D", "Store": "#6FA287"}

st.set_page_config(page_title="Sales dashboard", page_icon="📈", layout="wide")


@st.cache_data
def load() -> pd.DataFrame:
    return pd.read_csv(DATA, parse_dates=["date"])


def rupees(x: float) -> str:
    if x >= 1e5:
        return f"₹{x / 1e5:.2f} L"  # lakh, the way Indian businesses read revenue
    return f"₹{x:,.0f}"


def pct_change(now: float, before: float) -> str | None:
    return f"{(now - before) / before:+.1%}" if before else None


def query_date(key: str, fallback: date, low: date, high: date) -> date:
    try:
        return min(max(date.fromisoformat(st.query_params[key]), low), high)
    except (KeyError, ValueError):
        return fallback


sales = load()

# ---------- filters ----------
# Defaults can come from the URL (?start=2026-10-01&end=2026-10-31&channels=Amazon,Store),
# so a filtered view can be bookmarked or shared
st.sidebar.header("Filters")
min_d, max_d = sales["date"].min().date(), sales["date"].max().date()
default_range = (query_date("start", min_d, min_d, max_d), query_date("end", max_d, min_d, max_d))
picked = st.sidebar.date_input("Date range", default_range, min_value=min_d, max_value=max_d)
start, end = (picked if isinstance(picked, tuple) and len(picked) == 2 else (min_d, max_d))
all_channels = sorted(sales["channel"].unique())
url_channels = [c for c in st.query_params.get("channels", "").split(",") if c in all_channels]
channels = st.sidebar.multiselect("Channel", all_channels, default=url_channels or all_channels)
products = st.sidebar.multiselect("Product", sorted(sales["product"].unique()), placeholder="All products")

mask = sales["channel"].isin(channels)
if products:
    mask &= sales["product"].isin(products)
scoped = sales[mask]
start_ts, end_ts = pd.Timestamp(start), pd.Timestamp(end)
current = scoped[scoped["date"].between(start_ts, end_ts)]
span = end_ts - start_ts + pd.Timedelta(days=1)
previous = scoped[scoped["date"].between(start_ts - span, start_ts - pd.Timedelta(days=1))]

# ---------- header + KPIs ----------
st.title("Sales dashboard")
st.caption(f"{start:%d %b %Y} – {end:%d %b %Y} · Website, Amazon and store sales in one place · "
           "deltas compare with the previous period of the same length")

if current.empty:
    st.warning("No sales match these filters.")
    st.stop()

rev, prev_rev = current["revenue"].sum(), previous["revenue"].sum()
orders, prev_orders = current["order_id"].nunique(), previous["order_id"].nunique()
units, prev_units = current["qty"].sum(), previous["qty"].sum()
aov = rev / orders
prev_aov = prev_rev / prev_orders if prev_orders else 0

k1, k2, k3, k4 = st.columns(4)
k1.metric("Revenue", rupees(rev), pct_change(rev, prev_rev), border=True)
k2.metric("Orders", f"{orders:,}", pct_change(orders, prev_orders), border=True)
k3.metric("Units sold", f"{units:,}", pct_change(units, prev_units), border=True)
k4.metric("Avg order value", rupees(aov), pct_change(aov, prev_aov), border=True)

# ---------- charts ----------
left, right = st.columns([2, 1])
# A 7-day average instead of weekly totals: partial weeks at the edges would look like crashes
daily = (current.pivot_table(index="date", columns="channel", values="revenue", aggfunc="sum")
         .reindex(pd.date_range(start_ts, end_ts), fill_value=0).fillna(0))
trend = (daily.rolling(7, min_periods=1).mean().rename_axis("date").reset_index()
         .melt(id_vars="date", var_name="channel", value_name="revenue"))
fig = px.area(trend, x="date", y="revenue", color="channel", color_discrete_map=CHANNEL_COLORS,
              labels={"date": "", "revenue": "Revenue per day (₹)", "channel": ""},
              title="Daily revenue by channel (7-day average)")
fig.update_layout(legend=dict(orientation="h", y=1.1), margin=dict(t=60, b=10))
left.plotly_chart(fig, width="stretch")

split = current.groupby("channel")["revenue"].sum().reset_index()
donut = px.pie(split, names="channel", values="revenue", hole=0.6, color="channel",
               color_discrete_map=CHANNEL_COLORS, title="Channel share")
donut.update_traces(textinfo="percent", textposition="inside", insidetextfont_color="white")
donut.update_layout(margin=dict(t=60, b=10), legend=dict(orientation="h", y=-0.05))
right.plotly_chart(donut, width="stretch")

left, right = st.columns([3, 2])
by_product = current.groupby("product").agg(revenue=("revenue", "sum"), units=("qty", "sum")).reset_index()
by_product = by_product.sort_values("revenue")
bar = px.bar(by_product, x="revenue", y="product", orientation="h", text=by_product["revenue"].map(rupees),
             labels={"revenue": "Revenue (₹)", "product": ""}, title="Revenue by product",
             color_discrete_sequence=["#1F3D2B"])
bar.update_layout(margin=dict(t=60, b=10))
left.plotly_chart(bar, width="stretch")

monthly = (current.assign(Month=current["date"].dt.to_period("M").astype(str))
           .pivot_table(index="Month", columns="channel", values="revenue", aggfunc="sum", fill_value=0))
monthly["Total"] = monthly.sum(axis=1)
right.markdown("**Monthly revenue (₹)**")
right.dataframe(monthly.style.format("{:,.0f}"), width="stretch")

st.download_button("Download filtered data (CSV)", current.to_csv(index=False).encode("utf-8"),
                   file_name="sales_filtered.csv", mime="text/csv")
