# ABOUTME: Streamlit entry point for the daily performance dashboard, wiring sidebar
# ABOUTME: filters through the metric aggregations and into the Plotly panels.

import pandas as pd
import streamlit as st

from src.charts import (
    conversion_funnel_chart,
    profit_trend_chart,
    publisher_profit_chart,
    segment_profit_chart,
    segment_rate_chart,
)
from src.data import apply_filters, load_performance_data
from src.metrics import (
    conversion_funnel,
    daily_summary,
    weekly_summary,
    headline_kpis,
    publisher_leaderboard,
    segment_comparison,
)

st.set_page_config(
    page_title="Daily Performance Dashboard",
    page_icon="📊",
    layout="wide",
)


@st.cache_data
def get_data():
    return load_performance_data()


def render_sidebar(df):
    """Collect the filter selections and return the narrowed frame."""
    st.sidebar.header("Filters")

    earliest = df["Date"].min().date()
    latest = df["Date"].max().date()
    selected_range = st.sidebar.date_input(
        "Date range",
        value=(earliest, latest),
        min_value=earliest,
        max_value=latest,
    )
    # Streamlit returns a single date while the user is mid-way through picking a range.
    if not isinstance(selected_range, tuple) or len(selected_range) != 2:
        selected_range = (earliest, latest)

    selected_verticals = st.sidebar.multiselect(
        "Verticals", options=sorted(df["Vertical"].dropna().unique())
    )
    selected_sources = st.sidebar.multiselect(
        "Traffic sources", options=sorted(df["Traffic_Source"].dropna().unique())
    )
    funded_only = st.sidebar.checkbox(
        "Exclude test campaigns",
        value=True,
        help="Hides rows spending under $50, which are smoke tests rather than funded campaigns.",
    )

    st.sidebar.caption(
        "Partner and buyer names in this dataset are pseudonyms. "
        "The underlying metrics are unchanged."
    )

    return apply_filters(
        df,
        date_range=selected_range,
        verticals=selected_verticals,
        sources=selected_sources,
        funded_only=funded_only,
    )


def render_kpis(kpis):
    """Top strip of headline numbers."""
    columns = st.columns(5)
    columns[0].metric("Net Profit", f"${kpis['net_profit']:,.0f}")
    columns[1].metric("Revenue", f"${kpis['revenue']:,.0f}")
    columns[2].metric("Spend", f"${kpis['spend']:,.0f}")
    columns[3].metric("ROI", f"{kpis['roi_pct']:.1f}%")
    columns[4].metric("Conversion", f"{kpis['conversion_pct']:.1f}%")


def main():
    df = get_data()
    filtered = render_sidebar(df)

    st.title("Daily Performance Dashboard")
    st.caption(
        "Profitability and call-funnel performance across media buyers, "
        "verticals, and traffic sources."
    )

    if filtered.empty:
        st.warning("No campaigns match the current filters.")
        return

    render_kpis(headline_kpis(filtered))
    st.divider()

    summary = daily_summary(filtered)

    trend_header, grain_control = st.columns([3, 1])
    trend_header.subheader("Profit and ROI over time")
    weekly = grain_control.selectbox(
        "Granularity", ["Daily", "Weekly"], index=1 if len(summary) > 45 else 0
    ) == "Weekly"

    trend = weekly_summary(summary) if weekly else summary
    st.plotly_chart(profit_trend_chart(trend), use_container_width=True)
    if weekly and trend.get("Partial", pd.Series(dtype=bool)).any():
        st.caption("* Part-week at the edge of the selected range.")

    left, right = st.columns([3, 2])
    with left:
        st.subheader("Top media buyers by profit")
        st.plotly_chart(
            publisher_profit_chart(publisher_leaderboard(filtered)),
            use_container_width=True,
        )
    with right:
        st.subheader("Call conversion funnel")
        stages, values = conversion_funnel(filtered)
        st.plotly_chart(conversion_funnel_chart(stages, values), use_container_width=True)

    st.subheader("Affiliate versus internal traffic")
    segments = segment_comparison(filtered)
    profit_column, rate_column = st.columns(2)
    with profit_column:
        st.plotly_chart(segment_profit_chart(segments), use_container_width=True)
    with rate_column:
        st.plotly_chart(segment_rate_chart(segments), use_container_width=True)

    with st.expander("Daily summary table"):
        st.dataframe(summary, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
