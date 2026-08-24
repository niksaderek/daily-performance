# ABOUTME: Aggregates campaign rows into the headline KPIs, daily trends, publisher
# ABOUTME: rankings, funnel stages, and affiliate-versus-internal splits shown on the dashboard.

import pandas as pd

FUNNEL_STAGES = ["Incoming Calls", "Connected", "Converted"]


def _ratio(numerator, denominator, scale=100.0):
    """Percentage helper that returns 0 instead of raising when nothing was measured."""
    if not denominator:
        return 0
    return numerator / denominator * scale


def headline_kpis(df):
    """Roll the selection up into the numbers shown in the KPI strip."""
    spend = df["Spend"].sum()
    revenue = df["Revenue"].sum()
    connected = df["Connected"].sum()

    return {
        "spend": spend,
        "revenue": revenue,
        "net_profit": df["Net_Profit"].sum(),
        "roi_pct": _ratio(revenue - spend, spend),
        "margin_pct": _ratio(df["Net_Profit"].sum(), revenue),
        "conversion_pct": _ratio(df["Converted"].sum(), connected),
        "incoming": df["Incoming"].sum(),
        "campaigns": len(df),
    }


def daily_summary(df):
    """One row per date, so trends can be plotted across the reporting window."""
    rows = []
    for date, day_data in df.groupby("Date", sort=True):
        rows.append(
            {
                "Date": date,
                "Label": date.strftime("%b %d") if hasattr(date, "strftime") else str(date),
                "Day": day_data["Day"].iloc[0] if "Day" in day_data else "",
                "Campaigns": len(day_data),
                "Spend": day_data["Spend"].sum(),
                "Revenue": day_data["Revenue"].sum(),
                "Net_Profit": day_data["Net_Profit"].sum(),
                "ROI": round(_ratio(day_data["Revenue"].sum() - day_data["Spend"].sum(),
                                    day_data["Spend"].sum()), 2),
                "Incoming": day_data["Incoming"].sum(),
                "Connected": day_data["Connected"].sum(),
                "Converted": day_data["Converted"].sum(),
                "Conv_Rate": round(_ratio(day_data["Converted"].sum(),
                                          day_data["Connected"].sum()), 2),
            }
        )
    return pd.DataFrame(rows)


def publisher_leaderboard(df, limit=10):
    """Rank media buyers by profit, with the efficiency metrics used in tooltips."""
    board = df.groupby("Media_Buyer").agg(
        Net_Profit=("Net_Profit", "sum"),
        Revenue=("Revenue", "sum"),
        Spend=("Spend", "sum"),
        Conversion_Rate=("Conversion_Rate", "mean"),
    )
    board["ROI"] = board.apply(
        lambda r: round(_ratio(r["Revenue"] - r["Spend"], r["Spend"]), 2), axis=1
    )
    board["Margin"] = board.apply(
        lambda r: round(_ratio(r["Net_Profit"], r["Revenue"]), 2), axis=1
    )
    return board.sort_values("Net_Profit", ascending=False).head(limit).round(2)


def conversion_funnel(df):
    """Return the funnel stage labels alongside their call volumes."""
    values = [
        df["Incoming"].sum(),
        df["Connected"].sum(),
        df["Converted"].sum(),
    ]
    return list(FUNNEL_STAGES), values


def _segment_metrics(segment):
    revenue = segment["Revenue"].sum()
    spend = segment["Spend"].sum()

    return {
        "net_profit": segment["Net_Profit"].sum() if len(segment) else 0,
        "revenue": revenue,
        "spend": spend,
        "roi_pct": _ratio(revenue - spend, spend),
        "margin_pct": _ratio(segment["Net_Profit"].sum(), revenue),
        "conversion_pct": _ratio(segment["Converted"].sum(), segment["Connected"].sum()),
    }


def segment_comparison(df):
    """Compare affiliate-sourced traffic against internally bought traffic."""
    return {
        "Affiliate": _segment_metrics(df[df["Is_Affiliate"]]),
        "Internal": _segment_metrics(df[~df["Is_Affiliate"]]),
    }
