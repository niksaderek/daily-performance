# ABOUTME: Tests the profitability and funnel aggregations that drive the dashboard,
# ABOUTME: covering ordinary arithmetic and the zero-denominator guards.

import pandas as pd
import pytest

from src.metrics import (
    conversion_funnel,
    daily_summary,
    headline_kpis,
    publisher_leaderboard,
    segment_comparison,
)


@pytest.fixture
def frame():
    return pd.DataFrame(
        {
            "Date": pd.to_datetime(["2025-10-15", "2025-10-15", "2025-10-16"]),
            "Day": ["Wednesday", "Wednesday", "Thursday"],
            "Media_Buyer": ["ALPHA", "BETA", "ALPHA"],
            "Vertical": ["MEDICARE", "AUTO", "MEDICARE"],
            "Traffic_Source": ["GOOGLE", "FACEBOOK", "GOOGLE"],
            "Spend": [1000, 500, 800],
            "Revenue": [1500, 400, 1200],
            "Net_Profit": [400.0, -150.0, 350.0],
            "Incoming": [200.0, 100.0, 160.0],
            "Connected": [120.0, 50.0, 100.0],
            "Converted": [30.0, 10.0, 25.0],
            "Conversion_Rate": [0.25, 0.20, 0.25],
            "No_Connect_Pct": [0.40, 0.50, 0.375],
            "Is_Affiliate": [True, False, True],
        }
    )


def test_headline_kpis_sum_money_and_derive_rates(frame):
    kpis = headline_kpis(frame)

    assert kpis["spend"] == 2300
    assert kpis["revenue"] == 3100
    assert kpis["net_profit"] == pytest.approx(600.0)
    assert kpis["roi_pct"] == pytest.approx((3100 / 2300 - 1) * 100)
    assert kpis["conversion_pct"] == pytest.approx(65 / 270 * 100)


def test_headline_kpis_survive_zero_spend():
    empty = pd.DataFrame(
        {
            "Spend": [0], "Revenue": [0], "Net_Profit": [0.0],
            "Incoming": [0.0], "Connected": [0.0], "Converted": [0.0],
        }
    )
    kpis = headline_kpis(empty)

    assert kpis["roi_pct"] == 0
    assert kpis["conversion_pct"] == 0


def test_daily_summary_has_one_row_per_date(frame):
    summary = daily_summary(frame)

    assert list(summary["Date"]) == list(pd.to_datetime(["2025-10-15", "2025-10-16"]))
    assert summary.loc[0, "Campaigns"] == 2
    assert summary.loc[0, "Net_Profit"] == pytest.approx(250.0)


def test_publisher_leaderboard_ranks_by_profit_and_limits(frame):
    board = publisher_leaderboard(frame, limit=1)

    assert len(board) == 1
    assert board.index[0] == "ALPHA"
    assert board.loc["ALPHA", "Net_Profit"] == pytest.approx(750.0)


def test_conversion_funnel_returns_stages_in_order(frame):
    stages, values = conversion_funnel(frame)

    assert stages == ["Incoming Calls", "Connected", "Converted"]
    assert values == [460.0, 270.0, 65.0]


def test_segment_comparison_splits_affiliate_from_internal(frame):
    segments = segment_comparison(frame)

    assert segments["Affiliate"]["net_profit"] == pytest.approx(750.0)
    assert segments["Internal"]["net_profit"] == pytest.approx(-150.0)
    assert segments["Internal"]["roi_pct"] == pytest.approx((400 / 500 - 1) * 100)


def test_segment_comparison_handles_missing_segment(frame):
    affiliate_only = frame[frame["Is_Affiliate"]]
    segments = segment_comparison(affiliate_only)

    assert segments["Internal"]["net_profit"] == 0
    assert segments["Internal"]["roi_pct"] == 0


def test_daily_summary_provides_display_labels(frame):
    summary = daily_summary(frame)

    assert list(summary["Label"]) == ["Oct 15", "Oct 16"]


def test_apply_filters_narrows_to_the_date_range():
    from src.data import apply_filters

    df = pd.DataFrame(
        {
            "Date": pd.to_datetime(["2025-10-01", "2025-10-05", "2025-10-10"]),
            "Vertical": ["A", "A", "B"],
            "Traffic_Source": ["G", "G", "F"],
            "Spend": [100, 200, 300],
            "Is_Funded": [True, True, True],
        }
    )

    narrowed = apply_filters(df, date_range=("2025-10-04", "2025-10-08"))

    assert list(narrowed["Spend"]) == [200]


def test_weekly_summary_aggregates_days_into_weeks():
    from src.metrics import weekly_summary

    daily = pd.DataFrame(
        {
            "Date": pd.to_datetime(
                ["2025-10-06", "2025-10-07", "2025-10-13", "2025-10-14"]
            ),
            "Label": ["Oct 06", "Oct 07", "Oct 13", "Oct 14"],
            "Campaigns": [10, 10, 12, 12],
            "Spend": [100, 200, 300, 400],
            "Revenue": [150, 300, 450, 600],
            "Net_Profit": [50.0, 100.0, 150.0, 200.0],
            "ROI": [50.0, 50.0, 50.0, 50.0],
            "Incoming": [20.0, 30.0, 40.0, 50.0],
            "Connected": [10.0, 15.0, 20.0, 25.0],
            "Converted": [5.0, 5.0, 10.0, 10.0],
            "Conv_Rate": [50.0, 33.3, 50.0, 40.0],
        }
    )

    weeks = weekly_summary(daily)

    assert len(weeks) == 2
    assert weeks.loc[0, "Spend"] == 300
    assert weeks.loc[0, "Net_Profit"] == pytest.approx(150.0)
    # ROI is recomputed from the week's totals, never averaged from the daily rates.
    assert weeks.loc[0, "ROI"] == pytest.approx(50.0)
    assert weeks.loc[1, "Conv_Rate"] == pytest.approx(20 / 45 * 100, abs=0.1)


def test_weekly_summary_marks_partial_weeks():
    from src.metrics import weekly_summary

    daily = pd.DataFrame(
        {
            "Date": pd.to_datetime(["2025-10-11", "2025-10-12"]),
            "Label": ["Oct 11", "Oct 12"],
            "Campaigns": [10, 10],
            "Spend": [100, 100], "Revenue": [150, 150], "Net_Profit": [50.0, 50.0],
            "ROI": [50.0, 50.0], "Incoming": [20.0, 20.0],
            "Connected": [10.0, 10.0], "Converted": [5.0, 5.0], "Conv_Rate": [50.0, 50.0],
        }
    )

    weeks = weekly_summary(daily)

    assert bool(weeks.loc[0, "Partial"]) is True
    assert weeks.loc[0, "Label"].endswith("*")
