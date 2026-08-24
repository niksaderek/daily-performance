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
