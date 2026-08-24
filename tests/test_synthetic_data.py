# ABOUTME: Tests the synthetic history generator, covering schema fidelity, internal
# ABOUTME: arithmetic consistency, determinism, and the weekly volume pattern it imposes.

import pandas as pd
import pytest

from scripts.generate_history import (
    CAMPAIGN_PROFILES,
    build_history,
    generate_day,
    weekday_volume_factor,
)

EXPECTED_COLUMNS = [
    "Date", "Day", "Week", "Quarter", "Month", "Year", "Media\nBuyer", "Vertical",
    "Traffic \nSource", "Spend", "Revenue", "Platform \nFee", "Net \nProfit", "ROI I%",
    "Incoming", "Connected", "Converted", "No Connect", "ACL", "Aff Pub",
    "Conversion\nRate", "% of Total\nCalls", "CPC", "RPC", "Margin", "No-connect \n%",
    "Affiliate \nMargin",
]


@pytest.fixture(scope="module")
def history():
    return build_history(days=30, end_date=pd.Timestamp("2025-10-14"), seed=7)


def test_history_matches_the_workbook_schema(history):
    assert list(history.columns) == EXPECTED_COLUMNS


def test_history_covers_every_requested_day(history):
    dates = sorted(history["Date"].unique())
    assert len(dates) == 30
    assert dates[-1] == "10.14.25"


def test_generation_is_deterministic_for_a_seed():
    first = build_history(days=10, end_date=pd.Timestamp("2025-10-14"), seed=42)
    second = build_history(days=10, end_date=pd.Timestamp("2025-10-14"), seed=42)

    pd.testing.assert_frame_equal(first, second)


def test_different_seeds_produce_different_numbers():
    first = build_history(days=10, end_date=pd.Timestamp("2025-10-14"), seed=1)
    second = build_history(days=10, end_date=pd.Timestamp("2025-10-14"), seed=2)

    assert first["Spend"].sum() != second["Spend"].sum()


def test_funnel_counts_never_exceed_the_stage_above(history):
    assert (history["Connected"] <= history["Incoming"]).all()
    assert (history["Converted"] <= history["Connected"]).all()
    assert (history["No Connect"] >= 0).all()


def test_no_connect_complements_connected(history):
    reconstructed = history["Connected"] + history["No Connect"]
    assert (reconstructed == history["Incoming"]).all()


def test_profit_reconciles_with_revenue_spend_and_fee(history):
    expected = history["Revenue"] - history["Spend"] - history["Platform \nFee"]
    assert (expected - history["Net \nProfit"]).abs().max() < 0.01


def test_derived_rates_agree_with_their_components(history):
    row = history.iloc[0]
    assert row["CPC"] == pytest.approx(row["Spend"] / row["Incoming"], abs=0.01)
    assert row["RPC"] == pytest.approx(row["Revenue"] / row["Incoming"], abs=0.01)
    assert row["Conversion\nRate"] == pytest.approx(
        row["Converted"] / row["Connected"], abs=0.001
    )


def test_percent_of_total_calls_sums_to_one_within_a_day(history):
    for _, day in history.groupby("Date"):
        assert day["% of Total\nCalls"].sum() == pytest.approx(1.0, abs=0.01)


def test_calendar_fields_agree_with_the_date(history):
    row = history.iloc[0]
    stamp = pd.Timestamp(f"20{row['Date'][6:8]}-{row['Date'][0:2]}-{row['Date'][3:5]}")

    assert row["Day"] == stamp.day_name()
    assert row["Year"] == stamp.year
    assert row["Month"] == stamp.month_name()
    assert row["Quarter"] == stamp.quarter


def test_weekends_carry_lighter_call_volume():
    saturday = weekday_volume_factor(pd.Timestamp("2025-10-11"))
    wednesday = weekday_volume_factor(pd.Timestamp("2025-10-08"))

    assert saturday < wednesday


def test_a_generated_day_only_contains_active_campaigns():
    day = generate_day(pd.Timestamp("2025-09-10"), seed=3)
    names = set(day["Media\nBuyer"])

    assert names.issubset({profile.buyer for profile in CAMPAIGN_PROFILES})
    assert len(day) > 0


def test_affiliate_flag_is_consistent_per_buyer(history):
    flags = history.groupby("Media\nBuyer")["Aff Pub"].nunique()
    assert (flags == 1).all()
