# ABOUTME: Loads the daily performance workbook and normalizes its column names,
# ABOUTME: types, and derived fields into a tidy frame the dashboard can rely on.

from pathlib import Path

import pandas as pd

DATA_PATH = Path(__file__).resolve().parent.parent / "data" / "daily.xlsx"

COLUMN_NAMES = {
    "Media\nBuyer": "Media_Buyer",
    "Traffic \nSource": "Traffic_Source",
    "Platform \nFee": "Platform_Fee",
    "Net \nProfit": "Net_Profit",
    "ROI I%": "ROI_Pct",
    "Conversion\nRate": "Conversion_Rate",
    "% of Total\nCalls": "Pct_Total_Calls",
    "No-connect \n%": "No_Connect_Pct",
    "Affiliate \nMargin": "Affiliate_Margin",
    "No Connect": "No_Connect",
    "Aff Pub": "Aff_Pub",
}

# Rows below this spend are smoke tests rather than funded campaigns.
TEST_CAMPAIGN_SPEND_FLOOR = 50


def load_performance_data(path=DATA_PATH):
    """Read the workbook and return a frame with normalized columns and real dates."""
    df = pd.read_excel(path, engine="openpyxl")
    df = df.rename(columns=COLUMN_NAMES)

    # Source dates arrive as MM.DD.YY strings, which sort incorrectly as text.
    df["Date"] = pd.to_datetime(df["Date"], format="%m.%d.%y")
    df["Is_Affiliate"] = df["Aff_Pub"].fillna(0).astype(bool)
    df["Is_Funded"] = df["Spend"] > TEST_CAMPAIGN_SPEND_FLOOR

    return df.sort_values("Date").reset_index(drop=True)


def apply_filters(df, dates=None, verticals=None, sources=None, funded_only=True):
    """Narrow the frame to the selections made in the sidebar."""
    filtered = df
    if dates:
        filtered = filtered[filtered["Date"].isin(dates)]
    if verticals:
        filtered = filtered[filtered["Vertical"].isin(verticals)]
    if sources:
        filtered = filtered[filtered["Traffic_Source"].isin(sources)]
    if funded_only:
        filtered = filtered[filtered["Is_Funded"]]
    return filtered
