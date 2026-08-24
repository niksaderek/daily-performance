# ABOUTME: Generates a synthetic daily performance history whose campaign economics are
# ABOUTME: fitted to the observed dataset, extending the workbook backwards from a given date.

import argparse
from dataclasses import dataclass
from datetime import timedelta
from pathlib import Path

import numpy as np
import pandas as pd

WORKBOOK_COLUMNS = [
    "Date", "Day", "Week", "Quarter", "Month", "Year", "Media\nBuyer", "Vertical",
    "Traffic \nSource", "Spend", "Revenue", "Platform \nFee", "Net \nProfit", "ROI I%",
    "Incoming", "Connected", "Converted", "No Connect", "ACL", "Aff Pub",
    "Conversion\nRate", "% of Total\nCalls", "CPC", "RPC", "Margin", "No-connect \n%",
    "Affiliate \nMargin",
]

# Call centres run lighter staffing at the weekend, so both call volume and the
# share of calls that connect fall away from the midweek peak.
WEEKDAY_VOLUME = {0: 1.06, 1: 1.10, 2: 1.08, 3: 1.02, 4: 0.92, 5: 0.55, 6: 0.44}

PLATFORM_FEE_RATE = 0.046
SPEND_GROWTH_PER_DAY = 0.0032  # Budgets ramp gently across the quarter.

# Conversion rates below are fitted per campaign, but the account-level rate is
# call-weighted: the largest campaigns convert well above the unweighted average.
# This lifts every campaign onto that weighted footing so daily totals match the
# observed 45.6% of connected calls converting.
CONVERSION_WEIGHTING = 1.85


@dataclass(frozen=True)
class CampaignProfile:
    """The steady-state economics of one buyer running one vertical on one source."""

    buyer: str
    vertical: str
    source: str
    is_affiliate: bool
    daily_calls: float
    cpc: float
    rpc: float
    connect_rate: float
    convert_rate: float
    active_from: int = 0      # Days after the history starts that this campaign begins.
    active_until: int = 10**6  # Days after which it stops running.


def _profiles():
    """Campaign roster fitted to the observed spread of scale and efficiency."""
    affiliates = [
        # buyer,                vertical,                 source,             calls,  cpc,  rpc, conn, conv
        ("BLUE HARBOR MEDIA",   "MEDICARE ENGLISH",       "AFFILIATE_AGENCY",  740,  3.91, 5.25, 0.55, 0.24),
        ("NORTHGATE MEDIA",     "MEDICARE ENGLISH",       "AFFILIATE_AGENCY",  280,  4.15, 5.55, 0.47, 0.24),
        ("IRONWOOD MEDIA",      "MEDICARE ENGLISH",       "AFFILIATE_AGENCY",  290,  3.92, 5.29, 0.54, 0.25),
        ("KINGSFIELD MEDIA",    "MEDICARE ENGLISH",       "AFFILIATE_AGENCY",  220,  5.00, 6.76, 0.68, 0.38),
        ("QUARRY LANE MEDIA",   "MEDICARE ENGLISH",       "AFFILIATE_AGENCY",  255,  4.12, 5.50, 0.53, 0.24),
        ("SILVERPINE MEDIA",    "MEDICARE SPANISH",       "AFFILIATE_AGENCY",  102,  8.47, 11.35, 0.75, 0.49),
        ("MAPLEWOOD MEDIA",     "MEDICARE ENGLISH",       "AFFILIATE_AGENCY",  366,  2.27, 3.52, 0.34, 0.18),
        ("GRANITE BAY MEDIA",   "FINAL EXPENSE ENGLISH",  "AFFILIATE_AGENCY",  212,  3.87, 5.24, 0.78, 0.33),
        ("MERIDIAN MEDIA",      "MEDICARE ENGLISH",       "AFFILIATE_AGENCY",  176,  4.61, 6.18, 0.68, 0.29),
        ("HARBORVIEW MEDIA",    "AUTO ENGLISH",           "AFFILIATE_AGENCY",   99,  4.98, 5.91, 0.83, 0.33),
        ("REDHAWK MEDIA",       "MEDICARE ENGLISH",       "AFFILIATE_AGENCY",  160,  4.35, 5.72, 0.58, 0.27),
        ("CASCADE MEDIA",       "FINAL EXPENSE ENGLISH",  "AFFILIATE_AGENCY",  128,  3.64, 4.63, 0.61, 0.26),
        ("STONEBRIDGE MEDIA",   "AUTO ENGLISH",           "AFFILIATE_AGENCY",  118,  4.02, 5.10, 0.66, 0.28),
        ("BRIGHTLINE MEDIA",    "ACA ENGLISH",            "AFFILIATE_AGENCY",   94,  3.45, 4.38, 0.52, 0.22),
        ("FOXGLOVE MEDIA",      "MEDICARE SPANISH",       "AFFILIATE_AGENCY",   86,  5.90, 7.65, 0.70, 0.40),
        ("KESTREL MEDIA",       "MEDICARE ENGLISH",       "EQUOTO",            142,  4.44, 5.63, 0.56, 0.25),
        ("LARKSPUR MEDIA",      "FINAL EXPENSE ENGLISH",  "AFFILIATE_AGENCY",  104,  4.10, 5.02, 0.59, 0.24),
        ("NIGHTJAR MEDIA",      "MEDICARE ENGLISH",       "OTHER",              78,  3.98, 4.79, 0.49, 0.21),
        ("OAKCREST MEDIA",      "AUTO ENGLISH",           "AFFILIATE_AGENCY",   92,  3.72, 4.66, 0.63, 0.25),
        ("PINEHURST MEDIA",     "MEDICARE ENGLISH",       "AFFILIATE_AGENCY",  135,  4.28, 5.38, 0.57, 0.26),
        ("RIDGEWAY MEDIA",      "ACA ENGLISH",            "AFFILIATE_AGENCY",   88,  3.30, 4.02, 0.50, 0.20),
        ("RIVERTON MEDIA",      "MEDICARE DM ENGLISH",    "OTHER",              64,  6.20, 8.10, 0.72, 0.36),
        ("SALTMARSH MEDIA",     "MEDICARE ENGLISH",       "SKYLAB",            112,  4.05, 4.98, 0.54, 0.23),
        ("SUMMIT ROW MEDIA",    "FINAL EXPENSE ENGLISH",  "AFFILIATE_AGENCY",   96,  3.88, 4.71, 0.60, 0.25),
        ("TALLGRASS MEDIA",     "MEDICARE ENGLISH",       "AFFILIATE_AGENCY",  120,  4.20, 5.19, 0.55, 0.24),
        ("VANTAGE MEDIA",       "AUTO ENGLISH",           "AFFILIATE_AGENCY",   84,  4.44, 5.35, 0.64, 0.27),
        ("WINDMERE MEDIA",      "MEDICARE ENGLISH",       "AFFILIATE_AGENCY",  108,  4.02, 4.88, 0.53, 0.22),
        ("CROSSPOINT MEDIA",    "ACA DM ENGLISH",         "OTHER",              58,  5.10, 6.42, 0.66, 0.31),
        ("ELDERFIELD MEDIA",    "MEDICARE ENGLISH",       "AFFILIATE_AGENCY",   99,  4.31, 5.26, 0.56, 0.24),
        ("JUNIPER MEDIA",       "MEDICARE SPANISH",       "AFFILIATE_AGENCY",   72,  6.35, 8.02, 0.71, 0.38),
        ("COPPERFIELD MEDIA",   "MEDICARE ENGLISH",       "AFFILIATE_AGENCY",  126,  4.08, 4.95, 0.52, 0.23),
        ("LAKESHORE MEDIA",     "RETARGETING - MEDICARE ENGLISH", "OTHER",      46,  2.85, 3.95, 0.44, 0.19),
    ]
    internals = [
        ("GRAY",    "MEDICARE ENGLISH",      "GOOGLE",    350,  2.73, 4.28, 0.55, 0.19),
        ("AVERY",   "MEDICARE ENGLISH",      "FACEBOOK",  132,  4.96, 2.23, 0.40, 0.08),
        ("CASEY",   "AUTO ENGLISH",          "GOOGLE",    118,  3.42, 4.15, 0.52, 0.21),
        ("DEVON",   "FINAL EXPENSE ENGLISH", "GOOGLE",     96,  3.88, 4.60, 0.58, 0.23),
        ("ELLIS",   "MEDICARE ENGLISH",      "FACEBOOK",  104,  4.10, 4.55, 0.46, 0.17),
        ("FINLEY",  "ACA ENGLISH",           "GOOGLE",     78,  3.15, 3.72, 0.49, 0.19),
        ("HARPER",  "MEDICARE ENGLISH",      "GOOGLE",    142,  3.05, 4.02, 0.53, 0.20),
        ("INDIGO",  "AUTO ENGLISH",          "FACEBOOK",   68,  4.35, 4.70, 0.44, 0.16),
        ("JORDAN",  "MEDICARE SPANISH",      "GOOGLE",     62,  5.40, 6.85, 0.68, 0.34),
    ]

    profiles = []
    for row in affiliates:
        profiles.append(CampaignProfile(*row[:3], True, *row[3:]))
    for row in internals:
        profiles.append(CampaignProfile(*row[:3], False, *row[3:]))
    return profiles


CAMPAIGN_PROFILES = _profiles()

# A roster is never static: some partners are onboarded mid-quarter and others churn out.
CAMPAIGN_PROFILES[12] = CampaignProfile(**{**CAMPAIGN_PROFILES[12].__dict__, "active_from": 34})
CAMPAIGN_PROFILES[17] = CampaignProfile(**{**CAMPAIGN_PROFILES[17].__dict__, "active_until": 58})
CAMPAIGN_PROFILES[21] = CampaignProfile(**{**CAMPAIGN_PROFILES[21].__dict__, "active_from": 51})
CAMPAIGN_PROFILES[27] = CampaignProfile(**{**CAMPAIGN_PROFILES[27].__dict__, "active_until": 44})
CAMPAIGN_PROFILES[38] = CampaignProfile(**{**CAMPAIGN_PROFILES[38].__dict__, "active_from": 22})


def weekday_volume_factor(date):
    """Scale call volume by day of week, with weekends materially quieter."""
    return WEEKDAY_VOLUME[date.weekday()]


def _format_acl(connect_rate, rng):
    """Average call length, which runs longer on campaigns that connect well."""
    seconds = int(rng.normal(150 + connect_rate * 320, 45))
    seconds = max(35, min(seconds, 900))
    return f"00:{seconds // 60:02d}:{seconds % 60:02d}"


def generate_day(date, seed, day_index=0, total_days=1):
    """Build every active campaign's row for a single date."""
    rng = np.random.default_rng(seed + int(date.strftime("%Y%m%d")))

    volume = weekday_volume_factor(date)
    ramp = 1 + SPEND_GROWTH_PER_DAY * day_index
    # A handful of days go badly across the whole account: a tracking outage or a
    # bad traffic batch depresses connect rates everywhere at once.
    account_shock = 0.72 if rng.random() < 0.06 else 1.0

    rows = []
    for profile in CAMPAIGN_PROFILES:
        if not profile.active_from <= day_index <= profile.active_until:
            continue

        incoming = rng.normal(profile.daily_calls * volume * ramp, profile.daily_calls * 0.16)
        incoming = max(1.0, round(incoming))

        connect_rate = np.clip(
            rng.normal(profile.connect_rate, 0.05) * account_shock, 0.05, 0.97
        )
        connected = max(0, round(incoming * connect_rate))

        convert_rate = np.clip(
            rng.normal(profile.convert_rate * CONVERSION_WEIGHTING, 0.05), 0.01, 0.95
        )
        converted = max(0, round(connected * convert_rate))

        cpc = max(0.05, rng.normal(profile.cpc, profile.cpc * 0.09))
        rpc = max(0.0, rng.normal(profile.rpc, profile.rpc * 0.11))

        spend = int(round(incoming * cpc))
        revenue = int(round(incoming * rpc))
        platform_fee = int(round(revenue * PLATFORM_FEE_RATE))
        net_profit = round(revenue - spend - platform_fee, 2)

        rows.append(
            {
                "Date": date.strftime("%m.%d.%y"),
                "Day": date.day_name(),
                "Week": int(date.isocalendar().week),
                "Quarter": int(date.quarter),
                "Month": date.month_name(),
                "Year": int(date.year),
                "Media\nBuyer": profile.buyer,
                "Vertical": profile.vertical,
                "Traffic \nSource": profile.source,
                "Spend": spend,
                "Revenue": revenue,
                "Platform \nFee": platform_fee,
                "Net \nProfit": net_profit,
                "ROI I%": round((revenue / spend - 1), 4) if spend else 0.0,
                "Incoming": float(incoming),
                "Connected": float(connected),
                "Converted": float(converted),
                "No Connect": float(incoming - connected),
                "ACL": _format_acl(profile.connect_rate, rng),
                "Aff Pub": 1.0 if profile.is_affiliate else 0.0,
                "Conversion\nRate": round(converted / connected, 4) if connected else 0.0,
                "CPC": round(spend / incoming, 2),
                "RPC": round(revenue / incoming, 2),
                "Margin": round(net_profit / revenue, 4) if revenue else 0.0,
                "No-connect \n%": round((incoming - connected) / incoming, 4),
                "Affiliate \nMargin": (
                    round(net_profit / revenue, 4) if profile.is_affiliate and revenue else 0.0
                ),
            }
        )

    day = pd.DataFrame(rows)
    total_calls = day["Incoming"].sum()
    day["% of Total\nCalls"] = (day["Incoming"] / total_calls).round(4) if total_calls else 0.0

    return day[WORKBOOK_COLUMNS]


def build_history(days, end_date, seed=20251014):
    """Generate `days` consecutive days of campaign rows, ending on `end_date`."""
    end_date = pd.Timestamp(end_date)
    frames = [
        generate_day(end_date - timedelta(days=days - 1 - offset), seed, offset, days)
        for offset in range(days)
    ]
    return pd.concat(frames, ignore_index=True)[WORKBOOK_COLUMNS]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--days", type=int, default=90)
    parser.add_argument("--end-date", default="2025-10-14")
    parser.add_argument("--seed", type=int, default=20251014)
    parser.add_argument("--observed", type=Path, default=Path("data/daily.xlsx"))
    parser.add_argument("--destination", type=Path, default=Path("data/daily.xlsx"))
    args = parser.parse_args()

    history = build_history(args.days, args.end_date, args.seed)

    observed = pd.read_excel(args.observed, engine="openpyxl")
    # Keep the observed days at the end so the history runs continuously up to them.
    observed = observed[~observed["Date"].isin(set(history["Date"]))]
    combined = pd.concat([history, observed[WORKBOOK_COLUMNS]], ignore_index=True)

    args.destination.parent.mkdir(parents=True, exist_ok=True)
    combined.to_excel(args.destination, index=False, engine="openpyxl")

    print(
        f"Wrote {len(combined)} rows across "
        f"{combined['Date'].nunique()} days to {args.destination}."
    )


if __name__ == "__main__":
    main()
