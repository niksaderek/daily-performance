# ABOUTME: Maps confidential partner and buyer identities in a raw daily performance
# ABOUTME: export onto stable pseudonyms, writing a shareable workbook to data/daily.xlsx.

import argparse
import hashlib
from pathlib import Path

import pandas as pd

BUYER_COLUMN = "Media\nBuyer"
AFFILIATE_COLUMN = "Aff Pub"

AGENCY_HEADS = [
    "Northgate", "Silverpine", "Redhawk", "Blue Harbor", "Ironwood",
    "Meridian", "Copperfield", "Lakeshore", "Vantage", "Stonebridge",
    "Brightline", "Cascade", "Foxglove", "Kestrel", "Ridgeway",
    "Summit Row", "Tallgrass", "Windmere", "Crosspoint", "Elderfield",
    "Granite Bay", "Harborview", "Juniper", "Kingsfield", "Larkspur",
    "Maplewood", "Nightjar", "Oakcrest", "Pinehurst", "Quarry Lane",
    "Riverton", "Saltmarsh", "Thornfield", "Umberland",
]
AGENCY_TAILS = ["Media", "Digital", "Partners", "Group", "Labs", "Collective"]

INTERNAL_NAMES = [
    "Avery", "Casey", "Devon", "Ellis", "Finley", "Gray", "Harper",
    "Indigo", "Jordan", "Kai", "Lane", "Morgan", "Noor", "Onyx",
    "Parker", "Quinn", "Reese", "Sage", "Tatum", "Vale",
]

PRESERVED_LABELS = {"TESTING"}


def _stable_order(names):
    """Order names by a hash of the name so the mapping never depends on row order."""
    return sorted(names, key=lambda n: hashlib.sha256(n.encode("utf-8")).hexdigest())


def build_buyer_map(df):
    """Assign each real buyer a pseudonym matching its affiliate/internal character."""
    affiliate_share = df.groupby(BUYER_COLUMN)[AFFILIATE_COLUMN].mean()

    affiliates = _stable_order(
        [n for n, share in affiliate_share.items()
         if share >= 0.5 and n not in PRESERVED_LABELS]
    )
    internals = _stable_order(
        [n for n, share in affiliate_share.items()
         if share < 0.5 and n not in PRESERVED_LABELS]
    )

    mapping = {label: label for label in PRESERVED_LABELS}

    for i, real in enumerate(affiliates):
        head = AGENCY_HEADS[i % len(AGENCY_HEADS)]
        tail = AGENCY_TAILS[(i // len(AGENCY_HEADS)) % len(AGENCY_TAILS)]
        mapping[real] = f"{head} {tail}".upper()

    for i, real in enumerate(internals):
        suffix = "" if i < len(INTERNAL_NAMES) else f" {i // len(INTERNAL_NAMES) + 1}"
        mapping[real] = f"{INTERNAL_NAMES[i % len(INTERNAL_NAMES)]}{suffix}".upper()

    return mapping


def anonymize(source, destination):
    df = pd.read_excel(source, engine="openpyxl")
    mapping = build_buyer_map(df)
    df[BUYER_COLUMN] = df[BUYER_COLUMN].map(lambda n: mapping.get(n, n))

    destination.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(destination, index=False, engine="openpyxl")
    return mapping, len(df)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path, help="Raw export containing real partner names")
    parser.add_argument("destination", type=Path, nargs="?", default=Path("data/daily.xlsx"))
    args = parser.parse_args()

    mapping, rows = anonymize(args.source, args.destination)
    identities = len([k for k, v in mapping.items() if k != v])
    print(f"Wrote {rows} rows to {args.destination} with {identities} identities replaced.")


if __name__ == "__main__":
    main()
