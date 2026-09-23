"""
riaa_raw_export.csv  ->  us_music_revenue.csv

The RIAA export is one row per (format, metric, year). This collapses its 23
formats into the two bands the chart needs, keeping only the inflation-adjusted
value, and refuses to write anything unless the result reproduces the reference
table exactly.

    python prepare_data.py
"""

import collections
import csv
import io
from pathlib import Path

HERE = Path(__file__).parent

# RIAA's own format families. Synchronisation is neither physical nor digital
# and is left out of both bands — see NOTE below.
PHYSICAL = {"8 - Track", "CD", "CD Single", "Cassette", "Cassette Single",
            "DVD Audio", "LP/EP", "Music Video (Physical)", "Other Tapes",
            "SACD", "Vinyl Single"}
DIGITAL = {"Download Album", "Download Music Video", "Download Single", "Kiosk",
           "Limited Tier Paid Subscription", "On-Demand Streaming (Ad-Supported)",
           "Other Ad-Supported Streaming", "Other Digital", "Paid Subscription",
           "Paid Subscriptions", "Ringtones & Ringbacks",
           "SoundExchange Distributions"}
SYNC = {"Synchronization"}

FIRST, LAST = 1990, 2019

# Every value legible in the reference graphic, in $MM of 2019 dollars. This is
# what proves the format grouping above is RIAA's and not a guess: get one
# format into the wrong band and these stop matching.
REFERENCE = {
    1990: (14751, 0), 1991: (14705, 0), 1992: (16444, 0), 1993: (17775, 0),
    1994: (20818, 0), 1995: (20668, 0), 1996: (20423, 0), 1997: (19492, 0),
    1998: (21505, 0), 1999: (22381, 0), 2000: (21266, 0), 2001: (19836, 0),
    2002: (17926, 0), 2003: (16471, 0), 2004: (16450, 258), 2005: (14655, 1433),
    2006: (12515, 2398),
}


def band(fmt):
    if fmt in PHYSICAL:
        return "Physical"
    if fmt in DIGITAL:
        return "Digital"
    if fmt in SYNC:
        return "Sync"
    raise KeyError("format %r is in neither band — RIAA added one?" % fmt)


def aggregate(src):
    agg = collections.defaultdict(lambda: collections.defaultdict(float))
    with io.open(src, encoding="utf-8-sig") as fh:
        for row in csv.DictReader(fh):
            # "Value (Adjusted)" is RIAA's own inflation adjustment to 2019
            # dollars. "Value" is nominal, and mixing the two silently halves
            # the early years.
            if row["metric"] != "Value (Adjusted)" or not row["value_actual"]:
                continue
            agg[int(row["year"])][band(row["format"])] += float(row["value_actual"])
    return agg


def check(agg):
    bad = []
    for year, (want_p, want_d) in sorted(REFERENCE.items()):
        got_p = round(agg[year]["Physical"])
        got_d = round(agg[year]["Digital"])
        if (got_p, got_d) != (want_p, want_d):
            bad.append("%d: got (%d, %d), reference says (%d, %d)"
                       % (year, got_p, got_d, want_p, want_d))
    if bad:
        raise SystemExit("grouping does not reproduce the reference table:\n  "
                         + "\n  ".join(bad))
    return len(REFERENCE)


def main():
    agg = aggregate(HERE / "riaa_raw_export.csv")
    n_checked = check(agg)

    out = HERE / "us_music_revenue.csv"
    with io.open(out, "w", encoding="utf-8", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["Year", "Physical", "Digital", "Total", "Synchronization"])
        for year in range(FIRST, LAST + 1):
            p, d, s = (round(agg[year][k])
                       for k in ("Physical", "Digital", "Sync"))
            w.writerow([year, p, d, p + d, s])
    print("wrote %s  (%d years, %d reference values matched)"
          % (out, LAST - FIRST + 1, n_checked))


# NOTE on synchronisation. RIAA's headline total includes it, so RIAA's own
# 2019 figure is $11.1B against the $10.8B here. The reference table's Total
# column is Physical + Digital, which matches this file, and sync is zero in
# the data before 2009 — so the two only diverge at the right-hand end. It is
# carried as its own column rather than folded into either band.

if __name__ == "__main__":
    main()
