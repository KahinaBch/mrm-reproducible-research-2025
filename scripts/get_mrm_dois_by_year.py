#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Minimal script: fetch MRM DOIs for a given year via Crossref (cursor pagination).

Install:
  pip install requests pandas

Usage:
  python get_mrm_dois_by_year.py --year 2023 --out mrm_2023_dois.csv
"""

import argparse
import sys
import requests
import pandas as pd

MRM_ISSN = "1522-2594"
CROSSREF_WORKS = "https://api.crossref.org/works"


def iter_crossref(year: int, mailto: str, rows: int = 1000):
    cursor = "*"
    headers = {"User-Agent": f"mrm-doi-fetcher (mailto:{mailto})"}
    flt = f"issn:{MRM_ISSN},type:journal-article,from-pub-date:{year}-01-01,until-pub-date:{year}-12-31"

    while True:
        params = {"filter": flt, "rows": rows, "cursor": cursor}
        r = requests.get(CROSSREF_WORKS, params=params, headers=headers, timeout=45)
        if not r.ok:
            print("Crossref error:", r.status_code, r.text[:800], file=sys.stderr)
            r.raise_for_status()
        msg = r.json().get("message", {})
        items = msg.get("items", []) or []
        if not items:
            break

        for it in items:
            doi = (it.get("DOI") or "").strip()
            url = it.get("URL") or (f"https://doi.org/{doi}" if doi else "")
            title = (it.get("title") or [""])[0]
            yield {"doi": doi, "doi_url": url, "title": title}

        if len(items) < rows:
            break

        nxt = msg.get("next-cursor")
        if not nxt or nxt == cursor:
            break
        cursor = nxt


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--year", type=int, required=True)
    ap.add_argument("--out", required=True)
    args = ap.parse_args()

    rows = list(iter_crossref(args.year, args.mailto))
    df = pd.DataFrame(rows).dropna(subset=["doi"]).drop_duplicates(subset=["doi"])
    df.to_csv(args.out, index=False)
    print(f"Wrote {args.out} with {len(df)} DOIs")


if __name__ == "__main__":
    main()

