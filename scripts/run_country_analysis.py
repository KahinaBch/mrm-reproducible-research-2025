#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Country-based analysis of reproducible research sharing in MRM.

Computes:
1) Proportion of papers by country
2) Proportion of papers sharing code/data within each country
3) Chi-square test of independence (Country x Sharing)
4) Cramér's V effect size

Outputs:
- <workbook>_country_analysis.xlsx
- <workbook>_country_contingency.csv
"""

from pathlib import Path
import numpy as np
import pandas as pd
from scipy.stats import chi2_contingency

XLSX_PATH = Path(
    "/pathway/2025/2025-MRMH-ReproducibleResearch_acceptance_2.xlsx"
)

MONTH_SHEETS = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December","Sheet7"
]


def load_workbook_database(xlsx_path: Path) -> pd.DataFrame:
    xls = pd.ExcelFile(xlsx_path)
    dfs = []
    for m in MONTH_SHEETS:
        if m in xls.sheet_names:
            dfs.append(pd.read_excel(xls, m))
    return pd.concat(dfs, ignore_index=True)


def main():

    df = load_workbook_database(XLSX_PATH)

    # Normalize columns
    df["Country"] = df["First author affiliation country"].fillna("").astype(str).str.strip()
    df["Shared_code"] = df["Shared code?"].fillna("").astype(str).str.strip().str.lower() == "yes"
    df["Shared_data"] = df["Shared data?"].fillna("").astype(str).str.strip().str.lower() == "yes"

    # Define sharing = code OR data
    df["Shared"] = df["Shared_code"] | df["Shared_data"]

    # Remove rows without country
    df = df[df["Country"] != ""]

    total_papers = len(df)

    # ----------------------------------------------------
    # 1) Baseline proportion of papers by country
    # ----------------------------------------------------
    country_counts = df["Country"].value_counts()
    country_prop = country_counts / total_papers * 100

    # ----------------------------------------------------
    # 2) Conditional sharing probability by country
    # ----------------------------------------------------
    sharing_by_country = (
        df.groupby("Country")["Shared"]
        .agg(["sum", "count"])
        .rename(columns={"sum": "Shared_count", "count": "Total_count"})
    )

    sharing_by_country["Sharing_rate_%"] = (
        sharing_by_country["Shared_count"] /
        sharing_by_country["Total_count"] * 100
    )

    # Merge baseline proportion
    sharing_by_country["Proportion_of_all_papers_%"] = (
        country_prop
    )

    sharing_by_country = sharing_by_country.sort_values(
        "Total_count", ascending=False
    )

    # ----------------------------------------------------
    # 3) Statistical test (Country x Sharing)
    # ----------------------------------------------------
    contingency_table = pd.crosstab(df["Country"], df["Shared"])

    chi2, p, dof, expected = chi2_contingency(contingency_table)

    # Cramér’s V
    n = contingency_table.sum().sum()
    phi2 = chi2 / n
    r, k = contingency_table.shape
    cramers_v = np.sqrt(phi2 / min((k - 1), (r - 1)))

    # ----------------------------------------------------
    # Save outputs
    # ----------------------------------------------------
    output_analysis = XLSX_PATH.parent / f"{XLSX_PATH.stem}_country_analysis.xlsx"
    output_contingency = XLSX_PATH.parent / f"{XLSX_PATH.stem}_country_contingency.csv"

    sharing_by_country.to_excel(output_analysis)
    contingency_table.to_csv(output_contingency)

    # ----------------------------------------------------
    # Print summary
    # ----------------------------------------------------
    print("\n=== COUNTRY ANALYSIS ===\n")
    print("Total papers analysed:", total_papers)
    print("\nChi-square test:")
    print("  chi2 =", round(chi2, 3))
    print("  p-value =", round(p, 5))
    print("  dof =", dof)
    print("  Cramér's V =", round(cramers_v, 3))

    if p < 0.05:
        print("\nResult: Sharing is significantly associated with country.")
    else:
        print("\nResult: No statistically significant association between country and sharing.")

    print("\nSaved:")
    print("  -", output_analysis)
    print("  -", output_contingency)


if __name__ == "__main__":
    main()
