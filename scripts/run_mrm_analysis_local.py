#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Run the Boudreau-style MRM reproducible research analysis on a LOCAL workbook
(January..December sheets), adapted for your semantics where:

- "Shared code?" contains "Yes" or blank
- "Shared data?" contains "Yes" or blank
- "False Positive?" may be True/False OR Yes/blank (handled robustly)
- "Month" is a date (YYYY-MM-01) or blank (handled robustly)
- "Language(s)" and gender columns may be blank (handled robustly)

Outputs:
- Prints summary statistics
- Saves the statistics to <workbook_stem>_analysis.xlsx next to the workbook
- Saves link counts to <workbook_stem>_links_count.csv
"""

from pathlib import Path
import warnings

import numpy as np
import pandas as pd

XLSX_PATH = Path(
    "/pathway/2025/2025-MRMH-ReproducibleResearch_acceptance_2.xlsx"
)

MONTH_SHEETS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]


# -------------------------
# Helpers (robust parsing)
# -------------------------
def norm_str_series(s: pd.Series) -> pd.Series:
    """Lowercase, strip, replace NaN with empty string."""
    return s.fillna("").astype(str).str.strip().str.lower()


def yes_mask(s: pd.Series) -> pd.Series:
    """True if cell is 'yes' (case/space-insensitive)."""
    return norm_str_series(s).eq("yes")


def boolish_true_mask(s: pd.Series) -> pd.Series:
    """
    True if cell looks like True:
    - boolean True
    - 'true', 'yes', '1'
    """
    if s.dtype == bool:
        return s.fillna(False)
    v = norm_str_series(s)
    return v.isin(["true", "yes", "1"])


def boolish_false_mask(s: pd.Series) -> pd.Series:
    """
    True if cell looks like False:
    - boolean False
    - 'false', 'no', '0'
    """
    if s.dtype == bool:
        return ~s.fillna(False)
    v = norm_str_series(s)
    return v.isin(["false", "no", "0"])


def safe_month_number(month_col: pd.Series) -> pd.Series:
    """Convert Month column to month number 1..12 from date-like values."""
    dtv = pd.to_datetime(month_col, errors="coerce")
    return dtv.dt.month


# -------------------------
# Load all months into df
# -------------------------
def load_workbook_database(xlsx_path: Path) -> pd.DataFrame:
    xls = pd.ExcelFile(xlsx_path)
    dfs = []
    for m in MONTH_SHEETS:
        if m in xls.sheet_names:
            dfs.append(pd.read_excel(xls, m))
    if not dfs:
        raise ValueError("No month sheets found in workbook.")
    return pd.concat(dfs, ignore_index=True)


def main():
    if not XLSX_PATH.exists():
        raise FileNotFoundError(f"Workbook not found: {XLSX_PATH}")

    df = load_workbook_database(XLSX_PATH)

    # -------------------------
    # Cleanup / normalization
    # -------------------------
    df["Month"] = safe_month_number(df.get("Month", pd.Series(dtype=object)))

    df["Language(s)"] = norm_str_series(df.get("Language(s)", pd.Series(dtype=object)))
    df["First author gender"] = norm_str_series(df.get("First author gender", pd.Series(dtype=object)))
    df["Last author gender"] = norm_str_series(df.get("Last author gender", pd.Series(dtype=object)))
    df["Link"] = df.get("Link", pd.Series(dtype=object)).fillna("").astype(str)

    shared_code_mask = yes_mask(df.get("Shared code?", pd.Series(dtype=object)))
    shared_data_mask = yes_mask(df.get("Shared data?", pd.Series(dtype=object)))

    false_pos_mask = boolish_true_mask(df.get("False Positive?", pd.Series(dtype=object)))

    # "did actually share code/data" in the original notebook corresponds to False Positive? == False
    # Here we interpret "False" robustly; and treat blanks as "unknown" (not counted as shared).
    not_false_pos_mask = boolish_false_mask(df.get("False Positive?", pd.Series(dtype=object)))

    keywords_present = df.get("Keywords Matched", pd.Series(dtype=object)).notna()

    # -------------------------
    # Global stats (see Boudreau et al. (2022) notebook)
    # -------------------------
    papers_total = df.get("Filename", pd.Series(dtype=object)).notna().sum()
    keywords_total = int(keywords_present.sum())

    false_positives = int(false_pos_mask.sum())
    shared_codedata = int(not_false_pos_mask.sum())

    shared_code = int(shared_code_mask.sum())
    shared_data = int(shared_data_mask.sum())

    # -------------------------
    # Sanity checks
    # -------------------------
    if keywords_total != int(df.get("False Positive?", pd.Series(dtype=object)).notna().sum()):
        warnings.warn(
            "Keywords matched count does not match False Positive? filled rows. "
            "In the original workflow, these should be curated together."
        )

    # -------------------------
    # Website counts
    # -------------------------
    links_count = df["Link"].value_counts()

    github_total = int(df["Link"].str.contains("github", case=False, na=False).sum())
    gitlab_total = int(df["Link"].str.contains("gitlab", case=False, na=False).sum())
    zenodo_total = int(df["Link"].str.contains("zenodo", case=False, na=False).sum())
    osf_total = int(df["Link"].str.contains("osf", case=False, na=False).sum())

    # -------------------------
    # Language counts
    # -------------------------
    languages_count = df["Language(s)"].value_counts()

    matlab_total = int(languages_count.filter(regex="matlab").sum())
    python_total = int(languages_count.filter(regex="python").sum())
    julia_total = int(languages_count.filter(regex="julia").sum())
    cpp_total = int(languages_count.filter(regex=r"c\+\+").sum())

    # -------------------------
    # Gender counts (overall)
    # -------------------------
    first_author_gender_count = df["First author gender"].value_counts()
    last_author_gender_count = df["Last author gender"].value_counts()

    male_first = int(first_author_gender_count.get("male", 0))
    female_first = int(first_author_gender_count.get("female", 0))
    male_last = int(last_author_gender_count.get("male", 0))
    female_last = int(last_author_gender_count.get("female", 0))

    # -------------------------
    # Build statistics dict ONCE
    # -------------------------
    total_statistics = {}
    total_statistics["Total papers"] = int(papers_total)

    # -------------------------
    # Baseline gender proportions (ALL PAPERS)
    # -------------------------
    total_papers = int(len(df))

    male_first_total = int((df["First author gender"] == "male").sum())
    female_first_total = int((df["First author gender"] == "female").sum())

    male_last_total = int((df["Last author gender"] == "male").sum())
    female_last_total = int((df["Last author gender"] == "female").sum())

    total_statistics["% of papers that had male first authors"] = (male_first_total / total_papers * 100) if total_papers else 0
    total_statistics["% of papers that had female first authors"] = (female_first_total / total_papers * 100) if total_papers else 0
    total_statistics["% of papers that had male last authors"] = (male_last_total / total_papers * 100) if total_papers else 0
    total_statistics["% of papers that had female last authors"] = (female_last_total / total_papers * 100) if total_papers else 0

    # -------------------------
    # Notebook-like headline stats
    # -------------------------
    total_statistics["% papers with matched keyword"] = (keywords_total / papers_total * 100) if papers_total else 0
    total_statistics["% total papers that did actually share code/data"] = (shared_codedata / papers_total * 100) if papers_total else 0
    total_statistics["% matched papers that didn't actually share code/data"] = (false_positives / keywords_total * 100) if keywords_total else 0
    total_statistics["% matched papers that did actually share code/data"] = (shared_codedata / keywords_total * 100) if keywords_total else 0

    total_statistics["% of total papers that shared code"] = (shared_code / papers_total * 100) if papers_total else 0
    total_statistics["% of total papers that shared data"] = (shared_data / papers_total * 100) if papers_total else 0

    total_statistics["% of papers that shared code/data that hosted it on GitHub"] = (github_total / shared_codedata * 100) if shared_codedata else 0

    total_statistics["% of papers that shared code that used Python"] = (python_total / shared_code * 100) if shared_code else 0
    total_statistics["% of papers that shared code that used MATLAB"] = (matlab_total / shared_code * 100) if shared_code else 0
    total_statistics["% of papers that shared code that used C++"] = (cpp_total / shared_code * 100) if shared_code else 0
    total_statistics["% of papers that shared code that used Julia"] = (julia_total / shared_code * 100) if shared_code else 0

    # -------------------------
    # Gender among "shared code/data" papers (as in original notebook denominator)
    # Here "shared code/data" is approximated by False Positive? == False (curation)
    # -------------------------
    total_statistics["% of papers that shared code/data that had male first authors"] = (male_first / shared_codedata * 100) if shared_codedata else 0
    total_statistics["% of papers that shared code/data that had female first authors"] = (female_first / shared_codedata * 100) if shared_codedata else 0
    total_statistics["% of papers that shared code/data that had male last authors"] = (male_last / shared_codedata * 100) if shared_codedata else 0
    total_statistics["% of papers that shared code/data that had female last authors"] = (female_last / shared_codedata * 100) if shared_codedata else 0

    # -------------------------
    # Gender among code-sharing papers ONLY (Shared code? == Yes)
    # -------------------------
    df_shared_code = df[shared_code_mask]
    shared_code_total = int(len(df_shared_code))

    first_gender_code = df_shared_code["First author gender"].value_counts()
    last_gender_code = df_shared_code["Last author gender"].value_counts()

    male_first_code = int(first_gender_code.get("male", 0))
    female_first_code = int(first_gender_code.get("female", 0))
    male_last_code = int(last_gender_code.get("male", 0))
    female_last_code = int(last_gender_code.get("female", 0))

    total_statistics["% of papers that shared code that had male first authors"] = (male_first_code / shared_code_total * 100) if shared_code_total else 0
    total_statistics["% of papers that shared code that had female first authors"] = (female_first_code / shared_code_total * 100) if shared_code_total else 0
    total_statistics["% of papers that shared code that had male last authors"] = (male_last_code / shared_code_total * 100) if shared_code_total else 0
    total_statistics["% of papers that shared code that had female last authors"] = (female_last_code / shared_code_total * 100) if shared_code_total else 0

    # -------------------------
    # CONDITIONAL sharing rate by gender (what you asked for)
    # Define "shared" here as (Shared code? == Yes OR Shared data? == Yes)
    # -------------------------
    df_shared = df[shared_code_mask | shared_data_mask]

    male_first_shared = int((df_shared["First author gender"] == "male").sum())
    female_first_shared = int((df_shared["First author gender"] == "female").sum())
    male_last_shared = int((df_shared["Last author gender"] == "male").sum())
    female_last_shared = int((df_shared["Last author gender"] == "female").sum())

    total_statistics["% of male first-author papers that shared code/data"] = (
        male_first_shared / male_first_total * 100 if male_first_total else 0
    )
    total_statistics["% of female first-author papers that shared code/data"] = (
        female_first_shared / female_first_total * 100 if female_first_total else 0
    )
    total_statistics["% of male last-author papers that shared code/data"] = (
        male_last_shared / male_last_total * 100 if male_last_total else 0
    )
    total_statistics["% of female last-author papers that shared code/data"] = (
        female_last_shared / female_last_total * 100 if female_last_total else 0
    )

    # -------------------------
    # Optional chi-square test
    # -------------------------
    try:
        from scipy.stats import chi2_contingency

        contingency = [
            [male_first_shared, male_first_total - male_first_shared],
            [female_first_shared, female_first_total - female_first_shared],
        ]
        chi2, p, _, _ = chi2_contingency(contingency)
        total_statistics["Chi-square p-value (first authors, shared vs not)"] = float(p)
    except Exception:
        # If scipy isn't installed, we just skip this
        total_statistics["Chi-square p-value (first authors, shared vs not)"] = np.nan

    # -------------------------
    # Print + save
    # -------------------------
    ser = pd.Series(total_statistics)

    print("\n=== FINAL STATISTICS ===\n")
    print(ser.round(2))

    out_file = XLSX_PATH.parent / f"{XLSX_PATH.stem}_analysis.xlsx"
    ser.to_excel(out_file)
    print(f"\nSaved analysis to: {out_file}")

    links_out = XLSX_PATH.parent / f"{XLSX_PATH.stem}_links_count.csv"
    links_count.to_csv(links_out)
    print(f"Saved link counts to: {links_out}")


if __name__ == "__main__":
    main()
