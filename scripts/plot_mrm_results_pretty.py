#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt



ANALYSIS_XLSX = Path(
    "pathway/2025/2025-MRMH-ReproducibleResearch_acceptance_2_analysis.xlsx"
)

OUT_DIR = ANALYSIS_XLSX.parent / "plots_pretty"
OUT_DIR.mkdir(parents=True, exist_ok=True)

# Historical reference values (from paper / prior years)
HIST_SHARED_CODE_OR_DATA = {
    2019: 11,
    2020: 14,
    2021: 31,
}


def load_series(xlsx_path: Path) -> pd.Series:
    df = pd.read_excel(xlsx_path, header=None).dropna()
    keys = df[0].astype(str).str.strip()
    vals = pd.to_numeric(df[1], errors="coerce")
    s = pd.Series(vals.values, index=keys.values).dropna()
    return s


def get(s: pd.Series, key: str, default: float = 0.0) -> float:
    v = s.get(key, default)
    try:
        return float(v)
    except Exception:
        return float(default)


def prettify_axes(ax):
    ax.grid(axis="y", linestyle="--", linewidth=0.8, alpha=0.5)
    ax.set_axisbelow(True)
    # soften spines
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)


def annotate_bars(ax, bars, fmt="{:.1f}%"):
    ymax = 0
    for b in bars:
        h = b.get_height()
        ymax = max(ymax, h)
    pad = max(0.8, ymax * 0.02)

    for b in bars:
        h = b.get_height()
        ax.text(
            b.get_x() + b.get_width() / 2,
            h + pad,
            fmt.format(h),
            ha="center",
            va="bottom",
            fontsize=11,
        )


def save_bar(labels, values, title, ylabel, outpath: Path, ylim=(0, 100)):
    plt.figure(figsize=(8, 5))
    ax = plt.gca()
    bars = ax.bar(labels, values)
    ax.set_title(title, fontsize=14, pad=12)
    ax.set_ylabel(ylabel, fontsize=12)
    if ylim is not None:
        ax.set_ylim(*ylim)
    prettify_axes(ax)
    annotate_bars(ax, bars)
    plt.tight_layout()
    plt.savefig(outpath, dpi=220)
    plt.close()


def main():
    s = load_series(ANALYSIS_XLSX)

    # -------------------------
    # (A) Baseline gender repartition
    # -------------------------
    save_bar(
        ["Male", "Female"],
        [
            get(s, "% of papers that had male first authors"),
            get(s, "% of papers that had female first authors"),
        ],
        "Baseline: First-author gender (all papers)",
        "Percent of all papers",
        OUT_DIR / "01_baseline_gender_first_author.png",
    )

    save_bar(
        ["Male", "Female"],
        [
            get(s, "% of papers that had male last authors"),
            get(s, "% of papers that had female last authors"),
        ],
        "Baseline: Last-author gender (all papers)",
        "Percent of all papers",
        OUT_DIR / "02_baseline_gender_last_author.png",
    )

    # -------------------------
    # (B) Conditional sharing by gender
    # -------------------------
    pval = s.get("Chi-square p-value (first authors, shared vs not)", float("nan"))
    title = "Sharing (code/data) conditional on FIRST-author gender"
    if pd.notna(pval):
        try:
            title += f" (chi-square p={float(pval):.3f})"
        except Exception:
            pass

    save_bar(
        ["Male", "Female"],
        [
            get(s, "% of male first-author papers that shared code/data"),
            get(s, "% of female first-author papers that shared code/data"),
        ],
        title,
        "Percent within gender group",
        OUT_DIR / "03_conditional_sharing_first_author_gender.png",
    )

    save_bar(
        ["Male", "Female"],
        [
            get(s, "% of male last-author papers that shared code/data"),
            get(s, "% of female last-author papers that shared code/data"),
        ],
        "Sharing (code/data) conditional on LAST-author gender",
        "Percent within gender group",
        OUT_DIR / "04_conditional_sharing_last_author_gender.png",
    )

    # -------------------------
    # (C) Comparison figure: 2019–2021 vs your 2025
    # Use your "% of total papers that shared code" as 2025 proxy for "code or data"
    # -------------------------
    shared_2025_code = get(s, "% of total papers that shared code")
    shared_2025_data = get(s, "% of total papers that shared data")
    shared_2025_code_or_data = max(shared_2025_code, 0)  # data currently 0; keep logic simple

    years = list(HIST_SHARED_CODE_OR_DATA.keys()) + [2025]
    vals = [HIST_SHARED_CODE_OR_DATA[y] for y in HIST_SHARED_CODE_OR_DATA] + [shared_2025_code_or_data]

    plt.figure(figsize=(9, 5))
    ax = plt.gca()
    bars = ax.bar([str(y) for y in years], vals)

    ax.set_title("Share of papers that shared code or data (comparison)", fontsize=14, pad=12)
    ax.set_ylabel("Percent of papers", fontsize=12)
    ax.set_ylim(0, 100)
    prettify_axes(ax)
    annotate_bars(ax, bars)

    note = (
        "2019–2021 values provided by user.\n"
        "2025 value computed from your workbook as '% of total papers that shared code' "
        "(shared data currently 0)."
    )
    ax.text(
        0.5, -0.22, note,
        transform=ax.transAxes,
        ha="center", va="top",
        fontsize=10
    )

    plt.tight_layout()
    plt.savefig(OUT_DIR / "05_comparison_2019_2021_vs_2025.png", dpi=220)
    plt.close()

    # -------------------------
    # (D) Gender among code-sharing papers
    # -------------------------
    save_bar(
        ["Male first", "Female first", "Male last", "Female last"],
        [
            get(s, "% of papers that shared code that had male first authors"),
            get(s, "% of papers that shared code that had female first authors"),
            get(s, "% of papers that shared code that had male last authors"),
            get(s, "% of papers that shared code that had female last authors"),
        ],
        "Gender composition among CODE-sharing papers",
        "Percent of code-sharing papers",
        OUT_DIR / "06_gender_among_code_sharers.png",
    )

    print(f"Saved plots to: {OUT_DIR}")


if __name__ == "__main__":
    main()
