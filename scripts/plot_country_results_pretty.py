#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt


COUNTRY_ANALYSIS_XLSX = Path(
    "/pathway/2025/2025-MRMH-ReproducibleResearch_acceptance_2_country_analysis.xlsx"
)

OUT_DIR = COUNTRY_ANALYSIS_XLSX.parent / "country_plots"
OUT_DIR.mkdir(parents=True, exist_ok=True)

# Plot parameters (adjust as you like)
MIN_PAPERS_FOR_RATE_PLOTS = 5   # only show countries with >= this many papers
TOP_N_COUNTRIES = 20            # number of countries shown in top bars


def prettify_axes(ax):
    ax.grid(axis="y", linestyle="--", linewidth=0.8, alpha=0.5)
    ax.set_axisbelow(True)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)


def annotate_bars(ax, bars, fmt="{:.0f}", y_pad_frac=0.02):
    heights = [b.get_height() for b in bars]
    ymax = max(heights) if heights else 0
    pad = max(0.8, ymax * y_pad_frac)

    for b in bars:
        h = b.get_height()
        ax.text(
            b.get_x() + b.get_width() / 2,
            h + pad,
            fmt.format(h),
            ha="center",
            va="bottom",
            fontsize=10,
        )


def load_country_table(xlsx_path: Path) -> pd.DataFrame:
    """
    Expects the Excel created by run_country_analysis.py:
    index: Country
    columns: Shared_count, Total_count, Sharing_rate_%, Proportion_of_all_papers_%
    """
    df = pd.read_excel(xlsx_path, index_col=0)
    df.index = df.index.astype(str).str.strip()
    # Ensure numeric
    for col in ["Shared_count", "Total_count", "Sharing_rate_%", "Proportion_of_all_papers_%"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    df = df.dropna(subset=["Total_count"])
    return df


def plot_top_countries_by_count(df: pd.DataFrame):
    top = df.sort_values("Total_count", ascending=False).head(TOP_N_COUNTRIES)

    plt.figure(figsize=(10, 6))
    ax = plt.gca()
    bars = ax.bar(top.index.tolist(), top["Total_count"].tolist())
    ax.set_title(f"Top {TOP_N_COUNTRIES} countries by number of papers", fontsize=14, pad=12)
    ax.set_ylabel("Number of papers", fontsize=12)
    plt.xticks(rotation=60, ha="right")
    prettify_axes(ax)
    annotate_bars(ax, bars, fmt="{:.0f}")
    plt.tight_layout()
    plt.savefig(OUT_DIR / "01_top_countries_by_paper_count.png", dpi=220)
    plt.close()


def plot_sharing_rate_by_country(df: pd.DataFrame):
    sub = df[df["Total_count"] >= MIN_PAPERS_FOR_RATE_PLOTS].copy()
    sub = sub.sort_values("Sharing_rate_%", ascending=False)

    # show at most TOP_N_COUNTRIES (highest rates)
    sub = sub.head(TOP_N_COUNTRIES)

    plt.figure(figsize=(10, 6))
    ax = plt.gca()
    bars = ax.bar(sub.index.tolist(), sub["Sharing_rate_%"].tolist())
    ax.set_title(
        f"Sharing rate by country (countries with ≥ {MIN_PAPERS_FOR_RATE_PLOTS} papers)\nTop {min(TOP_N_COUNTRIES, len(sub))} by sharing rate",
        fontsize=14,
        pad=12,
    )
    ax.set_ylabel("Sharing rate (% of papers in country)", fontsize=12)
    ax.set_ylim(0, 100)
    plt.xticks(rotation=60, ha="right")
    prettify_axes(ax)
    annotate_bars(ax, bars, fmt="{:.1f}%")
    plt.tight_layout()
    plt.savefig(OUT_DIR / "02_sharing_rate_by_country_top.png", dpi=220)
    plt.close()


def plot_count_vs_sharing_scatter(df: pd.DataFrame):
    sub = df[df["Total_count"] >= MIN_PAPERS_FOR_RATE_PLOTS].copy()

    plt.figure(figsize=(8, 6))
    ax = plt.gca()
    ax.scatter(sub["Total_count"], sub["Sharing_rate_%"])
    ax.set_title(
        f"Country paper volume vs sharing rate (countries with ≥ {MIN_PAPERS_FOR_RATE_PLOTS} papers)",
        fontsize=14,
        pad=12,
    )
    ax.set_xlabel("Number of papers", fontsize=12)
    ax.set_ylabel("Sharing rate (%)", fontsize=12)
    ax.set_ylim(0, 100)
    prettify_axes(ax)

    # label a few biggest countries 
    biggest = sub.sort_values("Total_count", ascending=False).head(10)
    for country, row in biggest.iterrows():
        ax.text(row["Total_count"], row["Sharing_rate_%"], str(country), fontsize=9)

    plt.tight_layout()
    plt.savefig(OUT_DIR / "03_count_vs_sharing_rate_scatter.png", dpi=220)
    plt.close()


def plot_top_countries_by_share_count(df: pd.DataFrame):
    sub = df[df["Total_count"] >= MIN_PAPERS_FOR_RATE_PLOTS].copy()
    sub = sub.sort_values("Shared_count", ascending=False).head(TOP_N_COUNTRIES)

    plt.figure(figsize=(10, 6))
    ax = plt.gca()
    bars = ax.bar(sub.index.tolist(), sub["Shared_count"].tolist())
    ax.set_title(
        f"Top {min(TOP_N_COUNTRIES, len(sub))} countries by number of sharing papers\n(countries with ≥ {MIN_PAPERS_FOR_RATE_PLOTS} papers)",
        fontsize=14,
        pad=12,
    )
    ax.set_ylabel("Number of papers that shared code/data", fontsize=12)
    plt.xticks(rotation=60, ha="right")
    prettify_axes(ax)
    annotate_bars(ax, bars, fmt="{:.0f}")
    plt.tight_layout()
    plt.savefig(OUT_DIR / "04_top_countries_by_sharing_count.png", dpi=220)
    plt.close()


def main():
    if not COUNTRY_ANALYSIS_XLSX.exists():
        raise FileNotFoundError(f"Not found: {COUNTRY_ANALYSIS_XLSX}")

    df = load_country_table(COUNTRY_ANALYSIS_XLSX)

    # Basic sanity
    required = {"Shared_count", "Total_count", "Sharing_rate_%"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Missing columns in analysis file: {missing}")

    plot_top_countries_by_count(df)
    plot_sharing_rate_by_country(df)
    plot_count_vs_sharing_scatter(df)
    plot_top_countries_by_share_count(df)

    print(f"Saved country plots to: {OUT_DIR}")


if __name__ == "__main__":
    main()
