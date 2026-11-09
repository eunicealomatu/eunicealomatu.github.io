"""Static PNG charts for the README and the workbook."""
from __future__ import annotations

from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import pandas as pd  # noqa: E402
from matplotlib.ticker import PercentFormatter  # noqa: E402

# Colorblind-validated categorical palette (fixed order: base, upside, downside).
COLORS = {"base": "#2a78d6", "upside": "#1a9e75", "downside": "#eb6834"}
INK, MUTED, GRID, SURFACE = "#0b0b0b", "#52514e", "#e4e3df", "#fcfcfb"

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "font.size": 10,
    "axes.edgecolor": GRID,
    "axes.labelcolor": MUTED,
    "axes.titlecolor": INK,
    "axes.titlesize": 12,
    "axes.titleweight": "bold",
    "axes.titlelocation": "left",
    "xtick.color": MUTED,
    "ytick.color": MUTED,
    "axes.spines.top": False,
    "axes.spines.right": False,
    "axes.grid": True,
    "axes.grid.axis": "y",
    "grid.color": GRID,
    "grid.linewidth": 0.8,
    "figure.facecolor": SURFACE,
    "axes.facecolor": SURFACE,
    "savefig.facecolor": SURFACE,
})


def _label_end(ax, x, y, text, color):
    ax.annotate(text, (x, y), xytext=(6, 0), textcoords="offset points",
                va="center", fontsize=9, color=INK)
    ax.plot([x], [y], "o", color=color, markersize=6, markeredgecolor=SURFACE,
            markeredgewidth=1.5, zorder=4)


def scenario_lines(results: dict[str, pd.DataFrame], line: str, title: str,
                   fmt: str, path: Path, scale: float = 1.0) -> None:
    fig, ax = plt.subplots(figsize=(7.5, 4.2))
    hist_years = [2023, 2024, 2025]
    any_df = next(iter(results.values()))
    ax.plot(hist_years, any_df.loc[line, hist_years] * scale, color=MUTED, lw=2,
            label="Actual")
    for name, df in results.items():
        yrs = [2025] + [c for c in df.columns if c > 2025]
        vals = df.loc[line, yrs] * scale
        ax.plot(yrs, vals, color=COLORS[name], lw=2, label=name.capitalize(),
                ls="-" if name == "base" else "--")
        _label_end(ax, yrs[-1], vals.iloc[-1],
                   f"{name.capitalize()} {fmt.format(vals.iloc[-1])}", COLORS[name])
    ax.axvline(2025.5, color=GRID, lw=1)
    ax.text(2025.6, ax.get_ylim()[1], "Forecast →", va="top", fontsize=8, color=MUTED)
    ax.set_title(title)
    ax.set_xticks(list(any_df.columns))
    ax.set_xticklabels([f"FY{str(y)[2:]}" for y in any_df.columns])
    ax.set_xlim(2022.7, 2031.4)
    if "%" in fmt:
        ax.yaxis.set_major_formatter(PercentFormatter(1.0, decimals=0))
    ax.legend(frameon=False, loc="upper left", fontsize=9)
    fig.tight_layout()
    fig.savefig(path, dpi=160)
    plt.close(fig)


def tornado_chart(t: pd.DataFrame, path: Path, year: int) -> None:
    t = t.iloc[::-1]
    base = t["base"].iloc[0]
    fig, ax = plt.subplots(figsize=(7.5, 3.9))
    y = range(len(t))
    lo = t[["low_case", "high_case"]].min(axis=1) - base
    hi = t[["low_case", "high_case"]].max(axis=1) - base
    ax.barh(y, lo, color=COLORS["downside"], height=0.6, label="Driver moves against FCF")
    ax.barh(y, hi, color=COLORS["base"], height=0.6, label="Driver moves in favour of FCF")
    short = {"sub_growth": "Subscription growth", "wc_capture": "Working capital capture",
             "sm_pct": "Sales & marketing % rev", "cash_tax_rate": "Cash tax rate",
             "capex_pct": "Capex % rev", "rd_pct": "R&D % rev",
             "sub_gm": "Subscription gross margin", "ga_pct": "G&A % rev"}
    labels = [f"{short.get(r.driver, r.label)}  ±{r.shock * 100:.1f} pts" for r in t.itertuples()]
    ax.set_yticks(list(y))
    ax.set_yticklabels(labels, fontsize=9)
    ax.axvline(0, color=INK, lw=1)
    ax.grid(axis="x", color=GRID)
    ax.grid(axis="y", visible=False)
    ax.set_xlabel(f"Change in FY{year} free cash flow vs. base ($M)")
    fig.suptitle(f"What moves FY{year} free cash flow (base ${base:,.0f}M)",
                 x=0.02, ha="left", fontsize=12, fontweight="bold", color=INK)
    ax.legend(frameon=False, fontsize=8.5, loc="upper center", ncol=2,
              bbox_to_anchor=(0.4, 1.1))
    fig.tight_layout(rect=(0, 0, 1, 0.97))
    fig.savefig(path, dpi=160)
    plt.close(fig)


def build_all(results, tornado_df, year: int, out_dir: Path) -> list[Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    paths = {
        "revenue": out_dir / "revenue_by_scenario.png",
        "fcf_margin": out_dir / "fcf_margin_by_scenario.png",
        "headcount": out_dir / "headcount_by_scenario.png",
        "tornado": out_dir / "fcf_tornado.png",
    }
    scenario_lines(results, "total_rev", "Total revenue by scenario ($B)", "${:,.1f}B",
                   paths["revenue"], scale=0.001)
    scenario_lines(results, "fcf_margin", "Free cash flow margin by scenario", "{:.0%}",
                   paths["fcf_margin"])
    scenario_lines(results, "headcount", "Headcount by scenario (thousands)", "{:,.1f}K",
                   paths["headcount"], scale=0.001)
    tornado_chart(tornado_df, paths["tornado"], year)
    return list(paths.values())
