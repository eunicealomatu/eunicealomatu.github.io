"""Run the ServiceNow long-range plan end to end.

    python run_model.py                # all outputs into ./outputs
    python run_model.py --no-excel     # console + charts only
"""
from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

from lrp.analysis import evaluate_investment, tornado, two_way
from lrp.charts import build_all
from lrp.excel_export import write_workbook
from lrp.model import (DEFAULT_CONFIG, DEFAULT_HISTORICALS, guidance_check, load_config,
                       run_all, summary_table)

ROOT = Path(__file__).resolve().parent


def fmt_table(df: pd.DataFrame, pct_rows: set[str]) -> str:
    out = df.copy().astype(object)
    for idx in out.index:
        for col in out.columns:
            v = df.loc[idx, col]
            out.loc[idx, col] = f"{v:.1%}" if idx in pct_rows else f"{v:,.0f}"
    return out.to_markdown()


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--config", default=str(DEFAULT_CONFIG))
    ap.add_argument("--out", default=str(ROOT / "outputs"))
    ap.add_argument("--no-excel", action="store_true")
    args = ap.parse_args()

    out = Path(args.out)
    cfg = load_config(args.config)
    results = run_all(cfg)
    year = cfg.raw["sensitivity"]["target_year"]
    tor = tornado(cfg)
    grid = two_way(cfg)
    inv = evaluate_investment(cfg)
    gchk = guidance_check(cfg, results["base"])

    chart_paths = build_all(results, tor, year, out / "charts")
    charts = dict(zip(["revenue", "fcf_margin", "headcount", "tornado"], chart_paths))

    summ = summary_table(results, year)
    pct_rows = {"Subscription growth", "Non-GAAP operating margin", "FCF margin",
                "Rule of 40 (sub growth + FCF margin)"}

    lines = [
        "# Model output summary",
        "",
        f"## FY{year} by scenario ($M unless noted)",
        "",
        fmt_table(summ, pct_rows),
        "",
        "## FY2026 model vs. guidance",
        "",
        gchk.assign(
            model=[f"{v:,.0f}" if i == 0 else f"{v:.1%}" for i, v in enumerate(gchk.model)],
            guidance=[f"{v:,.0f}" if i == 0 else f"{v:.1%}" for i, v in enumerate(gchk.guidance)],
            difference=[f"{v:,.1f}" if i == 0 else f"{v * 1e4:+.0f} bp"
                        for i, v in enumerate(gchk.difference)],
        ).to_markdown(index=False),
        "",
        f"## Biggest drivers of FY{year} free cash flow (base case)",
        "",
        tor.assign(
            shock=tor.shock.map(lambda v: f"±{v * 100:.1f} pts"),
            low_case=tor.low_case.map("{:,.0f}".format),
            high_case=tor.high_case.map("{:,.0f}".format),
            swing=tor.swing.map("{:,.0f}".format),
        )[["label", "shock", "low_case", "high_case", "swing"]].to_markdown(index=False),
        "",
        f"## Investment case: {inv['name']}",
        "",
        f"- NPV including terminal value: ${inv['npv']:,.0f}M "
        f"({inv['pv_terminal'] / inv['npv']:.0%} of it from the terminal value)",
        f"- NPV of explicit five years only: ${inv['npv'] - inv['pv_terminal']:,.0f}M",
        f"- IRR on explicit cash flows: {inv['irr_no_terminal']:.1%}",
        f"- Payback: {inv['payback_years']:.1f} years vs. {inv['hurdle_payback_years']}-year hurdle",
        f"- Decision rule result: **{inv['decision']}**",
    ]
    out.mkdir(parents=True, exist_ok=True)
    (out / "summary.md").write_text("\n".join(lines) + "\n")
    print("\n".join(lines))

    if not args.no_excel:
        xlsx = write_workbook(out / "ServiceNow_LRP_Model.xlsx", cfg, results, tor, grid,
                              charts, Path(DEFAULT_HISTORICALS))
        print(f"\nWorkbook written: {xlsx}")
        print("Open it in Excel (formulas calculate on open), or run "
              "scripts/check_excel.py to recalculate and verify it with LibreOffice.")


if __name__ == "__main__":
    main()
