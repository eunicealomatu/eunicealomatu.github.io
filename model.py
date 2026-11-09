"""Driver-based long-range plan (LRP) engine for ServiceNow.

The forecast is deliberately simple and transparent: every line is either a reported
historical figure or one driver applied to another line. The same logic is written into
the Excel workbook as live formulas (see excel_export.py), and a check script confirms the
two give identical answers.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pandas as pd
import yaml

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_CONFIG = ROOT / "config" / "assumptions.yaml"
DEFAULT_HISTORICALS = ROOT / "data" / "historicals.csv"

HIST_YEARS = [2023, 2024, 2025]

# Order of lines in the output table (and in the Excel LRP sheet).
LINES = [
    ("sub_rev", "Subscription revenue"),
    ("ps_rev", "Professional services & other revenue"),
    ("total_rev", "Total revenue"),
    ("sub_growth", "  Subscription growth"),
    ("sub_cogs", "Cost of subscription revenue"),
    ("ps_cogs", "Cost of professional services revenue"),
    ("gross_profit", "Gross profit (GAAP)"),
    ("gross_margin", "  Gross margin"),
    ("sm", "Sales & marketing"),
    ("rd", "Research & development"),
    ("ga", "General & administrative"),
    ("total_opex", "Total operating expenses"),
    ("gaap_oi", "Operating income (GAAP)"),
    ("gaap_om", "  GAAP operating margin"),
    ("sbc", "Stock-based compensation"),
    ("other_adj", "Amortization & acquisition-related costs"),
    ("ng_oi", "Operating income (non-GAAP)"),
    ("ng_om", "  Non-GAAP operating margin"),
    ("cash_taxes", "Cash taxes"),
    ("wc_inflow", "Working capital inflow (mainly deferred revenue)"),
    ("net_interest", "Net interest & other cash income"),
    ("ocf", "Operating cash flow"),
    ("capex", "Capital expenditures"),
    ("fcf", "Free cash flow"),
    ("fcf_margin", "  FCF margin"),
    ("headcount", "Headcount (year end)"),
    ("rev_per_emp", "Revenue per employee ($K)"),
    ("crpo", "cRPO (year end)"),
    ("rule_of_40", "Rule of 40 (sub growth + FCF margin)"),
]
PCT_LINES = {"sub_growth", "gross_margin", "gaap_om", "ng_om", "fcf_margin", "rule_of_40"}


@dataclass
class Config:
    raw: dict

    @property
    def years(self) -> list[int]:
        return list(self.raw["forecast_years"])

    @property
    def scenarios(self) -> list[str]:
        return list(self.raw["scenarios"].keys())

    def drivers(self, scenario: str) -> dict[str, list[float]]:
        return {k: list(v) for k, v in self.raw["scenarios"][scenario]["drivers"].items()}


def load_config(path: Path | str = DEFAULT_CONFIG) -> Config:
    with open(path) as f:
        cfg = Config(yaml.safe_load(f))
    _validate(cfg)
    return cfg


def _validate(cfg: Config) -> None:
    n = len(cfg.years)
    expected = set(cfg.raw["drivers"].keys())
    for name in cfg.scenarios:
        d = cfg.drivers(name)
        missing = expected - set(d)
        if missing:
            raise ValueError(f"Scenario '{name}' is missing drivers: {sorted(missing)}")
        for key, vals in d.items():
            if len(vals) != n:
                raise ValueError(f"{name}.{key} has {len(vals)} values; expected {n}")


def load_historicals(path: Path | str = DEFAULT_HISTORICALS) -> pd.DataFrame:
    """Reported figures plus a few derived lines, indexed by metric, columns = years."""
    raw = pd.read_csv(path).set_index("metric")
    h = raw[[f"FY{y}" for y in HIST_YEARS]].astype(float)
    h.columns = HIST_YEARS

    out = pd.DataFrame(index=[k for k, _ in LINES], columns=HIST_YEARS, dtype=float)
    for k in ["sub_rev", "ps_rev", "total_rev", "sub_cogs", "ps_cogs", "sm", "rd", "ga",
              "gaap_oi", "sbc", "ocf", "capex", "fcf", "headcount", "crpo"]:
        out.loc[k] = h.loc[k]
    out.loc["gross_profit"] = out.loc["total_rev"] - out.loc["sub_cogs"] - out.loc["ps_cogs"]
    out.loc["gross_margin"] = out.loc["gross_profit"] / out.loc["total_rev"]
    out.loc["total_opex"] = out.loc["sm"] + out.loc["rd"] + out.loc["ga"]
    out.loc["gaap_om"] = out.loc["gaap_oi"] / out.loc["total_rev"]
    # Non-GAAP operating income is reported only as a rounded margin, so the
    # amortization/acquisition line is backed out as a plug (approximate).
    out.loc["ng_oi"] = h.loc["ng_om_reported"] * out.loc["total_rev"]
    out.loc["ng_om"] = h.loc["ng_om_reported"]
    out.loc["other_adj"] = out.loc["ng_oi"] - out.loc["gaap_oi"] - out.loc["sbc"]
    out.loc["fcf_margin"] = out.loc["fcf"] / out.loc["total_rev"]
    out.loc["rev_per_emp"] = out.loc["total_rev"] * 1000 / out.loc["headcount"]
    growth = out.loc["sub_rev"] / out.loc["sub_rev"].shift(1) - 1
    out.loc["sub_growth"] = growth
    out.loc["rule_of_40"] = growth + out.loc["fcf_margin"]
    return out


def project(cfg: Config, scenario: str, hist: pd.DataFrame | None = None,
            overrides: dict[str, list[float]] | None = None) -> pd.DataFrame:
    """Return historicals + forecast for one scenario as a DataFrame (lines x years)."""
    hist = load_historicals() if hist is None else hist
    d = cfg.drivers(scenario)
    if overrides:
        d.update(overrides)

    years = cfg.years
    df = pd.DataFrame(index=hist.index, columns=HIST_YEARS + years, dtype=float)
    df[HIST_YEARS] = hist[HIST_YEARS]

    for i, y in enumerate(years):
        p = df[y - 1]
        g = {k: v[i] for k, v in d.items()}
        c = {}
        c["sub_rev"] = p["sub_rev"] * (1 + g["sub_growth"])
        c["ps_rev"] = p["ps_rev"] * (1 + g["ps_growth"])
        c["total_rev"] = c["sub_rev"] + c["ps_rev"]
        c["sub_growth"] = g["sub_growth"]
        c["sub_cogs"] = c["sub_rev"] * (1 - g["sub_gm"])
        c["ps_cogs"] = c["ps_rev"] * (1 - g["ps_gm"])
        c["gross_profit"] = c["total_rev"] - c["sub_cogs"] - c["ps_cogs"]
        c["gross_margin"] = c["gross_profit"] / c["total_rev"]
        c["sm"] = c["total_rev"] * g["sm_pct"]
        c["rd"] = c["total_rev"] * g["rd_pct"]
        c["ga"] = c["total_rev"] * g["ga_pct"]
        c["total_opex"] = c["sm"] + c["rd"] + c["ga"]
        c["gaap_oi"] = c["gross_profit"] - c["total_opex"]
        c["gaap_om"] = c["gaap_oi"] / c["total_rev"]
        c["sbc"] = c["total_rev"] * g["sbc_pct"]
        c["other_adj"] = c["total_rev"] * g["other_adj_pct"]
        c["ng_oi"] = c["gaap_oi"] + c["sbc"] + c["other_adj"]
        c["ng_om"] = c["ng_oi"] / c["total_rev"]
        c["cash_taxes"] = -c["ng_oi"] * g["cash_tax_rate"]
        c["wc_inflow"] = (c["total_rev"] - p["total_rev"]) * g["wc_capture"]
        c["net_interest"] = c["total_rev"] * g["net_interest_pct"]
        c["ocf"] = c["ng_oi"] + c["cash_taxes"] + c["wc_inflow"] + c["net_interest"]
        c["capex"] = c["total_rev"] * g["capex_pct"]
        c["fcf"] = c["ocf"] - c["capex"]
        c["fcf_margin"] = c["fcf"] / c["total_rev"]
        c["rev_per_emp"] = p["rev_per_emp"] * (1 + g["rev_per_emp_growth"])
        c["headcount"] = c["total_rev"] * 1000 / c["rev_per_emp"]
        c["crpo"] = c["sub_rev"] * g["crpo_ratio"]
        c["rule_of_40"] = c["sub_growth"] + c["fcf_margin"]
        df[y] = pd.Series(c)
    return df


def run_all(cfg: Config) -> dict[str, pd.DataFrame]:
    hist = load_historicals()
    return {s: project(cfg, s, hist) for s in cfg.scenarios}


def guidance_check(cfg: Config, base: pd.DataFrame) -> pd.DataFrame:
    """Compare the model's FY2026 with company guidance."""
    gd = cfg.raw["guidance_2026"]
    y = 2026
    sub_gm = 1 - base.loc["sub_cogs", y] / base.loc["sub_rev", y]
    rows = [
        ("Subscription revenue ($M)", base.loc["sub_rev", y],
         (gd["sub_rev_low"] + gd["sub_rev_high"]) / 2),
        ("Subscription gross margin (GAAP)", sub_gm, gd["sub_gm_gaap"]),
        ("GAAP operating margin", base.loc["gaap_om", y], gd["gaap_om"]),
        ("Non-GAAP operating margin", base.loc["ng_om", y], gd["ng_om"]),
        ("Free cash flow margin", base.loc["fcf_margin", y], gd["fcf_margin"]),
    ]
    out = pd.DataFrame(rows, columns=["metric", "model", "guidance"])
    out["difference"] = out["model"] - out["guidance"]
    return out


def summary_table(results: dict[str, pd.DataFrame], year: int) -> pd.DataFrame:
    keys = ["total_rev", "sub_growth", "ng_om", "fcf", "fcf_margin", "headcount", "rule_of_40"]
    labels = dict(LINES)
    data = {s: df.loc[keys, year] for s, df in results.items()}
    t = pd.DataFrame(data)
    t.index = [labels[k].strip() for k in keys]
    return t
