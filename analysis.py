"""Sensitivity analysis and investment-proposal evaluation."""
from __future__ import annotations

import numpy as np
import pandas as pd

from .model import Config, load_historicals, project


def _shift(cfg: Config, scenario: str, driver: str, delta: float) -> list[float]:
    """Shift a driver by `delta` in every forecast year after FY2026 (guidance year)."""
    vals = cfg.drivers(scenario)[driver]
    return [v if i == 0 else v + delta for i, v in enumerate(vals)]


def tornado(cfg: Config, scenario: str = "base", metric: str = "fcf") -> pd.DataFrame:
    """One-at-a-time sensitivity of a target-year metric to each driver."""
    s = cfg.raw["sensitivity"]
    year = s["target_year"]
    hist = load_historicals()
    base_val = project(cfg, scenario, hist).loc[metric, year]
    labels = {k: v["label"] for k, v in cfg.raw["drivers"].items()}
    rows = []
    for driver, shock in s["shocks"].items():
        lo = project(cfg, scenario, hist, {driver: _shift(cfg, scenario, driver, -shock)})
        hi = project(cfg, scenario, hist, {driver: _shift(cfg, scenario, driver, shock)})
        rows.append({
            "driver": driver,
            "label": labels[driver],
            "shock": shock,
            "low_case": lo.loc[metric, year],
            "high_case": hi.loc[metric, year],
        })
    t = pd.DataFrame(rows)
    t["base"] = base_val
    t["swing"] = (t["high_case"] - t["low_case"]).abs()
    return t.sort_values("swing", ascending=False).reset_index(drop=True)


def two_way(cfg: Config, scenario: str = "base", metric: str = "fcf") -> pd.DataFrame:
    """Grid of a target-year metric across two drivers."""
    s = cfg.raw["sensitivity"]
    year = s["target_year"]
    r, c = s["two_way"]["rows"], s["two_way"]["cols"]
    hist = load_historicals()
    grid = pd.DataFrame(index=r["steps"], columns=c["steps"], dtype=float)
    for dr in r["steps"]:
        for dc in c["steps"]:
            df = project(cfg, scenario, hist, {
                r["driver"]: _shift(cfg, scenario, r["driver"], dr),
                c["driver"]: _shift(cfg, scenario, c["driver"], dc),
            })
            grid.loc[dr, dc] = df.loc[metric, year]
    grid.index.name = f"{r['driver']} shift"
    grid.columns.name = f"{c['driver']} shift"
    return grid


def irr(cash_flows: list[float]) -> float:
    """Internal rate of return by bisection (no numpy-financial dependency)."""
    def npv_at(rate):
        return sum(cf / (1 + rate) ** t for t, cf in enumerate(cash_flows))
    lo, hi = -0.99, 10.0
    if npv_at(lo) * npv_at(hi) > 0:
        return float("nan")
    for _ in range(200):
        mid = (lo + hi) / 2
        if npv_at(lo) * npv_at(mid) <= 0:
            hi = mid
        else:
            lo = mid
    return (lo + hi) / 2


def evaluate_investment(cfg: Config) -> dict:
    """NPV, IRR and payback for the investment case in the config.

    Cash flows are end-of-year; year 1 = first year of the proposal. NPV matches Excel's
    NPV(rate, flows) + terminal value discounted from the last year.
    """
    ic = cfg.raw["investment_case"]
    years = ic["years"]
    rev = np.array(ic["incremental_revenue"], dtype=float)
    upfront = np.array(ic["upfront_investment"], dtype=float)
    run_cost = np.array(ic["run_rate_cost"], dtype=float)
    contribution = rev * ic["contribution_margin"]
    pretax = contribution - run_cost - upfront
    tax = -np.maximum(pretax, 0) * ic["cash_tax_rate"]
    fcf = pretax + tax
    r, g = ic["discount_rate"], ic["terminal_growth"]
    disc = np.array([(1 + r) ** -(t + 1) for t in range(len(years))])
    pv = fcf * disc
    terminal = fcf[-1] * (1 + g) / (r - g)
    pv_terminal = terminal * disc[-1]
    npv = pv.sum() + pv_terminal

    # Payback: years until cumulative cash flow turns positive, interpolated within the year.
    cumulative = np.cumsum(fcf)
    payback = float("nan")
    prev = 0.0
    for t, cum in enumerate(cumulative):
        if cum >= 0:
            payback = t + (-prev / fcf[t] if prev < 0 else 0.0)
            break
        prev = cum

    table = pd.DataFrame({
        "incremental_revenue": rev,
        "contribution": contribution,
        "run_rate_cost": -run_cost,
        "upfront_investment": -upfront,
        "pre_tax_cash_flow": pretax,
        "cash_taxes": tax,
        "free_cash_flow": fcf,
        "cumulative_fcf": cumulative,
        "discount_factor": disc,
        "present_value": pv,
    }, index=years).T

    return {
        "name": ic["name"],
        "table": table,
        "npv": npv,
        "pv_terminal": pv_terminal,
        "irr_with_terminal": irr(list(fcf[:-1]) + [fcf[-1] + terminal]),
        "irr_no_terminal": irr(list(fcf)),
        "payback_years": payback,
        "hurdle_payback_years": ic["hurdle_payback_years"],
        "decision": "Approve" if npv > 0 and payback <= ic["hurdle_payback_years"]
                    else "Revisit",
    }
