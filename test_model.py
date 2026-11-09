"""Tests: historicals tie to filings, the forecast is internally consistent,
FY2026 ties to guidance, and the analysis tools behave sensibly."""
import math

import pytest

from lrp.analysis import evaluate_investment, irr, tornado, two_way
from lrp.model import guidance_check, load_config, load_historicals, project, run_all


@pytest.fixture(scope="module")
def cfg():
    return load_config()


@pytest.fixture(scope="module")
def results(cfg):
    return run_all(cfg)


def test_historical_gaap_operating_income_ties_to_filings():
    h = load_historicals()
    # Gross profit - opex should equal reported GAAP operating income every year.
    implied = h.loc["gross_profit"] - h.loc["total_opex"]
    assert (implied - h.loc["gaap_oi"]).abs().max() < 0.5


def test_historical_revenue_adds_up():
    h = load_historicals()
    assert (h.loc["sub_rev"] + h.loc["ps_rev"] - h.loc["total_rev"]).abs().max() < 0.5


@pytest.mark.parametrize("scenario", ["base", "upside", "downside"])
def test_forecast_identities(results, cfg, scenario):
    df = results[scenario]
    for y in cfg.years:
        c = df[y]
        assert c["total_rev"] == pytest.approx(c["sub_rev"] + c["ps_rev"])
        assert c["gaap_oi"] == pytest.approx(c["gross_profit"] - c["total_opex"])
        assert c["ng_oi"] == pytest.approx(c["gaap_oi"] + c["sbc"] + c["other_adj"])
        assert c["fcf"] == pytest.approx(c["ocf"] - c["capex"])
        assert c["headcount"] * c["rev_per_emp"] / 1000 == pytest.approx(c["total_rev"])


def test_fy2026_ties_to_guidance(cfg, results):
    g = guidance_check(cfg, results["base"])
    rev = g.iloc[0]
    assert abs(rev["difference"]) / rev["guidance"] < 0.001       # within 0.1%
    assert (g.iloc[1:]["difference"].abs() < 0.001).all()           # within 10 bp


def test_fy2026_identical_across_scenarios(results):
    base = results["base"][2026]
    for s in ("upside", "downside"):
        assert (results[s][2026] - base).abs().max() < 1e-9


def test_scenarios_are_ordered(results):
    fcf = {s: df.loc["fcf", 2030] for s, df in results.items()}
    assert fcf["downside"] < fcf["base"] < fcf["upside"]


def test_override_changes_only_what_it_should(cfg):
    base = project(cfg, "base")
    bumped = project(cfg, "base", overrides={"capex_pct": [0.03, 0.05, 0.05, 0.05, 0.05]})
    assert bumped.loc["total_rev", 2030] == pytest.approx(base.loc["total_rev", 2030])
    assert bumped.loc["fcf", 2030] < base.loc["fcf", 2030]


def test_tornado_sorted_and_growth_matters_most(cfg):
    t = tornado(cfg)
    assert list(t["swing"]) == sorted(t["swing"], reverse=True)
    assert t.iloc[0]["driver"] == "sub_growth"


def test_two_way_center_equals_base(cfg, results):
    grid = two_way(cfg)
    assert grid.loc[0.0, 0.0] == pytest.approx(results["base"].loc["fcf", 2030])


def test_irr_known_value():
    # -100 today, +110 in one year -> 10%
    assert irr([-100, 110]) == pytest.approx(0.10, abs=1e-6)


def test_investment_case_outputs(cfg):
    inv = evaluate_investment(cfg)
    assert inv["npv"] > inv["pv_terminal"] > 0
    assert 0 < inv["irr_no_terminal"] < 1
    assert not math.isnan(inv["payback_years"])
    assert inv["decision"] in {"Approve", "Revisit"}
