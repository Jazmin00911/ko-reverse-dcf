import numpy as np
import pandas as pd
import pytest

from src.reverse_dcf.config import calculate_wacc
from src.reverse_dcf.dcf import calculate_dcf, enterprise_value, project_fcff
from src.reverse_dcf.fcff import calculate_fcff
from src.reverse_dcf.reverse_dcf import solve_implied_growth, solve_implied_growth_bisection
from src.reverse_dcf.sensitivities import ev_grid, implied_growth_grid


def test_fcff():
    df = pd.DataFrame({"ebit": [100.0], "tax_rate": [0.25], "depreciation_amortization": [10.0],
                       "capital_expenditures": [15.0], "change_in_nwc": [5.0]})
    out = calculate_fcff(df)
    assert out.loc[0, "nopat"] == pytest.approx(75.0)
    assert out.loc[0, "fcff"] == pytest.approx(65.0)
    df["tax_rate"] = 1.2
    with pytest.raises(ValueError, match="tax_rate"):
        calculate_fcff(df)


def test_wacc():
    w = calculate_wacc(0.05, 0.58, 0.0445, 0.0625, 0.21, 400.0, 40.0)
    assert w.cost_of_equity == pytest.approx(0.05 + 0.58 * 0.0445)
    assert w.wacc == pytest.approx(400 / 440 * w.cost_of_equity + 40 / 440 * 0.0625 * 0.79)


@pytest.mark.parametrize("mid", [False, True])
def test_closed_form_equals_explicit_schedule(mid):
    base, g, w, tg, n = 123.4, 0.067, 0.081, 0.025, 10
    explicit = calculate_dcf(project_fcff(base, g, n), w, tg, mid)["enterprise_value"]
    assert float(enterprise_value(base, g, w, tg, n, mid)) == pytest.approx(explicit, rel=1e-13)


def test_closed_form_handles_q_equal_one():
    # growth == WACC -> q = 1, annuity term = N
    ev = float(enterprise_value(100.0, 0.08, 0.08, 0.02, 5))
    explicit = calculate_dcf(project_fcff(100.0, 0.08, 5), 0.08, 0.02)["enterprise_value"]
    assert ev == pytest.approx(explicit, rel=1e-13)


def test_dcf_guards():
    with pytest.raises(ValueError, match="WACC must be greater"):
        calculate_dcf(pd.Series([1.0, 2.0]), 0.03, 0.03)
    with pytest.raises(ValueError, match="cannot be empty"):
        calculate_dcf(pd.Series(dtype=float), 0.1, 0.03)


@pytest.mark.parametrize("mid", [False, True])
def test_solver_recovers_known_growth(mid):
    target = float(enterprise_value(100.0, 0.06, 0.10, 0.025, 7, mid))
    res = solve_implied_growth(100.0, 7, 0.10, 0.025, target, mid_year=mid)
    assert res["implied_growth_rate"] == pytest.approx(0.06, abs=1e-12)
    assert abs(res["valuation_difference"]) < 1e-8
    bis = solve_implied_growth_bisection(100.0, 7, 0.10, 0.025, target, mid_year=mid)
    assert bis == pytest.approx(res["implied_growth_rate"], abs=1e-15)


def test_solver_unbracketed():
    with pytest.raises(ValueError, match="bracket"):
        solve_implied_growth(100.0, 5, 0.10, 0.025, 1e9)


def test_grids_monotonic():
    w, t = [0.07, 0.08, 0.09], [0.02, 0.03]
    grid = ev_grid(100.0, 0.05, w, t, 10)
    assert (grid.diff(axis=0).iloc[1:] < 0).all().all()      # higher WACC -> lower EV
    assert (grid.diff(axis=1).iloc[:, 1:] > 0).all().all()   # higher gT -> higher EV
    growth = implied_growth_grid(100.0, 2000.0, w, t, 10, -0.2, 0.3)
    assert (growth.diff(axis=0).iloc[1:] > 0).all().all()    # higher WACC -> more growth required
