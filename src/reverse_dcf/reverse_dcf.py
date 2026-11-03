"""Reverse DCF: solve for the growth rate the market price already discounts."""
from __future__ import annotations

import numpy as np
from scipy.optimize import brentq

from .dcf import enterprise_value


def _validate(base_fcff, forecast_years, wacc, terminal_growth_rate,
              market_enterprise_value, lower_bound, upper_bound) -> None:
    values = [base_fcff, wacc, terminal_growth_rate, market_enterprise_value,
              lower_bound, upper_bound]
    if not np.isfinite(values).all():
        raise ValueError("All numeric inputs must be finite.")
    if base_fcff <= 0:
        raise ValueError("base_fcff must be positive.")
    if forecast_years < 1:
        raise ValueError("forecast_years must be at least 1.")
    if market_enterprise_value <= 0:
        raise ValueError("Market enterprise value must be positive.")
    if lower_bound >= upper_bound:
        raise ValueError("lower_bound must be less than upper_bound.")
    if lower_bound <= -1:
        raise ValueError("lower_bound must be greater than -100%.")
    if wacc <= terminal_growth_rate:
        raise ValueError("WACC must exceed terminal growth.")


def solve_implied_growth(
    base_fcff: float,
    forecast_years: int,
    wacc: float,
    terminal_growth_rate: float,
    market_enterprise_value: float,
    lower_bound: float = -0.20,
    upper_bound: float = 0.30,
    mid_year: bool = False,
) -> dict:
    """Constant annual FCFF growth g over the explicit period such that

        EV_model(g; WACC, gT, FCFF_0, N) = EV_market

    ``base_fcff`` is the FCFF of the year *before* the forecast (TTM).  EV is
    strictly increasing in g, so the root is unique and Brent's method on the
    closed form converges to machine precision.
    """
    _validate(base_fcff, forecast_years, wacc, terminal_growth_rate,
              market_enterprise_value, lower_bound, upper_bound)

    def gap(g: float) -> float:
        return float(enterprise_value(base_fcff, g, wacc, terminal_growth_rate,
                                      forecast_years, mid_year)) - market_enterprise_value

    f_lo, f_hi = gap(lower_bound), gap(upper_bound)
    if f_lo == 0:
        root = lower_bound
    elif f_hi == 0:
        root = upper_bound
    elif f_lo * f_hi > 0:
        raise ValueError("Could not bracket an implied growth rate within the "
                         "supplied bounds. Widen the bounds or review the inputs.")
    else:
        root = brentq(gap, lower_bound, upper_bound, xtol=1e-15, rtol=4 * np.finfo(float).eps,
                      maxiter=500)

    residual = gap(root)
    return {
        "implied_growth_rate": float(root),
        "market_enterprise_value": float(market_enterprise_value),
        "implied_enterprise_value": float(market_enterprise_value + residual),
        "valuation_difference": float(residual),
    }


def solve_implied_growth_bisection(
    base_fcff: float, forecast_years: int, wacc: float, terminal_growth_rate: float,
    market_enterprise_value: float, lower_bound: float = -0.20,
    upper_bound: float = 0.30, mid_year: bool = False, iterations: int = 60,
) -> float:
    """Bisection with a fixed iteration count - algorithmically identical to the
    Excel 'Solver' sheet, so the two can be compared bit-for-bit."""
    lo, hi = lower_bound, upper_bound
    for _ in range(iterations):
        mid = 0.5 * (lo + hi)
        ev = float(enterprise_value(base_fcff, mid, wacc, terminal_growth_rate,
                                    forecast_years, mid_year))
        if ev < market_enterprise_value:
            lo = mid
        else:
            hi = mid
    return 0.5 * (lo + hi)
