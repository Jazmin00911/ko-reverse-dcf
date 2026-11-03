"""Enterprise-value mechanics: explicit schedule *and* closed form.

Both implementations discount identical cash flows and must agree to
floating-point precision (enforced by tests).  The closed form is what the
reverse-DCF solver and the Excel bisection use, because with constant growth
``g`` the explicit-period PV is a finite geometric series in
``q = (1+g)/(1+WACC)``:

    EV(g) = (1+WACC)^(m/2) * F0 * [ q(1-q^N)/(1-q)  +  q^N (1+gT)/(WACC-gT) ]

where F0 is the base-year FCFF, N the explicit years, gT terminal growth and
m = 1 for the mid-year convention (0 for year-end discounting).
"""
from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import pandas as pd

ArrayLike = float | np.ndarray


@dataclass(frozen=True)
class DcfSchedule:
    years: np.ndarray
    fcff: np.ndarray
    discount_periods: np.ndarray
    discount_factors: np.ndarray
    pv_fcff: np.ndarray
    terminal_fcff: float
    terminal_value: float
    terminal_discount_factor: float
    pv_terminal_value: float
    sum_pv_fcff: float
    enterprise_value: float

    @property
    def terminal_share_of_ev(self) -> float:
        return self.pv_terminal_value / self.enterprise_value


def _validate_rates(wacc: ArrayLike, terminal_growth: ArrayLike) -> None:
    wacc_a, tg_a = np.asarray(wacc, float), np.asarray(terminal_growth, float)
    if not (np.isfinite(wacc_a).all() and np.isfinite(tg_a).all()):
        raise ValueError("WACC and terminal growth must be finite numbers.")
    if (wacc_a <= -1).any() or (tg_a <= -1).any():
        raise ValueError("WACC and terminal growth must exceed -100 %.")
    if (wacc_a <= tg_a).any():
        raise ValueError("WACC must be greater than the terminal growth rate.")


def project_fcff(base_fcff: float, growth: ArrayLike, years: int) -> np.ndarray:
    """FCFF_t = FCFF_0 (1+g)^t for t = 1..years (vectorised)."""
    t = np.arange(1, years + 1)
    return base_fcff * (1.0 + np.asarray(growth, float)[..., None]) ** t


def calculate_dcf(fcff_forecast: pd.Series | np.ndarray, wacc: float,
                  terminal_growth_rate: float, mid_year: bool = False) -> dict:
    """Explicit year-by-year DCF with a Gordon-growth terminal value.

    Terminal value is FCFF_N (1+gT)/(WACC-gT).  Under the mid-year convention
    the Gordon terminal value is discounted with the *same* N - 0.5 factor as
    the final explicit cash flow (the standard treatment when the perpetuity is
    also assumed to arrive mid-year).
    """
    cash_flows = np.asarray(fcff_forecast, dtype=float)
    if cash_flows.ndim != 1 or cash_flows.size == 0:
        raise ValueError("FCFF forecast cannot be empty.")
    if not np.isfinite(cash_flows).all():
        raise ValueError("FCFF forecast contains missing or non-finite values.")
    _validate_rates(wacc, terminal_growth_rate)

    years = np.arange(1, cash_flows.size + 1)
    periods = years - (0.5 if mid_year else 0.0)
    factors = (1.0 + wacc) ** (-periods)
    pv = cash_flows * factors

    terminal_fcff = cash_flows[-1] * (1.0 + terminal_growth_rate)
    terminal_value = terminal_fcff / (wacc - terminal_growth_rate)
    tv_factor = factors[-1]
    pv_tv = terminal_value * tv_factor
    ev = pv.sum() + pv_tv

    return {
        "schedule": DcfSchedule(years, cash_flows, periods, factors, pv,
                                float(terminal_fcff), float(terminal_value),
                                float(tv_factor), float(pv_tv),
                                float(pv.sum()), float(ev)),
        "discounted_fcff": pv,
        "sum_discounted_fcff": float(pv.sum()),
        "terminal_fcff": float(terminal_fcff),
        "terminal_value": float(terminal_value),
        "discounted_terminal_value": float(pv_tv),
        "enterprise_value": float(ev),
    }


def enterprise_value(base_fcff: float, growth: ArrayLike, wacc: ArrayLike,
                     terminal_growth: ArrayLike, years: int,
                     mid_year: bool = False) -> np.ndarray:
    """Closed-form EV; broadcasts over ``growth``, ``wacc`` and ``terminal_growth``."""
    g = np.asarray(growth, float)
    w = np.asarray(wacc, float)
    tg = np.asarray(terminal_growth, float)
    _validate_rates(w, tg)

    q = (1.0 + g) / (1.0 + w)
    with np.errstate(divide="ignore", invalid="ignore"):
        annuity = np.where(np.isclose(q, 1.0, rtol=0, atol=1e-12),
                           float(years), q * (1.0 - q ** years) / (1.0 - q))
    terminal = q ** years * (1.0 + tg) / (w - tg)
    shift = (1.0 + w) ** (0.5 if mid_year else 0.0)
    return shift * base_fcff * (annuity + terminal)
