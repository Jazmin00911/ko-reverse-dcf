"""Sensitivity grids (vectorised closed-form EV; re-solved implied growth)."""
from __future__ import annotations

import numpy as np
import pandas as pd

from .dcf import enterprise_value
from .reverse_dcf import solve_implied_growth


def centred_axis(centre: float, step: float, points: int = 5) -> np.ndarray:
    """Symmetric axis around ``centre`` (the centre cell ties to the base case)."""
    offsets = np.arange(points) - points // 2
    return centre + offsets * step


def ev_grid(base_fcff: float, growth: float, wacc_axis, tg_axis, years: int,
            mid_year: bool = False) -> pd.DataFrame:
    """EV for WACC (rows) x terminal growth (columns) holding explicit growth fixed."""
    w = np.asarray(wacc_axis, float)[:, None]
    t = np.asarray(tg_axis, float)[None, :]
    values = enterprise_value(base_fcff, growth, w, t, years, mid_year)
    return pd.DataFrame(values, index=pd.Index(wacc_axis, name="WACC"),
                        columns=pd.Index(tg_axis, name="Terminal growth"))


def implied_growth_grid(base_fcff: float, market_ev: float, wacc_axis, tg_axis, years: int,
                        lower: float, upper: float, mid_year: bool = False) -> pd.DataFrame:
    """Market-implied growth re-solved for each (WACC, terminal growth) pair."""
    out = np.empty((len(wacc_axis), len(tg_axis)))
    for i, w in enumerate(wacc_axis):
        for j, t in enumerate(tg_axis):
            out[i, j] = solve_implied_growth(base_fcff, years, w, t, market_ev,
                                             lower, upper, mid_year)["implied_growth_rate"]
    return pd.DataFrame(out, index=pd.Index(wacc_axis, name="WACC"),
                        columns=pd.Index(tg_axis, name="Terminal growth"))


def price_grid(base_fcff: float, growth_axis, wacc_axis, terminal_growth: float, years: int,
               ev_to_equity_bridge: float, shares: float, mid_year: bool = False) -> pd.DataFrame:
    """Intrinsic value per share for alternative growth paths (rows) x WACC (columns).

    ``ev_to_equity_bridge`` is (equity value - EV), i.e. cash & non-operating
    assets minus debt, leases and minority interests.
    """
    g = np.asarray(growth_axis, float)[:, None]
    w = np.asarray(wacc_axis, float)[None, :]
    ev = enterprise_value(base_fcff, g, w, terminal_growth, years, mid_year)
    return pd.DataFrame((ev + ev_to_equity_bridge) / shares,
                        index=pd.Index(growth_axis, name="FCFF growth"),
                        columns=pd.Index(wacc_axis, name="WACC"))
