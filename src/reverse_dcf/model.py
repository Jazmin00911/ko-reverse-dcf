"""End-to-end reverse-DCF calculation (Python; the Excel workbook re-implements it)."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date

import numpy as np
import pandas as pd

from .config import MarketInputs, ModelSettings, WaccResult, calculate_wacc
from .dcf import DcfSchedule, calculate_dcf, project_fcff
from .fcff import base_fcff_from_ttm
from .fundamentals import HISTORY_FIRST_YEAR, Fundamentals, compute_fundamentals
from .reverse_dcf import solve_implied_growth, solve_implied_growth_bisection
from .sec import CompanyFacts, SecSnapshot, annual_history, build_snapshot
from .sensitivities import centred_axis, ev_grid, implied_growth_grid, price_grid


@dataclass(frozen=True)
class ModelResult:
    market: MarketInputs
    settings: ModelSettings
    snapshot: SecSnapshot
    history: dict
    wacc: WaccResult
    # capital structure & bridge (USD millions)
    diluted_shares: float
    market_cap: float
    debt_total: float
    operating_leases: float
    noncontrolling_interests: float
    cash_and_securities: float
    equity_method_adjustment: float
    market_ev: float
    ev_to_equity_bridge: float
    # base-year FCFF
    base: dict
    # reverse DCF
    implied: dict
    implied_growth_bisection: float
    schedule: DcfSchedule
    equity_value: float
    implied_share_price: float
    fundamentals: Fundamentals
    # grids
    wacc_axis: np.ndarray
    tg_axis: np.ndarray
    ev_grid: pd.DataFrame
    growth_grid: pd.DataFrame
    price_grid: pd.DataFrame
    price_growth_axis: np.ndarray

    @property
    def implied_growth(self) -> float:
        return self.implied["implied_growth_rate"]


def run_model(market: MarketInputs, settings: ModelSettings, facts: CompanyFacts) -> ModelResult:
    market.validate()
    settings.validate()
    snap = build_snapshot(facts, market.valuation_date)
    history = annual_history(facts, market.valuation_date, HISTORY_FIRST_YEAR)

    b = snap.balance
    shares = snap.diluted_shares_m
    market_cap = market.share_price * shares
    leases = snap.operating_lease_liability.value
    debt_total = snap.total_debt + leases
    nci = b["noncontrolling_interests"].value
    cash = (b["cash_and_equivalents"].value + b["short_term_investments"].value
            + b["marketable_securities"].value)
    eq_method = b["equity_method_investments"].value if settings.include_nonoperating else 0.0

    market_ev = market_cap + debt_total + nci - cash - eq_method
    bridge = cash + eq_method - debt_total - nci

    wacc = calculate_wacc(market.risk_free_rate, market.beta, market.equity_risk_premium,
                          market.pre_tax_cost_of_debt, market.marginal_tax_rate,
                          market_cap, debt_total)

    t = {m: snap.ttm_value(m) for m in snap.ttm}
    base = base_fcff_from_ttm(t["ebit"], t["income_before_income_taxes"], t["income_tax_expense"],
                              t["depreciation_amortization"], t["capital_expenditures"],
                              t["change_in_nwc"])
    base.update(ebit=t["ebit"], d_and_a=t["depreciation_amortization"],
                capex=t["capital_expenditures"], change_in_nwc=t["change_in_nwc"],
                revenue=t["revenue"])
    if base["fcff"] <= 0:
        raise ValueError("Base FCFF is not positive; a growth-rate reverse DCF is undefined.")

    args = dict(base_fcff=base["fcff"], forecast_years=settings.forecast_years,
                wacc=wacc.wacc, terminal_growth_rate=market.terminal_growth_rate,
                market_enterprise_value=market_ev, lower_bound=settings.growth_lower_bound,
                upper_bound=settings.growth_upper_bound, mid_year=settings.mid_year)
    implied = solve_implied_growth(**args)
    bisect = solve_implied_growth_bisection(**args)
    g = implied["implied_growth_rate"]

    forecast = project_fcff(base["fcff"], g, settings.forecast_years)
    schedule = calculate_dcf(forecast, wacc.wacc, market.terminal_growth_rate,
                             settings.mid_year)["schedule"]
    equity_value = schedule.enterprise_value + bridge
    implied_price = equity_value / shares

    invested_capital = (b["total_equity"].value + debt_total - cash - eq_method)
    fundamentals = compute_fundamentals(
        history, ttm=t, ttm_tax_rate=base["tax_rate"], invested_capital=invested_capital,
        base_fcff=base["fcff"], implied_growth=g, forecast_years=settings.forecast_years)

    w_axis = centred_axis(wacc.wacc, settings.wacc_step)
    t_axis = centred_axis(market.terminal_growth_rate, settings.terminal_growth_step)
    g_axis = centred_axis(g, 0.01)
    grids = dict(
        ev=ev_grid(base["fcff"], g, w_axis, t_axis, settings.forecast_years, settings.mid_year),
        growth=implied_growth_grid(base["fcff"], market_ev, w_axis, t_axis,
                                   settings.forecast_years, settings.growth_lower_bound,
                                   settings.growth_upper_bound, settings.mid_year),
        price=price_grid(base["fcff"], g_axis, w_axis, market.terminal_growth_rate,
                         settings.forecast_years, bridge, shares, settings.mid_year),
    )
    return ModelResult(
        market=market, settings=settings, snapshot=snap, history=history, wacc=wacc,
        diluted_shares=shares, market_cap=market_cap, debt_total=debt_total,
        operating_leases=leases, noncontrolling_interests=nci, cash_and_securities=cash,
        equity_method_adjustment=eq_method, market_ev=market_ev, ev_to_equity_bridge=bridge,
        base=base, implied=implied, implied_growth_bisection=bisect, schedule=schedule,
        equity_value=equity_value, implied_share_price=implied_price, fundamentals=fundamentals,
        wacc_axis=w_axis, tg_axis=t_axis, ev_grid=grids["ev"], growth_grid=grids["growth"],
        price_grid=grids["price"], price_growth_axis=g_axis,
    )
