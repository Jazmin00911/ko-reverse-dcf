"""Historical fundamentals used to judge whether the implied growth is plausible.

Damodaran's test: growth must be *earned* - expected growth in operating income
= reinvestment rate x return on invested capital (+ efficiency gains).  We
therefore compare the market-implied FCFF growth with (i) observed revenue and
EBIT growth, (ii) the growth that today's reinvestment and ROIC can support and
(iii) the FCFF margin that would be required in year N.
"""
from __future__ import annotations

from dataclasses import dataclass

HISTORY_FIRST_YEAR = 2019            # FY2019..FY2025 shown in the model
CAGR_WINDOWS = (3, 5, 6)             # years back from the latest fiscal year
NORMALISED_WINDOW = 5                # years used for the normalised reinvestment rate
REVENUE_SCENARIOS = (0.03, 0.05, 0.07)


def cagr(end_value: float, start_value: float, years: int) -> float:
    if start_value <= 0 or end_value <= 0:
        raise ValueError("CAGR needs positive start and end values.")
    return (end_value / start_value) ** (1.0 / years) - 1.0


@dataclass(frozen=True)
class Fundamentals:
    latest_fiscal_year: int
    revenue_cagr: dict[int, float]
    ebit_cagr: dict[int, float]
    effective_tax_rate: dict[int, float]
    nopat_ttm: float
    invested_capital: float
    roic: float
    reinvestment_rate_ttm: float
    fundamental_growth_ttm: float
    reinvestment_rate_normalised: float
    fundamental_growth_normalised: float
    fcff_margin_ttm: float
    nopat_margin_ttm: float
    required_fcff_margin_year_n: dict[float, float]


def compute_fundamentals(history: dict[str, dict[int, float]], *, ttm: dict[str, float],
                         ttm_tax_rate: float, invested_capital: float, base_fcff: float,
                         implied_growth: float, forecast_years: int) -> Fundamentals:
    years = sorted(history["revenue"])
    last = years[-1]
    rev, ebit = history["revenue"], history["ebit"]
    tax, pretax = history["income_tax_expense"], history["income_before_income_taxes"]
    capex, dna = history["capital_expenditures"], history["depreciation_amortization"]

    rev_cagr = {n: cagr(rev[last], rev[last - n], n) for n in CAGR_WINDOWS}
    ebit_cagr = {n: cagr(ebit[last], ebit[last - n], n) for n in CAGR_WINDOWS}
    eff_tax = {y: tax[y] / pretax[y] for y in years}

    nopat_ttm = ttm["ebit"] * (1.0 - ttm_tax_rate)
    roic = nopat_ttm / invested_capital
    rr_ttm = (ttm["capital_expenditures"] - ttm["depreciation_amortization"]
              + ttm["change_in_nwc"]) / nopat_ttm

    window = [y for y in years if y > last - NORMALISED_WINDOW]
    net_capex = sum(capex[y] - dna[y] for y in window)
    nopat_window = sum(ebit[y] * (1.0 - eff_tax[y]) for y in window)
    rr_norm = net_capex / nopat_window

    required = {
        r: base_fcff * (1.0 + implied_growth) ** forecast_years
        / (ttm["revenue"] * (1.0 + r) ** forecast_years)
        for r in REVENUE_SCENARIOS
    }
    return Fundamentals(
        latest_fiscal_year=last, revenue_cagr=rev_cagr, ebit_cagr=ebit_cagr,
        effective_tax_rate=eff_tax, nopat_ttm=nopat_ttm, invested_capital=invested_capital,
        roic=roic, reinvestment_rate_ttm=rr_ttm, fundamental_growth_ttm=roic * rr_ttm,
        reinvestment_rate_normalised=rr_norm, fundamental_growth_normalised=roic * rr_norm,
        fcff_margin_ttm=base_fcff / ttm["revenue"], nopat_margin_ttm=nopat_ttm / ttm["revenue"],
        required_fcff_margin_year_n=required,
    )
