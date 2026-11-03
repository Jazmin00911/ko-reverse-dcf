"""Configuration objects and cost-of-capital maths.

All monetary amounts in this package are **USD millions** and all share counts
are **millions of shares** unless a name says otherwise.  Rates are decimals
(0.05 == 5 %).
"""
from __future__ import annotations

import csv
from dataclasses import dataclass, asdict
from datetime import date
from pathlib import Path
import numpy as np

PROJECT_ROOT = Path(__file__).resolve().parents[2]
DATA_DIR = PROJECT_ROOT / "data"
OUTPUT_DIR = PROJECT_ROOT / "outputs"


# --------------------------------------------------------------------------- #
# Market / macro inputs (the only numbers that do NOT come from the SEC)
# --------------------------------------------------------------------------- #
@dataclass(frozen=True)
class MarketInputs:
    """Market data and valuation assumptions, loaded from data/market_inputs.csv."""

    valuation_date: date
    share_price: float
    risk_free_rate: float
    equity_risk_premium: float
    beta: float
    pre_tax_cost_of_debt: float
    marginal_tax_rate: float
    terminal_growth_rate: float
    #: per-input provenance: name -> (units, source, as_of, notes)
    provenance: dict

    def validate(self) -> None:
        numeric = [
            self.share_price, self.risk_free_rate, self.equity_risk_premium,
            self.beta, self.pre_tax_cost_of_debt, self.marginal_tax_rate,
            self.terminal_growth_rate,
        ]
        if not np.isfinite(numeric).all():
            raise ValueError("All market inputs must be finite numbers.")
        if self.share_price <= 0:
            raise ValueError("share_price must be positive.")
        if self.beta < 0:
            raise ValueError("beta cannot be negative.")
        for name in ("risk_free_rate", "equity_risk_premium",
                     "pre_tax_cost_of_debt", "marginal_tax_rate"):
            if not 0 <= getattr(self, name) < 1:
                raise ValueError(f"{name} must be in [0, 1).")
        if self.terminal_growth_rate >= self.risk_free_rate:
            # Damodaran: a perpetual growth rate above the risk-free rate
            # implies the firm eventually outgrows the whole economy.
            raise ValueError(
                "terminal_growth_rate must be below the risk-free rate "
                "(Damodaran's cap on stable growth)."
            )


@dataclass(frozen=True)
class ModelSettings:
    """Structural switches shared by the Python engine and the Excel model."""

    forecast_years: int = 10
    mid_year: bool = False                 # mid-year discounting convention
    include_nonoperating: bool = True      # subtract equity-method inv. from EV
    growth_lower_bound: float = -0.20
    growth_upper_bound: float = 0.30
    wacc_step: float = 0.005               # sensitivity-grid step
    terminal_growth_step: float = 0.005

    def validate(self) -> None:
        if self.forecast_years < 1:
            raise ValueError("forecast_years must be >= 1.")
        if not self.growth_lower_bound < self.growth_upper_bound:
            raise ValueError("growth bounds must satisfy lower < upper.")
        if self.growth_lower_bound <= -1:
            raise ValueError("growth_lower_bound must exceed -100 %.")


def load_market_inputs(path: Path | None = None) -> MarketInputs:
    """Read data/market_inputs.csv (columns: input,value,units,source,as_of,notes)."""
    path = Path(path) if path else DATA_DIR / "market_inputs.csv"
    if not path.exists():
        raise FileNotFoundError(f"Market-input file not found: {path}")

    rows: dict[str, dict] = {}
    with path.open(newline="", encoding="utf-8") as handle:
        for row in csv.DictReader(handle):
            name = row["input"].strip()
            if name in rows:
                raise ValueError(f"Duplicate market input: {name}")
            rows[name] = row

    required = [
        "valuation_date", "share_price", "risk_free_rate",
        "equity_risk_premium", "beta", "pre_tax_cost_of_debt",
        "marginal_tax_rate", "terminal_growth_rate",
    ]
    missing = [name for name in required if name not in rows]
    if missing:
        raise ValueError(f"market_inputs.csv is missing: {missing}")

    def num(name: str) -> float:
        return float(rows[name]["value"])

    inputs = MarketInputs(
        valuation_date=date.fromisoformat(rows["valuation_date"]["value"].strip()),
        share_price=num("share_price"),
        risk_free_rate=num("risk_free_rate"),
        equity_risk_premium=num("equity_risk_premium"),
        beta=num("beta"),
        pre_tax_cost_of_debt=num("pre_tax_cost_of_debt"),
        marginal_tax_rate=num("marginal_tax_rate"),
        terminal_growth_rate=num("terminal_growth_rate"),
        provenance={
            name: (r["units"], r["source"], r["as_of"], r["notes"])
            for name, r in rows.items()
        },
    )
    inputs.validate()
    return inputs


# --------------------------------------------------------------------------- #
# Cost of capital (Damodaran: CAPM cost of equity, market-value weights)
# --------------------------------------------------------------------------- #
@dataclass(frozen=True)
class WaccResult:
    cost_of_equity: float
    after_tax_cost_of_debt: float
    equity_weight: float
    debt_weight: float
    total_capital: float
    wacc: float

    def as_dict(self) -> dict:
        return asdict(self)


def calculate_wacc(
    risk_free_rate: float,
    beta: float,
    equity_risk_premium: float,
    pre_tax_cost_of_debt: float,
    tax_rate: float,
    market_capitalization: float,
    total_debt: float,
) -> WaccResult:
    """WACC = E/(D+E)*[rf + beta*ERP] + D/(D+E)*kd*(1-t).

    ``market_capitalization`` and ``total_debt`` must be in the same units
    (only their ratio matters).  ``tax_rate`` is the *marginal* rate, because
    the interest tax shield is earned at the margin.
    """
    inputs = [risk_free_rate, beta, equity_risk_premium, pre_tax_cost_of_debt,
              tax_rate, market_capitalization, total_debt]
    if not np.isfinite(inputs).all():
        raise ValueError("All WACC inputs must be finite.")
    if not 0 <= risk_free_rate < 1:
        raise ValueError("Risk-free rate must be between 0 and 1.")
    if beta < 0:
        raise ValueError("Beta cannot be negative.")
    if not 0 <= equity_risk_premium < 1:
        raise ValueError("Equity risk premium must be between 0 and 1.")
    if not 0 <= pre_tax_cost_of_debt < 1:
        raise ValueError("Cost of debt must be between 0 and 1.")
    if not 0 <= tax_rate < 1:
        raise ValueError("Tax rate must be between 0 and 1.")
    if market_capitalization <= 0 or total_debt < 0:
        raise ValueError("Market cap must be positive and debt non-negative.")

    total_capital = market_capitalization + total_debt
    cost_of_equity = risk_free_rate + beta * equity_risk_premium
    after_tax_kd = pre_tax_cost_of_debt * (1.0 - tax_rate)
    w_e = market_capitalization / total_capital
    w_d = total_debt / total_capital
    return WaccResult(
        cost_of_equity=cost_of_equity,
        after_tax_cost_of_debt=after_tax_kd,
        equity_weight=w_e,
        debt_weight=w_d,
        total_capital=total_capital,
        wacc=w_e * cost_of_equity + w_d * after_tax_kd,
    )
