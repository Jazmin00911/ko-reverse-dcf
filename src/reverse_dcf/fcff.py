"""Free cash flow to the firm (FCFF)."""
from __future__ import annotations

import pandas as pd

REQUIRED_COLUMNS = (
    "ebit", "tax_rate", "depreciation_amortization",
    "capital_expenditures", "change_in_nwc",
)


def calculate_fcff(df: pd.DataFrame) -> pd.DataFrame:
    """Vectorised FCFF:  EBIT(1-t) + D&A - CapEx - change in NWC.

    Conventions
    -----------
    * ``tax_rate`` is a decimal (0.18 == 18 %).
    * ``capital_expenditures`` is a positive cash outflow.
    * ``change_in_nwc`` is positive when working capital *absorbs* cash
      (reduces FCFF) and negative when it releases cash.
    """
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"Cannot calculate FCFF. Missing columns: {missing}")
    if df[list(REQUIRED_COLUMNS)].isna().any().any():
        raise ValueError("Cannot calculate FCFF: required inputs contain missing values.")
    if not df["tax_rate"].between(0, 1).all():
        raise ValueError("tax_rate must be between 0 and 1.")

    result = df.copy()
    result["nopat"] = result["ebit"] * (1.0 - result["tax_rate"])
    result["fcff"] = (
        result["nopat"]
        + result["depreciation_amortization"]
        - result["capital_expenditures"]
        - result["change_in_nwc"]
    )
    return result


def base_fcff_from_ttm(ebit: float, pretax_income: float, income_tax: float,
                       d_and_a: float, capex: float, change_in_nwc: float) -> dict:
    """Single-period helper used for the base-year (TTM) FCFF.

    The tax rate is the TTM *effective* rate (tax expense / pre-tax income),
    a standard Damodaran proxy for the near-term cash tax burden.
    """
    if pretax_income <= 0:
        raise ValueError("Income before income taxes must be positive to infer a tax rate.")
    tax_rate = income_tax / pretax_income
    frame = calculate_fcff(pd.DataFrame([{
        "ebit": ebit, "tax_rate": tax_rate,
        "depreciation_amortization": d_and_a,
        "capital_expenditures": capex, "change_in_nwc": change_in_nwc,
    }]))
    row = frame.iloc[0]
    return {"tax_rate": float(tax_rate), "nopat": float(row["nopat"]), "fcff": float(row["fcff"])}
