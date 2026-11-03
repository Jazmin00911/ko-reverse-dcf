from __future__ import annotations
import json
from pathlib import Path

import pandas as pd

from .config import OUTPUT_DIR
from .model import ModelResult


def summary_dict(r: ModelResult) -> dict:
    f = r.fundamentals
    return {
        "valuation_date": r.market.valuation_date.isoformat(),
        "sec_basis": r.snapshot.basis_label,
        "latest_filing_used": r.snapshot.filing_date.isoformat(),
        "share_price": r.market.share_price,
        "diluted_shares_m": r.diluted_shares,
        "market_cap_usd_m": r.market_cap,
        "market_enterprise_value_usd_m": r.market_ev,
        "wacc": r.wacc.wacc, "cost_of_equity": r.wacc.cost_of_equity,
        "after_tax_cost_of_debt": r.wacc.after_tax_cost_of_debt,
        "equity_weight": r.wacc.equity_weight, "debt_weight": r.wacc.debt_weight,
        "base_fcff_usd_m": r.base["fcff"], "ttm_tax_rate": r.base["tax_rate"],
        "terminal_growth_rate": r.market.terminal_growth_rate,
        "forecast_years": r.settings.forecast_years, "mid_year": r.settings.mid_year,
        "implied_fcff_growth": r.implied_growth,
        "implied_fcff_growth_bisection": r.implied_growth_bisection,
        "valuation_residual_usd_m": r.implied["valuation_difference"],
        "model_enterprise_value_usd_m": r.schedule.enterprise_value,
        "terminal_value_share_of_ev": r.schedule.terminal_share_of_ev,
        "implied_share_price": r.implied_share_price,
        "roic_ttm": f.roic, "revenue_cagr": f.revenue_cagr, "ebit_cagr": f.ebit_cagr,
        "fundamental_growth_ttm": f.fundamental_growth_ttm,
        "fundamental_growth_normalised": f.fundamental_growth_normalised,
        "required_fcff_margin_year_n": {str(k): v for k, v in f.required_fcff_margin_year_n.items()},
    }


def export_results(r: ModelResult, directory: Path | None = None) -> dict[str, Path]:
    directory = Path(directory) if directory else OUTPUT_DIR
    directory.mkdir(parents=True, exist_ok=True)
    json_path = directory / "reverse_dcf_summary.json"
    json_path.write_text(json.dumps(summary_dict(r), indent=2), encoding="utf-8")

    rows = []
    for item in r.snapshot.ttm.values():
        for src in item.sources:
            rows.append({"item": item.metric, "source": src})
    for b in list(r.snapshot.balance.values()) + [r.snapshot.operating_lease_liability]:
        rows.append({"item": b.metric, "source": b.source})
    rows.append({"item": "shares", "source": r.snapshot.shares_source})
    prov_path = directory / "sec_provenance.csv"
    pd.DataFrame(rows).to_csv(prov_path, index=False)

    grids = directory / "sensitivity_grids.xlsx"
    with pd.ExcelWriter(grids, engine="openpyxl") as xw:
        r.ev_grid.to_excel(xw, sheet_name="EV (USD m)")
        r.growth_grid.to_excel(xw, sheet_name="Implied growth")
        r.price_grid.to_excel(xw, sheet_name="Value per share")
    return {"json": json_path, "provenance": prov_path, "grids": grids}
