"""Run the Coca-Cola reverse DCF and write the Excel model.

    python run_model.py                       # uses data/market_inputs.csv
    python run_model.py --valuation-date 2026-10-28 --price 95.10
    python run_model.py --mid-year --simple-bridge
    python run_model.py --refresh             # re-download SEC facts (needs SEC_USER_AGENT)
"""
from __future__ import annotations

import argparse
from dataclasses import replace
from datetime import date

from src.reverse_dcf.config import ModelSettings, OUTPUT_DIR, load_market_inputs
from src.reverse_dcf.excel_model import write_model_workbook
from src.reverse_dcf.model import run_model
from src.reverse_dcf.outputs import export_results
from src.reverse_dcf.sec import CompanyFacts


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--valuation-date", type=date.fromisoformat, help="override market_inputs.csv")
    p.add_argument("--price", type=float, help="override share price (USD)")
    p.add_argument("--mid-year", action="store_true", help="mid-year discounting convention")
    p.add_argument("--simple-bridge", action="store_true",
                   help="do not subtract equity-method investments from EV")
    p.add_argument("--refresh", action="store_true", help="download fresh SEC company facts")
    p.add_argument("--output", default="reverse_dcf_model.xlsx")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    market = load_market_inputs()
    if args.valuation_date:
        market = replace(market, valuation_date=args.valuation_date)
    if args.price:
        market = replace(market, share_price=args.price)
    settings = ModelSettings(mid_year=args.mid_year, include_nonoperating=not args.simple_bridge)

    facts = CompanyFacts.download() if args.refresh else CompanyFacts.from_path()
    result = run_model(market, settings, facts)
    path, cache = write_model_workbook(result, OUTPUT_DIR / args.output)
    export_results(result)

    r, f = result, result.fundamentals
    print("\nCOCA-COLA REVERSE DCF  (USD millions)")
    print("=" * 60)
    print(f"Valuation date / SEC basis : {market.valuation_date} / {r.snapshot.basis_label}")
    print(f"Share price x diluted shs  : ${market.share_price:,.2f} x {r.diluted_shares:,.1f}m = {r.market_cap:,.0f}")
    print(f"Market enterprise value    : {r.market_ev:,.0f}")
    print(f"WACC                       : {r.wacc.wacc:.4%}   (Ke {r.wacc.cost_of_equity:.2%}, E/V {r.wacc.equity_weight:.1%})")
    print(f"Base-year FCFF (TTM)       : {r.base['fcff']:,.0f}   (tax rate {r.base['tax_rate']:.2%})")
    print(f"Terminal growth            : {market.terminal_growth_rate:.2%}")
    print(f"MARKET-IMPLIED FCFF GROWTH : {r.implied_growth:.4%} a year for {settings.forecast_years} years")
    print(f"  residual (USD m)         : {r.implied['valuation_difference']:.2e}")
    print(f"  terminal value % of EV   : {r.schedule.terminal_share_of_ev:.1%}")
    print("Plausibility benchmarks    :")
    print(f"  revenue CAGR 3/5/6y      : " + " / ".join(f"{f.revenue_cagr[n]:.2%}" for n in (3, 5, 6)))
    print(f"  EBIT CAGR 3/5/6y         : " + " / ".join(f"{f.ebit_cagr[n]:.2%}" for n in (3, 5, 6)))
    print(f"  ROIC {f.roic:.1%} x reinvestment {f.reinvestment_rate_normalised:.1%} = fundamental growth {f.fundamental_growth_normalised:.2%}")
    print(f"Excel model                : {path}" + ("" if cache else "  (LibreOffice not found: no cached values)"))


if __name__ == "__main__":
    main()
