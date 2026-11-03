
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parents[2]
DATA_DIR = ROOT / "data"

FY2025_PATH = DATA_DIR / "extracted" / "fy2025_financials.csv"
Q3_2026_PATH = DATA_DIR / "q3_financials.csv"
OUTPUT_PATH = DATA_DIR / "ttm_financials.csv"

ADDITITIVE_METRICS = {
    "revenue",
    "ebit",
    "income_before_income_taxes",
    "income_tax_expense",
    "depreciation_amortization",
    "capital_expenditures_reported",
    "net_change_operating_assets_liabilities",
}


def main():
    for path in [FY2025_PATH, Q3_2026_PATH]:
        if not path.exists():
            raise FileNotFoundError(f"Required input file not found: {path}")

    annual = pd.read_csv(FY2025_PATH)
    interim = pd.read_csv(Q3_2026_PATH)

    required_columns = {
        "metric", "period", "value_usd_millions"
    }

    for name, frame in [("FY2025", annual), ("Q3 2026", interim)]:
        missing = required_columns - set(frame.columns)
        if missing:
            raise ValueError(f"{name} dataset is missing columns: {missing}")

    annual_values = annual.set_index("metric")["value_usd_millions"]
    interim_values = interim.pivot(
        index="metric",
        columns="period",
        values="value_usd_millions",
    )

    required_interim_periods = {"9M 2026", "9M 2025"}
    missing_periods = required_interim_periods - set(interim_values.columns)
    if missing_periods:
        raise ValueError(
            f"Q3 dataset is missing periods: {missing_periods}"
        )

    records = []

    for metric in sorted(ADDITITIVE_METRICS):
        if metric not in annual_values.index:
            raise ValueError(f"FY2025 data is missing metric: {metric}")
        if metric not in interim_values.index:
            raise ValueError(f"Q3 dataset is missing metric: {metric}")

        fy2025 = float(annual_values.loc[metric])
        nine_months_2026 = float(interim_values.loc[metric, "9M 2026"])
        nine_months_2025 = float(interim_values.loc[metric, "9M 2025"])

        ttm = fy2025 + nine_months_2026 - nine_months_2025

        records.append({
            "metric": metric,
            "period": "TTM through Q3 2026",
            "value_usd_millions": ttm,
            "fy2025": fy2025,
            "9m_2026": nine_months_2026,
            "9m_2025": nine_months_2025,
            "calculation": "FY2025 + 9M 2026 - 9M 2025",
            "source": "Coca-Cola FY2025 Form 10-K and Q3 2026 Form 10-Q",
        })

    result = pd.DataFrame(records)

    # Effective tax rate based on the reported income-tax expense and
    # income before income taxes, for reference only.
    pretax = result.loc[
        result["metric"] == "income_before_income_taxes",
        "value_usd_millions",
    ].iloc[0]
    tax = result.loc[
        result["metric"] == "income_tax_expense",
        "value_usd_millions",
    ].iloc[0]

    if pretax != 0:
        effective_tax_rate = tax / pretax
        print(f"TTM reported effective tax rate: {effective_tax_rate:.2%}")

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    result.to_csv(OUTPUT_PATH, index=False)

    print("\nTTM financials ($ millions):")
    print(result.to_string(index=False))
    print(f"\nSaved to: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
