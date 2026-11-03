
from pathlib import Path
import re
import pandas as pd

ROOT = Path(__file__).resolve().parents[2]
EXTRACTED = ROOT / "data" / "extracted"
OUTPUT = ROOT / "data" / "q3_financials.csv"


def parse_number(value):
    text = str(value).strip().replace(",", "").replace("$", "")

    if not text or text in {"—", "-", "nan", "None"}:
        return None

    if text.startswith("(") and text.endswith(")"):
        return -float(text[1:-1])

    try:
        return float(text)
    except ValueError:
        return None


def normalize_label(value):
    text = str(value).replace("\xa0", " ").strip()
    return re.sub(r"\s+", " ", text).casefold()


def extract_values(df, label):
    matches = df[
        df.iloc[:, 0].apply(normalize_label)
        == normalize_label(label)
    ]

    if len(matches) != 1:
        raise ValueError(
            f"Expected one row for '{label}', found {len(matches)}"
        )

    row = matches.iloc[0]

    # Ignore the first three repeated label columns.
    values = []
    for cell in row.iloc[3:]:
        number = parse_number(cell)
        if number is not None:
            if not values or number != values[-1]:
                values.append(number)

    return values


def main():
    income = pd.read_csv(
        EXTRACTED / "q3_income_statement.csv",
        header=None,
        dtype=str,
    ).fillna("")

    cash_flow = pd.read_csv(
        EXTRACTED / "q3_cash_flow_statement.csv",
        header=None,
        dtype=str,
    ).fillna("")

    specifications = [
        (income, "Net Operating Revenues", "revenue",
         ["Q3 2026", "Q3 2025", "9M 2026", "9M 2025"]),
        (income, "Operating Income", "ebit",
         ["Q3 2026", "Q3 2025", "9M 2026", "9M 2025"]),
        (income, "Income taxes", "income_tax_expense",
         ["Q3 2026", "Q3 2025", "9M 2026", "9M 2025"]),
        (income, "Income before income taxes", "income_before_income_taxes",
         ["Q3 2026", "Q3 2025", "9M 2026", "9M 2025"]),
        (cash_flow, "Depreciation and amortization",
         "depreciation_amortization",
         ["9M 2026", "9M 2025"]),
        (cash_flow, "Purchases of property, plant and equipment",
         "capital_expenditures_reported",
         ["9M 2026", "9M 2025"]),
        (cash_flow, "Net change in operating assets and liabilities",
         "net_change_operating_assets_liabilities",
         ["9M 2026", "9M 2025"]),
    ]

    records = []

    for df, source_label, metric, periods in specifications:
        values = extract_values(df, source_label)

        if len(values) != len(periods):
            raise ValueError(
                f"{source_label}: expected {len(periods)} values "
                f"for {periods}, found {values}"
            )

        for period, value in zip(periods, values):
            records.append({
                "metric": metric,
                "period": period,
                "value_usd_millions": value,
                "source": "Coca-Cola Q3 2026 Form 10-Q",
                "source_label": source_label,
            })

    result = pd.DataFrame(records)

    # Present capital expenditures as a positive outflow.
    capex_mask = result["metric"] == "capital_expenditures_reported"
    result.loc[capex_mask, "value_usd_millions"] = (
        result.loc[capex_mask, "value_usd_millions"].abs()
    )

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    result.to_csv(OUTPUT, index=False)

    print(result.to_string(index=False))
    print(f"\nSaved verified candidate dataset to: {OUTPUT}")


if __name__ == "__main__":
    main()
