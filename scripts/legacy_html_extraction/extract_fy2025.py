from pathlib import Path
import re
import pandas as pd

ROOT = Path(__file__).resolve().parents[2]
FILING_PATH = ROOT / "data" / "raw" / "ko-2025-10k.htm"
OUTPUT_PATH = ROOT / "data" / "extracted" / "fy2025_financials.csv"

SOURCE = "Coca-Cola FY2025 Form 10-K"


INCOME_STATEMENT_TABLE = 23
BALANCE_SHEET_TABLE = 25
CASH_FLOW_TABLE = 26


def normalize(value):
    text = str(value).replace("\xa0", " ").strip().casefold()
    return re.sub(r"\s+", " ", text)


def parse_number(value):
    text = str(value).strip()
    text = text.replace("\xa0", "").replace(",", "").replace("$", "")
    text = text.replace("−", "-")

    if not text or text.casefold() in {"nan", "none", "-", "—", "–"}:
        return None

    if text.startswith("(") and text.endswith(")"):
        text = "-" + text[1:-1].strip()

    try:
        return float(text)
    except ValueError:
        return None


def extract_row_value(table, label, year_index=0):
    target = normalize(label)
    matches = []

    for _, row in table.iterrows():
        cells = [str(value) for value in row.tolist()]

        if not any(normalize(cell) == target for cell in cells[:3]):
            continue

        # SEC tables repeat labels and sometimes repeat numeric values.
        values = []
        for cell in cells[3:]:
            number = parse_number(cell)
            if number is not None and (
                not values or number != values[-1]
            ):
                values.append(number)

        if len(values) > year_index:
            matches.append(values[year_index])

    if len(matches) != 1:
        raise ValueError(
            f"Expected one match for '{label}', found {matches}"
        )

    return matches[0]


def main():
    if not FILING_PATH.exists():
        raise FileNotFoundError(f"Filing not found: {FILING_PATH}")

    tables = pd.read_html(FILING_PATH)
    print(f"Read {len(tables)} HTML tables.")

    income = tables[INCOME_STATEMENT_TABLE]
    cash_flow = tables[CASH_FLOW_TABLE]

    specifications = [
        (income, "Net Operating Revenues", "revenue"),
        (income, "Operating Income", "ebit"),
        (income, "Income before income taxes", "income_before_income_taxes"),
        (income, "Income taxes", "income_tax_expense"),
        (cash_flow, "Depreciation and amortization", "depreciation_amortization"),
        (
            cash_flow,
            "Purchases of property, plant and equipment",
            "capital_expenditures_reported",
        ),
        (
            cash_flow,
            "Net change in operating assets and liabilities",
            "net_change_operating_assets_liabilities",
        ),
    ]

    records = []

    for table, label, metric in specifications:
        value = extract_row_value(table, label)

        records.append({
            "metric": metric,
            "period": "FY2025",
            "value_usd_millions": value,
            "source": SOURCE,
            "source_label": label,
        })

    result = pd.DataFrame(records)

    capex_mask = result["metric"] == "capital_expenditures_reported"
    result.loc[capex_mask, "value_usd_millions"] = (
        result.loc[capex_mask, "value_usd_millions"].abs()
    )


    expected = {
        "revenue": 47941.0,
        "ebit": 13762.0,
    }

    for metric, expected_value in expected.items():
        actual = result.loc[
            result["metric"] == metric, "value_usd_millions"
        ].iloc[0]

        if actual != expected_value:
            raise ValueError(
                f"Unexpected {metric}: {actual}; expected {expected_value}"
            )

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    result.to_csv(OUTPUT_PATH, index=False)

    print("\nExtracted FY2025 financials ($ millions):")
    print(result.to_string(index=False))
    print(f"\nSaved to: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
