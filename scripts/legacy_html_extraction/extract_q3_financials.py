
from pathlib import Path
import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[2]s[2]

FILING_PATH = PROJECT_ROOT / "data" / "raw" / "ko-20261002.htm"
OUTPUT_DIR = PROJECT_ROOT / "data" / "extracted"

# These are the statement-table positions used in Q2 extraction.
# verifying the Q3 table positions before relying on them.
TABLE_NUMBERS = {
    "income_statement": 5,
    "balance_sheet": 7,
    "cash_flow_statement": 8,
}


def main():
    if not FILING_PATH.exists():
        raise FileNotFoundError(
            f"Q3 filing not found: {FILING_PATH}"
        )

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    print("Reading Coca-Cola's Q3 2026 SEC filing...")
    tables = pd.read_html(FILING_PATH)

    print(f"Found {len(tables)} HTML tables.")

    for statement, table_number in TABLE_NUMBERS.items():
        if table_number >= len(tables):
            raise IndexError(
                f"Table {table_number} does not exist. "
                f"Only {len(tables)} tables were found."
            )

        table = tables[table_number]
        output_path = (
            OUTPUT_DIR / f"q3_{statement}.csv"
        )

        table.to_csv(
            output_path,
            index=False,
            header=False,
        )

        print(
            f"Saved {statement}: {output_path} "
            f"({table.shape[0]} rows, {table.shape[1]} columns)"
        )


if __name__ == "__main__":
    main()
