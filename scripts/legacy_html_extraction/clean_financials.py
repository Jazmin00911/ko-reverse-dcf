from pathlib import Path
import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[2]s[2]
INPUT_DIR = PROJECT_ROOT / "data" / "extracted"
OUTPUT_DIR = PROJECT_ROOT / "data" / "cleaned"

FILES = {
    "income_statement": "q2_income_statement.csv",
    "balance_sheet": "q2_balance_sheet.csv",
    "cash_flow_statement": "q2_cash_flow_statement.csv",
}


def clean_label(value):
    if pd.isna(value):
        return ""
    return " ".join(str(value).replace("\xa0", " ").split())


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    for statement, filename in FILES.items():
        input_path = INPUT_DIR / filename
        if not input_path.exists():
            raise FileNotFoundError(f"Missing input file: {input_path}")

        df = pd.read_csv(input_path, header=None, dtype=str).fillna("")
        df.insert(
            0,
            "source_row",
            range(1, len(df) + 1),
        )

        # Preserve every original column, but normalize text whitespace.
        for column in df.columns:
            df[column] = df[column].map(clean_label)

        output_path = OUTPUT_DIR / f"{statement}_labeled.csv"
        df.to_csv(output_path, index=False)

        print(f"Saved {statement}: {output_path}")
        print(f"Rows: {len(df)} | Columns: {len(df.columns)}")

        # Show only rows likely to be useful for FCFF.
        keywords = (
            "revenue",
            "operating income",
            "depreciation",
            "capital expenditures",
            "property, plant",
            "working capital",
            "receivables",
            "inventories",
            "accounts payable",
            "cash provided by operating activities",
        )

        matches = df[
            df.iloc[:, 1:].apply(
                lambda row: any(
                    keyword in " ".join(row.astype(str)).lower()
                    for keyword in keywords
                ),
                axis=1,
            )
        ]

        print(f"Potentially relevant rows: {len(matches)}")


if __name__ == "__main__":
    main()
