import os
from pathlib import Path
import requests

ROOT = Path(__file__).resolve().parents[2]
RAW_DIR = ROOT / "data" / "raw"

CIK = "21344"

FILINGS = {
    "q3_2026": {
        "accession": "0001628280-26-067844",
        "primary_document": "ko-20261002.htm",
        "output_filename": "ko-20261002.htm",
    },
    "fy2025": {
        "accession": "0001628280-26-010047",
        "primary_document": "ko-20251231.htm",
        "output_filename": "ko-2025-10k.htm",
    },
}


def download_filing(filing_name: str) -> Path:
    if filing_name not in FILINGS:
        raise ValueError(
            f"Unknown filing: {filing_name}. "
            f"Choose one of: {', '.join(FILINGS)}"
        )

    user_agent = os.getenv("SEC_USER_AGENT")
    if not user_agent:
        raise ValueError(
            "Set SEC_USER_AGENT in your PyCharm run configuration."
        )

    filing = FILINGS[filing_name]
    accession_path = filing["accession"].replace("-", "")

    url = (
        f"https://www.sec.gov/Archives/edgar/data/{CIK}/"
        f"{accession_path}/{filing['primary_document']}"
    )

    RAW_DIR.mkdir(parents=True, exist_ok=True)
    output_path = RAW_DIR / filing["output_filename"]

    response = requests.get(
        url,
        headers={"User-Agent": user_agent},
        timeout=30,
    )
    response.raise_for_status()

    if "<html" not in response.text[:2000].lower():
        raise ValueError("The SEC response does not appear to be an HTML filing.")

    output_path.write_text(response.text, encoding="utf-8")

    print(f"Downloaded {filing_name}: {output_path}")
    print(f"File size: {output_path.stat().st_size:,} bytes")

    return output_path


if __name__ == "__main__":
    print("Available filings:")
    print("  1. q3_2026")
    print("  2. fy2025")

    choice = input("Which filing should be downloaded? Enter 1 or 2: ").strip()

    filing_name = {
        "1": "q3_2026",
        "2": "fy2025",
    }.get(choice)

    if filing_name is None:
        raise ValueError("Invalid choice. Enter 1 or 2.")

    download_filing(filing_name)