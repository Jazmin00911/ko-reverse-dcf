import os, sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from src.reverse_dcf.sec import CompanyFacts

import requests

SEC_SUBMISSIONS_URL = (
    "https://data.sec.gov/submissions/CIK0000021344.json"
)


def check_sec_filings():
    user_agent = os.getenv("SEC_USER_AGENT")
    if not user_agent:
        raise ValueError("SEC_USER_AGENT is not set.")

    response = requests.get(
        SEC_SUBMISSIONS_URL,
        headers={"User-Agent": user_agent},
        timeout=30,
    )
    response.raise_for_status()
    data = response.json()

    recent = data["filings"]["recent"]

    print("\nLatest SEC filings:")
    for i, form in enumerate(recent["form"]):
        if form in {"10-Q", "10-K"}:
            filing_date = recent["filingDate"][i]
            report_date = recent["reportDate"][i]

            print(f"{filing_date} | {form} | {report_date}")

            print("   accession:", recent["accessionNumber"][i], "| document:", recent["primaryDocument"][i])



def main():
    print("Downloading Coca-Cola company facts from the SEC...")

    facts = CompanyFacts.download(save=True).raw
    us_gaap = facts.get("facts", {}).get("us-gaap", {})

    latest_filing = max(
        (
            item.get("filed", "")
            for concept in us_gaap.values()
            for entries in concept.get("units", {}).values()
            for item in entries
            if item.get("filed")
        ),
        default="No filing dates found",
    )

    print("Latest filing date in SEC response:", latest_filing)
    print("Download successful: data/raw/coca_cola_companyfacts.json")
    print(f"Company: {facts.get('entityName', 'Unknown')}")
    print(f"XBRL namespaces: {len(facts.get('facts', {}))}")
    check_sec_filings()


if __name__ == "__main__":
    main()
