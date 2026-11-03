from datetime import date

import pandas as pd
import pytest

from src.reverse_dcf.config import DATA_DIR
from src.reverse_dcf.sec import LookAheadError, build_snapshot


def test_oct23_uses_q2_not_q3(facts):
    snap = build_snapshot(facts, date(2026, 10, 23))
    assert snap.period_end == date(2026, 7, 3)          # Q3 10-Q was filed 2026-10-28
    assert snap.filing_date <= date(2026, 10, 23)
    snap.assert_no_lookahead()


def test_q3_available_after_filing(facts):
    snap = build_snapshot(facts, date(2026, 10, 28))
    assert snap.period_end == date(2026, 10, 2)


def test_lookahead_guard_triggers(facts):
    snap = build_snapshot(facts, date(2026, 10, 28))
    from dataclasses import replace
    early = replace(snap, valuation_date=date(2026, 10, 23))
    with pytest.raises(LookAheadError):
        early.assert_no_lookahead()


def test_companyfacts_ttm_matches_html_extraction(facts):
    """XBRL TTM (Q3 basis) must equal the figures scraped from the 10-K/10-Q HTML."""
    legacy = pd.read_csv(DATA_DIR / "ttm_financials.csv").set_index("metric")["value_usd_millions"]
    snap = build_snapshot(facts, date(2026, 10, 28))
    mapping = {"revenue": "revenue", "ebit": "ebit", "income_before_income_taxes": "income_before_income_taxes",
               "income_tax_expense": "income_tax_expense", "depreciation_amortization": "depreciation_amortization",
               "capital_expenditures": "capital_expenditures_reported", "change_in_nwc": "net_change_operating_assets_liabilities"}
    for metric, legacy_name in mapping.items():
        expected = legacy[legacy_name] * (-1 if metric == "change_in_nwc" else 1)
        assert snap.ttm_value(metric) == pytest.approx(expected), metric
