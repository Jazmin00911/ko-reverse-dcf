"""Point-in-time extraction of Coca-Cola's financials from SEC XBRL company facts.

Data source: https://data.sec.gov/api/xbrl/companyfacts/CIK0000021344.json
(the same SEC API the original project used).

Look-ahead protection
---------------------
Every XBRL fact carries a ``filed`` date.  All selection helpers accept an
``asof`` date and only consider facts with ``filed <= asof``; when a figure has
been restated, the version that was public on ``asof`` is used.  The final
snapshot re-validates every fact it used (``assert_no_lookahead``).
"""
from __future__ import annotations
import json
import os
from dataclasses import dataclass, field
from datetime import date, timedelta
from pathlib import Path
from typing import Iterable
import requests
from config import DATA_DIR

SEC_CIK = "0000021344"
SEC_FACTS_URL = f"https://data.sec.gov/api/xbrl/companyfacts/CIK{SEC_CIK}.json"
RAW_FACTS_PATH = DATA_DIR / "raw" / "coca_cola_companyfacts.json"
USD_PER_MILLION = 1_000_000.0

# metric -> us-gaap tag (flow items, summed to trailing-twelve-months)
FLOW_TAGS: dict[str, str] = {
    "revenue": "Revenues",
    "ebit": "OperatingIncomeLoss",
    "income_before_income_taxes":
        "IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
    "income_tax_expense": "IncomeTaxExpenseBenefit",
    "depreciation_amortization": "DepreciationDepletionAndAmortization",
    "capital_expenditures": "PaymentsToAcquirePropertyPlantAndEquipment",
    # XBRL sign: positive = operating assets/liabilities *absorbed* cash,
    # i.e. an increase in net working capital -> reduces FCFF.
    "change_in_nwc": "IncreaseDecreaseInOperatingCapital",
}

# balance-sheet label -> us-gaap tag (instant items at the latest period end)
BALANCE_TAGS: dict[str, str] = {
    "cash_and_equivalents": "CashAndCashEquivalentsAtCarryingValue",
    "short_term_investments": "OtherShortTermInvestments",
    "marketable_securities": "MarketableSecurities",
    "loans_and_notes_payable": "NotesAndLoansPayable",
    "current_maturities_of_debt": "LongTermDebtAndCapitalLeaseObligationsCurrent",
    "long_term_debt": "LongTermDebtAndCapitalLeaseObligations",
    "noncontrolling_interests": "MinorityInterest",
    "equity_method_investments": "EquityMethodInvestments",
    "total_equity": "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest",
}

HISTORY_TAGS: dict[str, str] = {
    "revenue": "Revenues",
    "ebit": "OperatingIncomeLoss",
    "income_before_income_taxes": FLOW_TAGS["income_before_income_taxes"],
    "income_tax_expense": "IncomeTaxExpenseBenefit",
    "depreciation_amortization": "DepreciationDepletionAndAmortization",
    "capital_expenditures": "PaymentsToAcquirePropertyPlantAndEquipment",
    "change_in_nwc": "IncreaseDecreaseInOperatingCapital",
}


class LookAheadError(RuntimeError):
    """Raised when a fact filed after the valuation date would be used."""


class SecDataError(RuntimeError):
    """Raised when a required XBRL fact cannot be found as of the date."""


@dataclass(frozen=True)
class Fact:
    tag: str
    value: float            # raw units (USD or shares)
    start: date | None
    end: date
    filed: date
    form: str
    accn: str

    @property
    def millions(self) -> float:
        return self.value / USD_PER_MILLION

    @property
    def days(self) -> int | None:
        return (self.end - self.start).days if self.start else None

    def cite(self) -> str:
        return f"{self.form} filed {self.filed.isoformat()} ({self.accn}) | us-gaap:{self.tag}"


def _d(text: str) -> date:
    return date.fromisoformat(text)


class CompanyFacts:
    """Thin, read-only wrapper over the SEC companyfacts payload."""

    def __init__(self, raw: dict):
        self.raw = raw
        self._us = raw["facts"]["us-gaap"]
        self._dei = raw["facts"].get("dei", {})

    # ------------------------------------------------------------------ io
    @classmethod
    def from_path(cls, path: Path | None = None) -> "CompanyFacts":
        path = Path(path) if path else RAW_FACTS_PATH
        if not path.exists():
            raise FileNotFoundError(
                f"{path} not found. Run `python scripts/download_sec_data.py` "
                "(requires the SEC_USER_AGENT environment variable)."
            )
        with path.open(encoding="utf-8") as handle:
            return cls(json.load(handle))

    @classmethod
    def download(cls, save: bool = True) -> "CompanyFacts":
        user_agent = os.getenv("SEC_USER_AGENT")
        if not user_agent:
            raise ValueError(
                "Set the SEC_USER_AGENT environment variable to identify your "
                "application and provide contact information (SEC fair-access policy)."
            )
        response = requests.get(SEC_FACTS_URL, headers={"User-Agent": user_agent}, timeout=30)
        response.raise_for_status()
        payload = response.json()
        if save:
            RAW_FACTS_PATH.parent.mkdir(parents=True, exist_ok=True)
            RAW_FACTS_PATH.write_text(json.dumps(payload), encoding="utf-8")
        return cls(payload)

    # ------------------------------------------------------------- queries
    def _facts(self, tag: str, namespace: dict | None = None, unit: str = "USD") -> list[Fact]:
        concept = (namespace if namespace is not None else self._us).get(tag)
        if not concept:
            return []
        out = []
        for entry in concept["units"].get(unit, []):
            out.append(Fact(
                tag=tag, value=float(entry["val"]),
                start=_d(entry["start"]) if entry.get("start") else None,
                end=_d(entry["end"]), filed=_d(entry["filed"]),
                form=entry.get("form", ""), accn=entry.get("accn", ""),
            ))
        return out

    @staticmethod
    def _latest_filed(facts: Iterable[Fact], asof: date) -> Fact | None:
        eligible = [f for f in facts if f.filed <= asof]
        return max(eligible, key=lambda f: f.filed) if eligible else None

    def duration(self, tag: str, start: date, end: date, asof: date) -> Fact:
        matches = [f for f in self._facts(tag) if f.start == start and f.end == end]
        fact = self._latest_filed(matches, asof)
        if fact is None:
            raise SecDataError(f"{tag} for {start}..{end} not available as of {asof}.")
        return fact

    def instant(self, tag: str, end: date, asof: date) -> Fact:
        matches = [f for f in self._facts(tag) if f.start is None and f.end == end]
        fact = self._latest_filed(matches, asof)
        if fact is None:
            raise SecDataError(f"{tag} at {end} not available as of {asof}.")
        return fact

    def latest_instant(self, tag: str, asof: date) -> Fact:
        facts = [f for f in self._facts(tag) if f.start is None and f.filed <= asof]
        if not facts:
            raise SecDataError(f"{tag} not available as of {asof}.")
        return max(facts, key=lambda f: (f.end, f.filed))

    def cover_shares(self, asof: date) -> Fact:
        facts = [f for f in self._facts("EntityCommonStockSharesOutstanding", self._dei, "shares")
                 if f.filed <= asof]
        if not facts:
            raise SecDataError(f"Cover-page share count not available as of {asof}.")
        return max(facts, key=lambda f: (f.filed, f.end))

    def annual_series(self, tag: str, asof: date) -> dict[int, Fact]:
        """Fiscal-year duration facts from 10-Ks, keyed by fiscal-year-end year."""
        best: dict[date, Fact] = {}
        for f in self._facts(tag):
            if f.form not in ("10-K", "10-K/A") or f.days is None or not 350 <= f.days <= 380:
                continue
            if f.filed > asof:
                continue
            if f.end not in best or f.filed > best[f.end].filed:
                best[f.end] = f
        return {end.year: fact for end, fact in sorted(best.items())}


# --------------------------------------------------------------------------- #
# Snapshot dataclasses
# --------------------------------------------------------------------------- #
@dataclass(frozen=True)
class TtmItem:
    """TTM = last fiscal year + current YTD - prior-year YTD (all USD millions)."""
    metric: str
    fiscal_year: float
    ytd_current: float
    ytd_prior: float
    sources: tuple[str, ...]

    @property
    def ttm(self) -> float:
        return self.fiscal_year + self.ytd_current - self.ytd_prior


@dataclass(frozen=True)
class BalanceItem:
    metric: str
    value: float            # USD millions
    as_of: date
    source: str


@dataclass(frozen=True)
class SecSnapshot:
    valuation_date: date
    basis_label: str
    period_end: date
    ytd_start: date
    ytd_prior_end: date
    fiscal_year_end: date
    filing_date: date
    ttm: dict[str, TtmItem]
    balance: dict[str, BalanceItem]
    operating_lease_liability: BalanceItem
    shares_outstanding_m: float
    dilutive_shares_m: float
    shares_source: str
    facts_used: tuple[Fact, ...] = field(repr=False, default=())

    @property
    def diluted_shares_m(self) -> float:
        return self.shares_outstanding_m + self.dilutive_shares_m

    @property
    def total_debt(self) -> float:
        b = self.balance
        return (b["loans_and_notes_payable"].value
                + b["current_maturities_of_debt"].value
                + b["long_term_debt"].value)

    def ttm_value(self, metric: str) -> float:
        return self.ttm[metric].ttm

    def assert_no_lookahead(self) -> None:
        late = [f for f in self.facts_used if f.filed > self.valuation_date]
        if late:
            raise LookAheadError(
                "Facts filed after the valuation date were used: "
                + "; ".join(f.cite() for f in late)
            )


# --------------------------------------------------------------------------- #
# Snapshot construction
# --------------------------------------------------------------------------- #
def _current_ytd_period(facts: CompanyFacts, asof: date) -> Fact:
    """Latest operating-income fact (longest fiscal-year-to-date duration)."""
    candidates = [
        f for f in facts._facts(FLOW_TAGS["ebit"])
        if f.filed <= asof and f.form in ("10-Q", "10-K") and f.days and f.days <= 380
        # fiscal-year-to-date facts start at (or within a week of) the year start
    ]
    if not candidates:
        raise SecDataError(f"No operating-income facts filed on/before {asof}.")
    latest_end = max(f.end for f in candidates)
    same_end = [f for f in candidates if f.end == latest_end]
    return max(same_end, key=lambda f: (f.days, f.filed))


def _basis_label(period: Fact) -> str:
    quarter = {3: "Q1", 6: "Q2", 9: "Q3", 12: "FY"}
    # fiscal quarters end a few days off calendar quarter-ends
    months = round(period.days / 30.4)
    q = quarter.get(months if months in quarter else 12 if months >= 11 else months, "")
    return f"{q} {period.end.year} {'10-K' if period.form == '10-K' else '10-Q'} (TTM to {period.end.isoformat()})"


def build_snapshot(facts: CompanyFacts, valuation_date: date) -> SecSnapshot:
    """Assemble every SEC-derived model input using only facts public on ``valuation_date``."""
    asof = valuation_date
    current = _current_ytd_period(facts, asof)
    used: list[Fact] = []

    is_full_year = current.days >= 350
    ytd_start, period_end = current.start, current.end

    # prior fiscal year (ends the day before the YTD window starts)
    fy_candidates = [
        f for f in facts._facts(FLOW_TAGS["ebit"])
        if f.filed <= asof and f.days and 350 <= f.days <= 380
        and abs((f.end - (ytd_start - timedelta(days=1))).days) <= 7
    ]
    if is_full_year:
        fy_end = period_end
    else:
        if not fy_candidates:
            raise SecDataError("Prior fiscal-year operating income not found.")
        fy_end = max(fy_candidates, key=lambda f: f.filed).end
    fy_start = max(f for f in fy_candidates if f.end == fy_end).start if not is_full_year else ytd_start

    # prior-year comparable YTD window
    prior_candidates = [
        f for f in facts._facts(FLOW_TAGS["ebit"])
        if f.filed <= asof and f.days and abs(f.days - current.days) <= 10
        and abs((f.start - (ytd_start - timedelta(days=365))).days) <= 7
        and abs((f.end - (period_end - timedelta(days=365))).days) <= 10
    ]
    if not is_full_year and not prior_candidates:
        raise SecDataError("Prior-year YTD window not found for the current period.")
    prior = max(prior_candidates, key=lambda f: f.filed) if prior_candidates else None

    ttm_items: dict[str, TtmItem] = {}
    for metric, tag in FLOW_TAGS.items():
        fy = facts.duration(tag, fy_start, fy_end, asof)
        if is_full_year:
            cur_v = prior_v = 0.0
            cites = (fy.cite(),)
            used.append(fy)
        else:
            cur = facts.duration(tag, ytd_start, period_end, asof)
            pri = facts.duration(tag, prior.start, prior.end, asof)
            cur_v, prior_v = cur.millions, pri.millions
            cites = (fy.cite(), cur.cite(), pri.cite())
            used.extend([fy, cur, pri])
        ttm_items[metric] = TtmItem(metric, fy.millions, cur_v, prior_v, cites)

    balance: dict[str, BalanceItem] = {}
    for label, tag in BALANCE_TAGS.items():
        fact = facts.instant(tag, period_end, asof)
        used.append(fact)
        balance[label] = BalanceItem(label, fact.millions, fact.end, fact.cite())

    lease = facts.latest_instant("OperatingLeaseLiability", asof)
    used.append(lease)
    lease_item = BalanceItem("operating_lease_liability", lease.millions, lease.end, lease.cite())

    cover = facts.cover_shares(asof)
    used.append(cover)
    # Dilution from the latest 3-month (or YTD) weighted-average share counts.
    dilutive = 0.0
    dil_src = "n/a (cover-page shares only)"
    for window_days in (range(80, 100), range(150, 200), range(250, 290)):
        diluted = [f for f in facts._facts("WeightedAverageNumberOfDilutedSharesOutstanding", unit="shares")
                   if f.end == period_end and f.filed <= asof and f.days in window_days]
        basic = [f for f in facts._facts("WeightedAverageNumberOfSharesOutstandingBasic", unit="shares")
                 if f.end == period_end and f.filed <= asof and f.days in window_days]
        if diluted and basic:
            d_f, b_f = max(diluted, key=lambda f: f.filed), max(basic, key=lambda f: f.filed)
            dilutive = (d_f.value - b_f.value) / USD_PER_MILLION
            used.extend([d_f, b_f])
            dil_src = f"{d_f.form} filed {d_f.filed.isoformat()} ({d_f.accn}) | diluted - basic weighted avg"
            break

    snapshot = SecSnapshot(
        valuation_date=valuation_date,
        basis_label=_basis_label(current),
        period_end=period_end,
        ytd_start=ytd_start,
        ytd_prior_end=prior.end if prior else period_end,
        fiscal_year_end=fy_end,
        filing_date=max(f.filed for f in used),
        ttm=ttm_items,
        balance=balance,
        operating_lease_liability=lease_item,
        shares_outstanding_m=cover.value / USD_PER_MILLION,
        dilutive_shares_m=dilutive,
        shares_source=f"{cover.cite()} (cover page, as of {cover.end.isoformat()}); dilution: {dil_src}",
        facts_used=tuple(used),
    )
    snapshot.assert_no_lookahead()
    return snapshot


def annual_history(facts: CompanyFacts, asof: date, first_year: int) -> dict[str, dict[int, float]]:
    """Fiscal-year series (USD millions) per metric from ``first_year`` onward."""
    history: dict[str, dict[int, float]] = {}
    for metric, tag in HISTORY_TAGS.items():
        series = facts.annual_series(tag, asof)
        history[metric] = {y: f.millions for y, f in series.items() if y >= first_year}
    years = set.intersection(*(set(v) for v in history.values()))
    return {m: {y: v[y] for y in sorted(years)} for m, v in history.items()}
