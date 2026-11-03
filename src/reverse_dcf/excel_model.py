"""Institutional-format Excel workbook (live formulas) for the reverse DCF.

Design rules
------------
* Every number is either a **blue hard-code** (inputs, SEC data, Python parity
  values) or a **black formula**.  No pasted results drive the valuation: the
  implied growth rate is solved *inside* Excel by a 60-step bisection on the
  closed-form EV (sheet "Solver"), so the model re-solves when any input changes.
* USD millions, Arial 10, gridlines off, consistent number formats, dynamic
  date headers (EDATE off the SEC period end).
* ``Model Checks`` reconciles the live Excel results with the Python engine.
"""
from __future__ import annotations

import datetime as dt
import shutil
import subprocess
import tempfile
from pathlib import Path

import xlsxwriter

from .fundamentals import HISTORY_FIRST_YEAR, NORMALISED_WINDOW, REVENUE_SCENARIOS
from .model import ModelResult

# --------------------------------------------------------------------------- #
# Sheet names and fixed layout (rows are 1-indexed Excel rows)
# --------------------------------------------------------------------------- #
S_SUM, S_INP, S_SEC, S_HIS = "Summary", "Inputs", "SEC Data", "Historical"
S_DCF, S_SOL, S_SEN, S_GRD, S_CHK, S_MET = ("DCF Engine", "Solver", "Sensitivity",
                                            "Solver Grid", "Model Checks", "Methodology")
SHEETS = [S_SUM, S_INP, S_SEC, S_HIS, S_DCF, S_SOL, S_SEN, S_GRD, S_CHK, S_MET]

NAVY, MID_BLUE, LIGHT_BLUE, GREY = "#1F3864", "#2F5597", "#D9E1F2", "#F2F2F2"
BLUE_INPUT = "#0000FF"          # pure blue: hard-coded inputs / historical data
FILL_INPUT = "#FFF9E5"

NF_MM = '#,##0_);(#,##0);"–"_)'
NF_MM1 = '#,##0.0_);(#,##0.0);"–"_)'
NF_PCT2, NF_PCT1 = "0.00%;(0.00%);0.00%", "0.0%;(0.0%);0.0%"
NF_USD = '$#,##0.00_);($#,##0.00)'
NF_DATE, NF_FACTOR, NF_X = "d-mmm-yy", "0.0000", '0.0"x"'
NF_INT, NF_SCI = "0", "0.0E+00"
NF_FY = '"FY"0'

YEARS = list(range(HISTORY_FIRST_YEAR, HISTORY_FIRST_YEAR + 7))
HIS_COLS = "CDEFGHI"                                  # FY2019..FY2025
FLOW_ORDER = ["revenue", "ebit", "income_before_income_taxes", "income_tax_expense",
              "depreciation_amortization", "capital_expenditures", "change_in_nwc"]
FLOW_LABEL = {
    "revenue": "Net operating revenues", "ebit": "Operating income (EBIT)",
    "income_before_income_taxes": "Income before income taxes",
    "income_tax_expense": "Income taxes", "depreciation_amortization": "Depreciation & amortization",
    "capital_expenditures": "Purchases of PP&E (CapEx)",
    "change_in_nwc": "Net change in operating assets & liabilities (+ = cash absorbed)",
}
SEC_ROW = {m: 12 + i for i, m in enumerate(FLOW_ORDER)}          # SEC Data rows 12..18
N_ITER = 60


def q(sheet: str) -> str:
    return f"'{sheet}'!"


def ev_expr(f0: str, g: str, w: str, tg: str, n: str, m: str) -> str:
    """Closed-form EV as an Excel expression (mirrors dcf.enterprise_value)."""
    qq = f"((1+{g})/(1+{w}))"
    return (f"(1+{w})^(0.5*{m})*{f0}*(IF(ABS(1-{qq})<1E-12,{n},{qq}*(1-{qq}^{n})/(1-{qq}))"
            f"+{qq}^{n}*(1+{tg})/({w}-{tg}))")


# fixed addresses shared across sheets ------------------------------------------------
I = {k: f"{q(S_INP)}$C${r}" for k, r in dict(
    date=6, price=7, n=8, mid=9, eqm=10, lb=11, ub=12, rf=15, erp=16, beta=17, kd=18,
    tax=19, tg=20, wstep=23, tstep=24, gstep=25).items()}
SEC = {
    "basis": f"{q(S_SEC)}$C$5", "period_end": f"{q(S_SEC)}$C$6", "filed": f"{q(S_SEC)}$C$7",
    "cash_total": f"{q(S_SEC)}$C$26", "debt_total": f"{q(S_SEC)}$C$32", "nci": f"{q(S_SEC)}$C$33",
    "eqm": f"{q(S_SEC)}$C$34", "equity": f"{q(S_SEC)}$C$35", "shares": f"{q(S_SEC)}$C$41",
    "tax_rate": f"{q(S_SEC)}$F$19",
}
SEC.update({m: f"{q(S_SEC)}$F${r}" for m, r in SEC_ROW.items()})
D = {k: f"{q(S_DCF)}$C${r}" for k, r in dict(
    rf=5, beta=6, erp=7, ke=8, kd=9, tax=10, kd_at=11, shares=12, price=13, mcap=14, debt=15,
    cap=16, we=17, wd=18, wacc=19, ev_target=27, ebit=30, rate=31, taxes=32, nopat=33, dna=34,
    capex=35, dnwc=36, fcff0=37, margin=38, g=41, tg=42, w=43, tfcff=54, tv=55, tvf=56,
    pvtv=57, sumpv=58, ev=59, ev_t=60, resid=61, tvshare=62, ev_ebit=63, ev_fcff=64,
    eq_ev=66, eq_debt=67, eq_nci=68, eq_cash=69, eq_eqm=70, eq=71, eq_sh=72, px=73, px_mkt=74,
    px_diff=75).items()}
SOL = {k: f"{q(S_SOL)}$C${r}" for k, r in dict(
    f0=4, w=5, tg=6, n=7, m=8, target=9, lb=10, ub=11, bracket=12, g=13, resid=14).items()}


class _Writer:
    """Format cache + formula-value cache around an XlsxWriter workbook."""

    def __init__(self, path: Path, cache: dict | None):
        self.wb = xlsxwriter.Workbook(str(path))
        self.wb.set_calc_mode("auto")
        self.cache = cache or {}
        self._formats: dict = {}

    def fmt(self, **props):
        base = dict(font_name="Arial", font_size=10, valign="vcenter")
        base.update({k: v for k, v in props.items() if v is not None})
        key = tuple(sorted(base.items()))
        if key not in self._formats:
            self._formats[key] = self.wb.add_format(base)
        return self._formats[key]

    def put(self, ws, addr: str, value, nf: str | None = None, kind: str = "calc", **style):
        props = dict(style)
        fill = props.pop("fill", False)
        if nf:
            props["num_format"] = nf
        if kind == "input":
            props["font_color"] = BLUE_INPUT
            if fill:
                props.setdefault("bg_color", FILL_INPUT)
        elif kind == "note":
            props.update(font_color="#7F7F7F", italic=True)
        f = self.fmt(**props)
        if isinstance(value, str) and value.startswith("="):
            cached = self.cache.get((ws.get_name(), addr), 0)
            if isinstance(cached, dt.datetime):
                cached = (cached - dt.datetime(1899, 12, 30)).total_seconds() / 86400.0
            ws.write_formula(addr, value, f, cached)
        elif isinstance(value, (dt.date, dt.datetime)):
            ws.write_datetime(addr, dt.datetime.combine(value, dt.time()) if not isinstance(value, dt.datetime) else value, f)
        elif value is None:
            ws.write_blank(addr, None, f)
        else:
            ws.write(addr, value, f)

    # convenience styles --------------------------------------------------------
    def title(self, ws, text, sub):
        self.put(ws, "B1", text, bold=True, font_size=14, font_color=NAVY)
        self.put(ws, "B2", sub, kind="note")

    def section(self, ws, row: int, text: str, c0="B", c1="G"):
        for c in range(ord(c0), ord(c1) + 1):
            self.put(ws, f"{chr(c)}{row}", text if chr(c) == c0 else None, bold=True,
                     font_color="#FFFFFF", bg_color=NAVY)

    def header(self, ws, row: int, labels: dict[str, str], nf=None):
        for col, text in labels.items():
            self.put(ws, f"{col}{row}", text, nf, bold=True, bg_color=LIGHT_BLUE,
                     bottom=1, align="center" if col != "B" else "left", text_wrap=True)


def build_workbook(result: ModelResult, path: Path, cache: dict | None = None) -> Path:
    path = Path(path)
    w = _Writer(path, cache)
    sheets = {n: w.wb.add_worksheet(n) for n in SHEETS}
    for ws in sheets.values():
        ws.hide_gridlines(2)
        ws.set_default_row(15)
        ws.set_landscape()
        ws.fit_to_pages(1, 0)
        ws.set_paper(9)
    sheets[S_SUM].set_tab_color(NAVY)
    sheets[S_INP].set_tab_color("#0000FF")
    sheets[S_SEC].set_tab_color("#0000FF")

    _inputs(w, sheets[S_INP], result)
    _sec(w, sheets[S_SEC], result)
    _historical(w, sheets[S_HIS], result)
    _dcf(w, sheets[S_DCF], result)
    _solver(w, sheets[S_SOL])
    _sensitivity(w, sheets[S_SEN], result)
    _grid(w, sheets[S_GRD])
    _checks(w, sheets[S_CHK], result)
    _methodology(w, sheets[S_MET], result)
    _summary(w, sheets[S_SUM], result)
    sheets[S_SUM].activate()
    w.wb.close()
    return path


# =========================================================================== #
# Inputs
# =========================================================================== #
def _inputs(w: _Writer, ws, r: ModelResult):
    m, s = r.market, r.settings
    ws.set_column("A:A", 2); ws.set_column("B:B", 44); ws.set_column("C:C", 15)
    ws.set_column("D:D", 10); ws.set_column("E:E", 44); ws.set_column("F:F", 12)
    ws.set_column("G:G", 70)
    w.title(ws, "The Coca-Cola Company (NYSE: KO) — Reverse DCF",
            "Inputs & assumptions  |  blue = hard-coded input, black = formula  |  USD millions unless stated")
    w.header(ws, 4, {"B": "Input", "C": "Value", "D": "Units", "E": "Source", "F": "As of", "G": "Notes"})
    prov = m.provenance

    def row(rr, label, value, nf, key, fill=True, units=None, **kw):
        u, src, asof, note = prov.get(key, (units or "", "", "", ""))
        w.put(ws, f"B{rr}", label)
        w.put(ws, f"C{rr}", value, nf, kind="input", fill=fill, align="right", **kw)
        w.put(ws, f"D{rr}", units or u, align="center")
        w.put(ws, f"E{rr}", src); w.put(ws, f"F{rr}", asof, align="center"); w.put(ws, f"G{rr}", note, kind="note")

    w.section(ws, 5, "Valuation settings")
    row(6, "Valuation date", m.valuation_date, NF_DATE, "valuation_date")
    row(7, "Share price (USD)", m.share_price, NF_USD, "share_price")
    row(8, "Explicit forecast years (N)", s.forecast_years, NF_INT, "n", units="years")
    ws.write("E8", "Model structure", w.fmt()); ws.write("G8", "Fixed at 10 by the projection table (see Model Checks).", w.fmt(font_color="#7F7F7F", italic=True))
    row(9, "Mid-year discounting (1 = on, 0 = year-end)", int(s.mid_year), NF_INT, "mid", units="switch")
    ws.write("E9", "Switch", w.fmt()); ws.write("G9", "Damodaran default is year-end. Mid-year also discounts the Gordon TV at N-0.5.", w.fmt(font_color="#7F7F7F", italic=True))
    row(10, "Subtract equity-method investments from EV (1 = yes)", int(s.include_nonoperating), NF_INT, "eqm", units="switch")
    ws.write("E10", "Switch", w.fmt()); ws.write("G10", "EBIT excludes equity income, so the investments (book value) are removed from the operating EV.", w.fmt(font_color="#7F7F7F", italic=True))
    row(11, "Growth solver lower bound", s.growth_lower_bound, NF_PCT1, "lb", units="decimal")
    row(12, "Growth solver upper bound", s.growth_upper_bound, NF_PCT1, "ub", units="decimal")
    for rr in (9, 10):
        ws.data_validation(f"C{rr}", {"validate": "list", "source": [0, 1]})
    w.section(ws, 14, "Cost of capital & terminal growth (market / macro inputs)")
    row(15, "Risk-free rate", m.risk_free_rate, NF_PCT2, "risk_free_rate")
    row(16, "Equity risk premium", m.equity_risk_premium, NF_PCT2, "equity_risk_premium")
    row(17, "Beta", m.beta, "0.00", "beta")
    row(18, "Pre-tax cost of debt", m.pre_tax_cost_of_debt, NF_PCT2, "pre_tax_cost_of_debt")
    row(19, "Marginal tax rate", m.marginal_tax_rate, NF_PCT2, "marginal_tax_rate")
    row(20, "Terminal growth rate", m.terminal_growth_rate, NF_PCT2, "terminal_growth_rate")
    w.section(ws, 22, "Sensitivity-table steps")
    row(23, "WACC step", s.wacc_step, NF_PCT2, "wstep", units="decimal")
    row(24, "Terminal-growth step", s.terminal_growth_step, NF_PCT2, "tstep", units="decimal")
    row(25, "FCFF-growth step (price table)", 0.01, NF_PCT2, "gstep", units="decimal")
    w.put(ws, "B27", "Operating inputs from SEC filings are on the 'SEC Data' sheet; "
          "market inputs above are not available from the SEC and must be refreshed manually.", kind="note")
    ws.freeze_panes(4, 0)


# =========================================================================== #
# SEC data
# =========================================================================== #
def _sec(w: _Writer, ws, r: ModelResult):
    snap = r.snapshot
    ws.set_column("A:A", 2); ws.set_column("B:B", 58); ws.set_column("C:F", 15)
    ws.set_column("G:G", 9); ws.set_column("H:H", 120)
    w.title(ws, "The Coca-Cola Company (NYSE: KO) — SEC Data",
            "Point-in-time extraction from SEC XBRL company facts (data.sec.gov)  |  USD millions")
    w.section(ws, 4, "Reporting basis", "B", "H")
    for rr, label, val, nf in [(5, "Latest filing used", snap.basis_label, None),
                               (6, "Latest period end", snap.period_end, NF_DATE),
                               (7, "Most recent SEC filing date used", snap.filing_date, NF_DATE)]:
        w.put(ws, f"B{rr}", label); w.put(ws, f"C{rr}", val, nf, kind="input", align="left" if nf is None else "right")
    w.put(ws, "B8", "Look-ahead check (filing date <= valuation date)")
    w.put(ws, "C8", f'=IF(C7<={I["date"]},"Pass","FAIL")', bold=True, align="right")
    w.section(ws, 10, "Trailing-twelve-month (TTM) flow items = last fiscal year + current YTD - prior-year YTD", "B", "H")
    w.header(ws, 11, {"B": "Metric", "C": f"FY{snap.fiscal_year_end.year}", "D": "YTD current",
                      "E": "YTD prior year", "F": "TTM", "G": "Units", "H": "Source (SEC XBRL fact; filing date; accession)"})
    for metric in FLOW_ORDER:
        rr = SEC_ROW[metric]; it = snap.ttm[metric]
        w.put(ws, f"B{rr}", FLOW_LABEL[metric])
        w.put(ws, f"C{rr}", it.fiscal_year, NF_MM, kind="input")
        w.put(ws, f"D{rr}", it.ytd_current, NF_MM, kind="input")
        w.put(ws, f"E{rr}", it.ytd_prior, NF_MM, kind="input")
        w.put(ws, f"F{rr}", f"=C{rr}+D{rr}-E{rr}", NF_MM, bold=True)
        w.put(ws, f"G{rr}", "USD m", align="center")
        w.put(ws, f"H{rr}", " ;  ".join(it.sources), kind="note")
    w.put(ws, "B19", "TTM effective tax rate (income taxes / pre-tax income)")
    w.put(ws, "F19", "=F15/F14", NF_PCT2, bold=True)
    w.section(ws, 21, f"Balance sheet at {snap.period_end.isoformat()}", "B", "H")
    w.header(ws, 22, {"B": "Item", "C": "USD m", "H": "Source"})
    b = snap.balance
    items = [(23, "Cash and cash equivalents", b["cash_and_equivalents"], None),
             (24, "Short-term investments", b["short_term_investments"], None),
             (25, "Marketable securities", b["marketable_securities"], None),
             (26, "Cash, short-term investments & marketable securities", None, "=SUM(C23:C25)"),
             (27, "Loans and notes payable", b["loans_and_notes_payable"], None),
             (28, "Current maturities of long-term debt", b["current_maturities_of_debt"], None),
             (29, "Long-term debt", b["long_term_debt"], None),
             (30, "Reported debt", None, "=SUM(C27:C29)"),
             (31, "Operating lease liabilities (latest 10-K; not itemised in 10-Q)", snap.operating_lease_liability, None),
             (32, "Total debt incl. leases", None, "=C30+C31"),
             (33, "Noncontrolling interests (book)", b["noncontrolling_interests"], None),
             (34, "Equity-method investments (book)", b["equity_method_investments"], None),
             (35, "Total equity incl. noncontrolling interests (book)", b["total_equity"], None)]
    for rr, label, item, formula in items:
        total = formula is not None
        w.put(ws, f"B{rr}", label, bold=total)
        if total:
            w.put(ws, f"C{rr}", formula, NF_MM, bold=True, top=1)
        else:
            w.put(ws, f"C{rr}", item.value, NF_MM, kind="input")
            w.put(ws, f"H{rr}", item.source, kind="note")
    w.section(ws, 37, "Share count (millions)", "B", "H")
    w.put(ws, "B38", "Source", kind="note"); w.put(ws, "H38", snap.shares_source, kind="note")
    w.put(ws, "B39", "Basic shares outstanding (10-Q cover page)"); w.put(ws, "C39", snap.shares_outstanding_m, "#,##0.000", kind="input")
    w.put(ws, "B40", "Dilutive securities (diluted - basic weighted average)"); w.put(ws, "C40", snap.dilutive_shares_m, "#,##0.000", kind="input")
    w.put(ws, "B41", "Diluted shares used", bold=True); w.put(ws, "C41", "=C39+C40", "#,##0.000", bold=True, top=1)
    ws.freeze_panes(4, 2)


# =========================================================================== #
# Historical fundamentals
# =========================================================================== #
def _historical(w: _Writer, ws, r: ModelResult):
    ws.set_column("A:A", 2); ws.set_column("B:B", 52); ws.set_column("C:J", 13); ws.set_column("K:K", 3)
    ws.set_column("L:L", 60)
    w.title(ws, "The Coca-Cola Company (NYSE: KO) — Historical Fundamentals",
            "Fiscal years from SEC 10-K filings (blue) and the TTM build  |  USD millions  |  used to test the plausibility of the implied growth")
    years = sorted(r.history["revenue"])
    if years != YEARS:
        raise ValueError(f"Historical sheet expects fiscal years {YEARS}, got {years}.")
    w.put(ws, "B4", "USD millions", bold=True, bg_color=LIGHT_BLUE, bottom=1)
    for i, col in enumerate(HIS_COLS):
        w.put(ws, f"{col}4", YEARS[0] if i == 0 else f"={HIS_COLS[i-1]}4+1", NF_FY, kind="input" if i == 0 else "calc",
              bold=True, bg_color=LIGHT_BLUE, bottom=1, align="right")
    w.put(ws, "J4", "TTM", bold=True, bg_color=LIGHT_BLUE, bottom=1, align="right")
    rows = {"revenue": 5, "ebit": 7, "income_before_income_taxes": 10, "income_tax_expense": 11,
            "depreciation_amortization": 14, "capital_expenditures": 15, "change_in_nwc": 17}
    labels = {5: "Net operating revenues", 6: "  % growth", 7: "Operating income (EBIT)", 8: "  % margin",
              9: "  % growth", 10: "Income before income taxes", 11: "Income taxes", 12: "  Effective tax rate",
              13: "NOPAT = EBIT x (1 - effective tax rate)", 14: "Depreciation & amortization",
              15: "Purchases of PP&E (CapEx)", 16: "Net CapEx (CapEx - D&A)",
              17: "Change in operating assets & liabilities (+ = cash absorbed)",
              18: "FCFF = NOPAT + D&A - CapEx - change in NWC", 19: "  FCFF margin"}
    for rr, text in labels.items():
        w.put(ws, f"B{rr}", text, bold=rr in (5, 7, 13, 18))
    for metric, rr in rows.items():
        for i, col in enumerate(HIS_COLS):
            w.put(ws, f"{col}{rr}", r.history[metric][YEARS[i]], NF_MM, kind="input")
        w.put(ws, f"J{rr}", f"={SEC[metric]}", NF_MM)
    for i, col in enumerate(list(HIS_COLS) + ["J"]):
        prev = HIS_COLS[i - 1] if 0 < i < 7 else None
        if prev:
            w.put(ws, f"{col}6", f"={col}5/{prev}5-1", NF_PCT1, italic=True)
            w.put(ws, f"{col}9", f"={col}7/{prev}7-1", NF_PCT1, italic=True)
        w.put(ws, f"{col}8", f"={col}7/{col}5", NF_PCT1, italic=True)
        w.put(ws, f"{col}12", f"={col}11/{col}10", NF_PCT1, italic=True)
        w.put(ws, f"{col}13", f"={col}7*(1-{col}12)", NF_MM, bold=True)
        w.put(ws, f"{col}16", f"={col}15-{col}14", NF_MM)
        w.put(ws, f"{col}18", f"={col}13+{col}14-{col}15-{col}17", NF_MM, bold=True, top=1)
        w.put(ws, f"{col}19", f"={col}18/{col}5", NF_PCT1, italic=True)
    w.put(ws, "L17", "FY2024-25 include very large outflows tied to income-tax payments; "
          "historical FCFF is therefore not a clean growth benchmark.", kind="note")
    w.put(ws, "L5", "FY2020 is the COVID trough and refranchising reshaped FY2016-19 revenue.", kind="note")

    w.section(ws, 22, "Historical growth benchmarks (compound annual growth to latest fiscal year)", "B", "J")
    w.header(ws, 23, {"B": "Benchmark", "C": "CAGR", "D": "Window"})
    spec = [(24, "Revenue CAGR - 3 years", 5, "F", 3), (25, "Revenue CAGR - 5 years", 5, "D", 5),
            (26, "Revenue CAGR - 6 years", 5, "C", 6), (27, "EBIT CAGR - 3 years", 7, "F", 3),
            (28, "EBIT CAGR - 5 years", 7, "D", 5), (29, "EBIT CAGR - 6 years", 7, "C", 6)]
    for rr, label, base_row, start_col, n in spec:
        w.put(ws, f"B{rr}", label)
        w.put(ws, f"C{rr}", f"=(I{base_row}/{start_col}{base_row})^(1/{n})-1", NF_PCT2, bold=True)
        w.put(ws, f"D{rr}", f'="FY"&{start_col}4&"–FY"&$I$4', align="center", kind="note")

    w.section(ws, 31, "Fundamental growth: g = reinvestment rate x return on invested capital (Damodaran)", "B", "J")
    f = [(32, "TTM NOPAT", "=J13", NF_MM),
         (33, "Invested capital = equity + debt & leases - cash & securities - equity-method inv.",
          f"={SEC['equity']}+{SEC['debt_total']}-{SEC['cash_total']}-{SEC['eqm']}*{I['eqm']}", NF_MM),
         (34, "Return on invested capital (ROIC)", "=C32/C33", NF_PCT2),
         (35, "Reinvestment rate - TTM  (CapEx - D&A + change in NWC) / NOPAT", "=(J15-J14+J17)/J13", NF_PCT2),
         (36, "Fundamental growth - TTM reinvestment", "=C34*C35", NF_PCT2),
         (37, f"Reinvestment rate - normalised, {NORMALISED_WINDOW}-yr (net CapEx / NOPAT, ex-NWC)", "=SUM(E16:I16)/SUM(E13:I13)", NF_PCT2),
         (38, "Fundamental growth - normalised reinvestment", "=C34*C37", NF_PCT2)]
    for rr, label, formula, nf in f:
        w.put(ws, f"B{rr}", label, bold=rr in (36, 38)); w.put(ws, f"C{rr}", formula, nf, bold=rr in (36, 38))

    w.section(ws, 41, "Required FCFF margin in year N implied by the market (vs today's margins)", "B", "J")
    w.header(ws, 42, {"B": "Revenue growth scenario (p.a.)", "C": "Revenue growth", "D": "Year-N revenue",
                      "E": "Year-N FCFF", "F": "Required margin", "G": "vs NOPAT margin"})
    for i, g in enumerate(REVENUE_SCENARIOS):
        rr = 43 + i
        w.put(ws, f"B{rr}", f"Revenue compounds at scenario rate for N years")
        w.put(ws, f"C{rr}", g, NF_PCT1, kind="input", fill=True)
        w.put(ws, f"D{rr}", f"=J5*(1+C{rr})^{I['n']}", NF_MM)
        w.put(ws, f"E{rr}", f"={D['fcff0']}*(1+{D['g']})^{I['n']}", NF_MM)
        w.put(ws, f"F{rr}", f"=E{rr}/D{rr}", NF_PCT1, bold=True)
        w.put(ws, f"G{rr}", f"=F{rr}-$C$47", NF_PCT1)
    w.put(ws, "B46", "Current FCFF margin (TTM)"); w.put(ws, "C46", "=J18/J5", NF_PCT1)
    w.put(ws, "B47", "Current NOPAT margin (TTM) - ceiling for FCFF margin if net reinvestment >= 0"); w.put(ws, "C47", "=J13/J5", NF_PCT1)
    w.put(ws, "B46", "Base-year FCFF margin (model, TTM)"); w.put(ws, "C46", f"={D['margin']}", NF_PCT1)
    ws.freeze_panes(4, 2)


# =========================================================================== #
# DCF engine
# =========================================================================== #
def _dcf(w: _Writer, ws, r: ModelResult):
    ws.set_column("A:A", 2); ws.set_column("B:B", 50); ws.set_column("C:M", 12.5); ws.set_column("N:N", 3)
    ws.set_column("O:O", 70)
    w.title(ws, "The Coca-Cola Company (NYSE: KO) — DCF Engine",
            "WACC, base-year FCFF, market-implied growth, forecast and equity bridge  |  USD millions")

    def line(row, label, formula, nf, bold=False, top=False, note=None, indent=0):
        w.put(ws, f"B{row}", label, bold=bold, indent=indent)
        extra = dict(top=1) if top else {}
        w.put(ws, f"C{row}", formula, nf, bold=bold, **extra)
        if note:
            w.put(ws, f"O{row}", note, kind="note")

    w.section(ws, 4, "1. Cost of capital (CAPM cost of equity, market-value weights)", "B", "M")
    line(5, "Risk-free rate", f"={I['rf']}", NF_PCT2)
    line(6, "Beta", f"={I['beta']}", "0.00")
    line(7, "Equity risk premium", f"={I['erp']}", NF_PCT2)
    line(8, "Cost of equity = rf + beta x ERP", "=C5+C6*C7", NF_PCT2, bold=True)
    line(9, "Pre-tax cost of debt", f"={I['kd']}", NF_PCT2)
    line(10, "Marginal tax rate", f"={I['tax']}", NF_PCT2)
    line(11, "After-tax cost of debt", "=C9*(1-C10)", NF_PCT2, bold=True)
    line(12, "Diluted shares (millions)", f"={SEC['shares']}", "#,##0.0")
    line(13, "Share price (USD)", f"={I['price']}", NF_USD)
    line(14, "Market capitalisation", "=C12*C13", NF_MM)
    line(15, "Total debt incl. operating leases", f"={SEC['debt_total']}", NF_MM)
    line(16, "Total capital", "=C14+C15", NF_MM, top=True)
    line(17, "Weight of equity", "=C14/C16", NF_PCT2)
    line(18, "Weight of debt", "=C15/C16", NF_PCT2)
    line(19, "WACC", "=C17*C8+C18*C11", NF_PCT2, bold=True, top=True)

    w.section(ws, 21, "2. Market enterprise value — the reverse-DCF target", "B", "M")
    line(22, "Market capitalisation", "=C14", NF_MM)
    line(23, "Plus: total debt incl. operating leases", "=C15", NF_MM)
    line(24, "Plus: noncontrolling interests", f"={SEC['nci']}", NF_MM)
    line(25, "Less: cash, short-term investments & marketable securities", f"=-{SEC['cash_total']}", NF_MM)
    line(26, "Less: equity-method investments (x switch)", f"=-{SEC['eqm']}*{I['eqm']}", NF_MM,
         note="EBIT excludes equity income, so these non-operating assets are removed from EV.")
    line(27, "Market enterprise value (target)", "=SUM(C22:C26)", NF_MM, bold=True, top=True)

    w.section(ws, 29, "3. Base-year FCFF (trailing twelve months)", "B", "M")
    line(30, "Operating income (EBIT)", f"={SEC['ebit']}", NF_MM)
    line(31, "Effective tax rate (TTM)", f"={SEC['tax_rate']}", NF_PCT2)
    line(32, "Less: taxes on EBIT", "=-C30*C31", NF_MM)
    line(33, "NOPAT", "=C30+C32", NF_MM, bold=True, top=True)
    line(34, "Plus: depreciation & amortization", f"={SEC['depreciation_amortization']}", NF_MM)
    line(35, "Less: capital expenditures", f"=-{SEC['capital_expenditures']}", NF_MM)
    line(36, "Less: increase in net working capital", f"=-{SEC['change_in_nwc']}", NF_MM,
         note="TTM change in operating assets & liabilities from the cash-flow statement (negative increase = cash released).")
    line(37, "Base-year FCFF (FCFF0)", "=SUM(C33:C36)", NF_MM, bold=True, top=True)
    line(38, "FCFF margin (FCFF0 / TTM revenue)", f"=C37/{SEC['revenue']}", NF_PCT1)

    w.section(ws, 40, "4. Market-implied growth (solved live on the 'Solver' sheet)", "B", "M")
    line(41, "Market-implied FCFF growth, years 1-N", f"={SOL['g']}", NF_PCT2, bold=True,
         note="Solved inside Excel by bisection - no pasted Goal-Seek value.")
    line(42, "Terminal growth rate", f"={I['tg']}", NF_PCT2)
    line(43, "WACC", "=C19", NF_PCT2)

    w.section(ws, 45, "5. Forecast & discounting", "B", "M")
    cols = "CDEFGHIJKLM"
    w.put(ws, "B46", "Year", bold=True, bg_color=LIGHT_BLUE, bottom=1)
    w.put(ws, "B47", "Period end (dynamic: SEC period end + 12 x year)")
    w.put(ws, "B48", "FCFF", bold=True); w.put(ws, "B49", "Discount period (years)")
    w.put(ws, "B50", "Discount factor"); w.put(ws, "B51", "Present value of FCFF", bold=True)
    for i, col in enumerate(cols):
        w.put(ws, f"{col}46", 0 if i == 0 else f"={cols[i-1]}46+1", NF_INT, bold=True, bg_color=LIGHT_BLUE, bottom=1, align="right",
              kind="input" if i == 0 else "calc")
        w.put(ws, f"{col}47", f"={SEC['period_end']}" if i == 0 else f"=EDATE($C$47,12*{col}46)", NF_DATE, align="right")
        if i == 0:
            w.put(ws, "C48", "=C37", NF_MM, bold=True)
        else:
            w.put(ws, f"{col}48", f"={cols[i-1]}48*(1+$C$41)", NF_MM, bold=True)
            w.put(ws, f"{col}49", f"={col}46-0.5*{I['mid']}", "0.0")
            w.put(ws, f"{col}50", f"=1/(1+$C$43)^{col}49", NF_FACTOR)
            w.put(ws, f"{col}51", f"={col}48*{col}50", NF_MM, bold=True)
    w.put(ws, "O46", "Year 0 = TTM base period.", kind="note")

    w.section(ws, 53, "6. Terminal value (Gordon growth) & enterprise value", "B", "M")
    line(54, "Terminal-year FCFF = FCFF(N) x (1 + gT)", "=M48*(1+C42)", NF_MM)
    line(55, "Terminal value at year N = FCFF(N+1) / (WACC - gT)", "=C54/(C43-C42)", NF_MM)
    line(56, "Terminal-value discount factor", "=M50", NF_FACTOR, note="Equals the year-N factor (N-0.5 under mid-year).")
    line(57, "PV of terminal value", "=C55*C56", NF_MM)
    line(58, "Sum of PV of explicit-period FCFF", "=SUM(D51:M51)", NF_MM)
    line(59, "Model enterprise value", "=C57+C58", NF_MM, bold=True, top=True)
    line(60, "Target market enterprise value", "=C27", NF_MM)
    line(61, "Valuation residual (market - model)", "=C60-C59", '#,##0.000000_);(#,##0.000000)',
         note="Should be ~0; solver tolerance < 1e-9.")
    line(62, "Terminal value as % of enterprise value", "=C57/C59", NF_PCT1)
    line(63, "Implied EV / TTM EBIT", "=C59/C30", NF_X)
    line(64, "Implied EV / base-year FCFF", "=C59/C37", NF_X)

    w.section(ws, 65, "7. Equity bridge & implied value per share", "B", "M")
    line(66, "Model enterprise value", "=C59", NF_MM)
    line(67, "Less: total debt incl. operating leases", "=-C15", NF_MM)
    line(68, "Less: noncontrolling interests", f"=-{SEC['nci']}", NF_MM)
    line(69, "Plus: cash, short-term investments & marketable securities", f"={SEC['cash_total']}", NF_MM)
    line(70, "Plus: equity-method investments (x switch)", f"={SEC['eqm']}*{I['eqm']}", NF_MM)
    line(71, "Implied equity value", "=SUM(C66:C70)", NF_MM, bold=True, top=True)
    line(72, "Diluted shares (millions)", "=C12", "#,##0.0")
    line(73, "Implied value per share (USD)", "=C71/C72", NF_USD, bold=True)
    line(74, "Market share price (USD)", "=C13", NF_USD)
    line(75, "Difference (should be ~0 by construction)", "=C73-C74", '$#,##0.000000_);($#,##0.000000)')
    ws.freeze_panes(3, 2)


# =========================================================================== #
# Solver (bisection)
# =========================================================================== #
def _solver(w: _Writer, ws):
    ws.set_column("A:A", 2); ws.set_column("B:B", 40); ws.set_column("C:G", 20)
    w.title(ws, "The Coca-Cola Company (NYSE: KO) — Implied-Growth Solver",
            f"{N_ITER}-step bisection on the closed-form EV: finds g such that EV(g) = market EV (EV is strictly increasing in g)")
    rows = [(4, "Base-year FCFF (FCFF0)", f"={D['fcff0']}", NF_MM), (5, "WACC", f"={D['wacc']}", "0.0000%"),
            (6, "Terminal growth", f"={D['tg']}", "0.0000%"), (7, "Forecast years (N)", f"={I['n']}", NF_INT),
            (8, "Mid-year switch", f"={I['mid']}", NF_INT), (9, "Target market EV", f"={D['ev_target']}", NF_MM),
            (10, "Lower bound for g", f"={I['lb']}", NF_PCT1), (11, "Upper bound for g", f"={I['ub']}", NF_PCT1)]
    for rr, label, formula, nf in rows:
        w.put(ws, f"B{rr}", label); w.put(ws, f"C{rr}", formula, nf)
    ev = lambda g: ev_expr("$C$4", g, "$C$5", "$C$6", "$C$7", "$C$8")
    w.put(ws, "B12", "Bracket valid? (EV(lower) < target < EV(upper))")
    w.put(ws, "C12", f'=IF(AND({ev("$C$10")}<$C$9,{ev("$C$11")}>$C$9),"Pass","FAIL")', bold=True, align="right")
    w.put(ws, "B13", "Market-implied FCFF growth (result)", bold=True)
    first, last = 17, 17 + N_ITER                         # row `last` is the final interval
    w.put(ws, "C13", f"=(C{last}+D{last})/2", "0.000000%", bold=True, bg_color=LIGHT_BLUE)
    w.put(ws, "B14", "EV(result) - target (residual)")
    w.put(ws, "C14", f"={ev('$C$13')}-$C$9", '#,##0.000000_);(#,##0.000000)')
    w.header(ws, 16, {"B": "Iteration", "C": "Low", "D": "High", "E": "Mid", "F": "EV(mid)", "G": "EV(mid) - target"})
    for i in range(N_ITER + 1):
        rr = first + i
        w.put(ws, f"B{rr}", i + 1 if i < N_ITER else "final", align="center")
        if i == 0:
            w.put(ws, f"C{rr}", "=C10", "0.000000000000%"); w.put(ws, f"D{rr}", "=C11", "0.000000000000%")
        else:
            p = rr - 1
            w.put(ws, f"C{rr}", f"=IF(G{p}<0,E{p},C{p})", "0.000000000000%")
            w.put(ws, f"D{rr}", f"=IF(G{p}<0,D{p},E{p})", "0.000000000000%")
        if i < N_ITER:
            w.put(ws, f"E{rr}", f"=(C{rr}+D{rr})/2", "0.000000000000%")
            w.put(ws, f"F{rr}", "=" + ev(f"E{rr}"), NF_MM1)
            w.put(ws, f"G{rr}", f"=F{rr}-$C$9", '#,##0.000000_);(#,##0.000000)')
    ws.freeze_panes(16, 0)


# =========================================================================== #
# Sensitivity tables
# =========================================================================== #
SEN_W0, SEN_T0 = 7, 4        # first data row of table 1 / first data column (D=4th col index 3)


def _sensitivity(w: _Writer, ws, r: ModelResult):
    ws.set_column("A:A", 2); ws.set_column("B:B", 30); ws.set_column("C:H", 14)
    w.title(ws, "The Coca-Cola Company (NYSE: KO) — Sensitivity Analysis",
            "Live formulas (no data tables). Axes are centred on the active WACC / terminal growth so the centre cell ties to the model.")
    tcols = "DEFGH"
    ks = [-2, -1, 0, 1, 2]
    # ---------------------------------------------------------------- table 1
    w.section(ws, 4, "A. Enterprise value (USD m) - WACC vs terminal growth, market-implied growth held fixed", "B", "H")
    w.put(ws, "D5", "Terminal growth rate  →", bold=True)
    w.put(ws, "B6", "WACC  ↓", bold=True, align="right")
    for j, k in enumerate(ks):
        w.put(ws, f"{tcols[j]}6", f"={D['tg']}+({k})*{I['tstep']}", NF_PCT2, bold=True, bg_color=LIGHT_BLUE, bottom=1, align="right")
    for i, k in enumerate(ks):
        rr = 7 + i
        w.put(ws, f"C{rr}", f"={D['wacc']}+({k})*{I['wstep']}", NF_PCT2, bold=True, bg_color=LIGHT_BLUE, align="right")
        for j in range(5):
            w.put(ws, f"{tcols[j]}{rr}", "=" + ev_expr(D["fcff0"], D["g"], f"$C{rr}", f"{tcols[j]}$6", I["n"], I["mid"]),
                  NF_MM, bold=(i == 2 and j == 2), bg_color="#FFF2CC" if (i == 2 and j == 2) else None)
    # ---------------------------------------------------------------- table 2
    w.section(ws, 14, "B. Market-implied FCFF growth (re-solved) - WACC vs terminal growth", "B", "H")
    w.put(ws, "D15", "Terminal growth rate  →", bold=True)
    w.put(ws, "B16", "WACC  ↓", bold=True, align="right")
    for j in range(5):
        w.put(ws, f"{tcols[j]}16", f"={tcols[j]}6", NF_PCT2, bold=True, bg_color=LIGHT_BLUE, bottom=1, align="right")
    for i in range(5):
        rr = 17 + i
        w.put(ws, f"C{rr}", f"=C{7+i}", NF_PCT2, bold=True, bg_color=LIGHT_BLUE, align="right")
        for j in range(5):
            s = i * 5 + j
            col = _col_letter(2 + 3 * s + 2)
            w.put(ws, f"{tcols[j]}{rr}", f"={q(S_GRD)}{col}$6", NF_PCT2, bold=(i == 2 and j == 2),
                  bg_color="#FFF2CC" if (i == 2 and j == 2) else None)
    # ---------------------------------------------------------------- table 3
    w.section(ws, 24, "C. Intrinsic value per share (USD) - alternative FCFF growth vs WACC (what if growth is lower than priced in?)", "B", "H")
    w.put(ws, "D25", "WACC  →", bold=True)
    w.put(ws, "B26", "FCFF growth (yrs 1-N)  ↓", bold=True, align="right")
    for j, k in enumerate(ks):
        w.put(ws, f"{tcols[j]}26", f"={D['wacc']}+({k})*{I['wstep']}", NF_PCT2, bold=True, bg_color=LIGHT_BLUE, bottom=1, align="right")
    bridge = f"({D['eq']}-{D['eq_ev']})"
    for i, k in enumerate(ks):
        rr = 27 + i
        w.put(ws, f"C{rr}", f"={D['g']}+({k})*{I['gstep']}", NF_PCT2, bold=True, bg_color=LIGHT_BLUE, align="right")
        for j in range(5):
            w.put(ws, f"{tcols[j]}{rr}", f"=({ev_expr(D['fcff0'], f'$C{rr}', f'{tcols[j]}$26', D['tg'], I['n'], I['mid'])}+{bridge})/{D['shares']}",
                  NF_USD, bold=(i == 2 and j == 2), bg_color="#FFF2CC" if (i == 2 and j == 2) else None)
    w.put(ws, "B33", "Market share price (USD)"); w.put(ws, "C33", f"={I['price']}", NF_USD)
    green = w.fmt(font_color="#006100", bg_color="#C6EFCE")
    red = w.fmt(font_color="#9C0006", bg_color="#FFC7CE")
    ws.conditional_format("D27:H31", {"type": "cell", "criteria": ">=", "value": f"{I['price']}", "format": green})
    ws.conditional_format("D27:H31", {"type": "cell", "criteria": "<", "value": f"{I['price']}", "format": red})
    w.put(ws, "B34", "Green = value per share at or above the market price; red = below.", kind="note")


def _col_letter(idx: int) -> str:
    return xlsxwriter.utility.xl_col_to_name(idx)


# =========================================================================== #
# Solver grid (25 bisection solves behind sensitivity table B)
# =========================================================================== #
def _grid(w: _Writer, ws):
    ws.set_column("A:A", 2); ws.set_column("B:B", 12); ws.set_column(2, 2 + 75, 13)
    w.title(ws, "The Coca-Cola Company (NYSE: KO) — Solver Grid",
            "25 independent bisection solves (one per WACC x terminal-growth cell) feeding Sensitivity table B; columns = low | high | EV(mid)-target")
    w.put(ws, "B4", "WACC", bold=True); w.put(ws, "B5", "Terminal g", bold=True); w.put(ws, "B6", "Implied g", bold=True)
    w.put(ws, "B8", "Iteration", bold=True, bg_color=LIGHT_BLUE)
    tcols = "DEFGH"
    for it in range(N_ITER + 1):
        w.put(ws, f"B{9+it}", it + 1 if it < N_ITER else "final", align="center")
    for s in range(25):
        i, j = divmod(s, 5)
        lo, hi, fc = (_col_letter(2 + 3 * s + k) for k in range(3))
        w.put(ws, f"{lo}4", f"={q(S_SEN)}$C${7+i}", NF_PCT2, bold=True)
        w.put(ws, f"{lo}5", f"={q(S_SEN)}${tcols[j]}$6", NF_PCT2, bold=True)
        w.put(ws, f"{fc}6", f"=({lo}{9+N_ITER}+{hi}{9+N_ITER})/2", "0.000000%", bold=True, bg_color=LIGHT_BLUE)
        w.put(ws, f"{lo}8", "low", bold=True, bg_color=LIGHT_BLUE, align="center")
        w.put(ws, f"{hi}8", "high", bold=True, bg_color=LIGHT_BLUE, align="center")
        w.put(ws, f"{fc}8", "f(mid)", bold=True, bg_color=LIGHT_BLUE, align="center")
        for it in range(N_ITER + 1):
            rr = 9 + it
            if it == 0:
                w.put(ws, f"{lo}{rr}", f"={SOL['lb']}", "0.0000000%"); w.put(ws, f"{hi}{rr}", f"={SOL['ub']}", "0.0000000%")
            else:
                p = rr - 1
                w.put(ws, f"{lo}{p+1}", f"=IF({fc}{p}<0,({lo}{p}+{hi}{p})/2,{lo}{p})", "0.0000000%")
                w.put(ws, f"{hi}{p+1}", f"=IF({fc}{p}<0,{hi}{p},({lo}{p}+{hi}{p})/2)", "0.0000000%")
            if it < N_ITER:
                mid = f"(({lo}{rr}+{hi}{rr})/2)"
                w.put(ws, f"{fc}{rr}", "=" + ev_expr(SOL["f0"], mid, f"{lo}$4", f"{lo}$5", SOL["n"], SOL["m"]) + f"-{SOL['target']}",
                      '#,##0.000_);(#,##0.000)')
    ws.freeze_panes(8, 2)


# =========================================================================== #
# Checks and Python parity
# =========================================================================== #
def parity_items(r: ModelResult) -> list[tuple[str, float, str, str]]:
    """(label, python value, excel reference, number format) — shared with the tests."""
    f = r.fundamentals
    return [
        ("WACC", r.wacc.wacc, D["wacc"], "0.0000000%"),
        ("Cost of equity", r.wacc.cost_of_equity, D["ke"], "0.0000000%"),
        ("Market capitalisation (USD m)", r.market_cap, D["mcap"], NF_MM1),
        ("Market enterprise value (USD m)", r.market_ev, D["ev_target"], NF_MM1),
        ("Base-year FCFF (USD m)", r.base["fcff"], D["fcff0"], NF_MM1),
        ("TTM effective tax rate", r.base["tax_rate"], SEC["tax_rate"], "0.0000000%"),
        ("Market-implied FCFF growth", r.implied_growth, D["g"], "0.000000000%"),
        ("Terminal value, year N (USD m)", r.schedule.terminal_value, D["tv"], NF_MM1),
        ("PV of terminal value (USD m)", r.schedule.pv_terminal_value, D["pvtv"], NF_MM1),
        ("Sum of PV of FCFF (USD m)", r.schedule.sum_pv_fcff, D["sumpv"], NF_MM1),
        ("Model enterprise value (USD m)", r.schedule.enterprise_value, D["ev"], NF_MM1),
        ("Implied equity value (USD m)", r.equity_value, D["eq"], NF_MM1),
        ("Implied value per share (USD)", r.implied_share_price, D["px"], "$#,##0.0000"),
        ("Revenue CAGR - 5 years", f.revenue_cagr[5], f"{q(S_HIS)}$C$25", "0.0000%"),
        ("EBIT CAGR - 5 years", f.ebit_cagr[5], f"{q(S_HIS)}$C$28", "0.0000%"),
        ("ROIC (TTM)", f.roic, f"{q(S_HIS)}$C$34", "0.0000%"),
        ("Fundamental growth - TTM reinvestment", f.fundamental_growth_ttm, f"{q(S_HIS)}$C$36", "0.0000%"),
        ("Fundamental growth - normalised", f.fundamental_growth_normalised, f"{q(S_HIS)}$C$38", "0.0000%"),
        ("EV grid: low WACC / low terminal g", float(r.ev_grid.iloc[0, 0]), f"{q(S_SEN)}$D$7", NF_MM1),
        ("EV grid: high WACC / high terminal g", float(r.ev_grid.iloc[4, 4]), f"{q(S_SEN)}$H$11", NF_MM1),
        ("Implied-growth grid: low WACC / low g", float(r.growth_grid.iloc[0, 0]), f"{q(S_SEN)}$D$17", "0.000000%"),
        ("Implied-growth grid: high WACC / high g", float(r.growth_grid.iloc[4, 4]), f"{q(S_SEN)}$H$21", "0.000000%"),
        ("Price grid: low growth / low WACC (USD)", float(r.price_grid.iloc[0, 0]), f"{q(S_SEN)}$D$27", "$#,##0.0000"),
        ("Price grid: high growth / high WACC (USD)", float(r.price_grid.iloc[4, 4]), f"{q(S_SEN)}$H$31", "$#,##0.0000"),
    ]


def _checks(w: _Writer, ws, r: ModelResult):
    ws.set_column("A:A", 2); ws.set_column("B:B", 6); ws.set_column("C:C", 62); ws.set_column("D:F", 18)
    ws.set_column("G:G", 12)
    w.title(ws, "The Coca-Cola Company (NYSE: KO) — Model Checks & Python Parity",
            "Integrity checks and a line-by-line reconciliation of this workbook against the Python engine")
    w.put(ws, "C4", "Overall status", bold=True, font_size=11)
    w.section(ws, 6, "A. Model integrity checks", "B", "G")
    w.header(ws, 7, {"B": "#", "C": "Check", "D": "Value", "E": "Tolerance", "F": "", "G": "Status"})
    chk = [
        ("Valuation residual: |model EV - market EV| (USD m)", f"=ABS({D['resid']})", 1e-6, "<="),
        ("Gordon growth constraint: WACC - terminal growth > 0", f"={D['wacc']}-{D['tg']}", 0.0, ">"),
        ("Terminal growth below risk-free rate (Damodaran cap)", f"={I['rf']}-{D['tg']}", 0.0, ">"),
        ("Solver bracket valid (EV(lower) < target < EV(upper))", f'=IF({SOL["bracket"]}="Pass",0,1)', 0.0, "<="),
        ("Implied growth strictly inside solver bounds", f"=MIN({D['g']}-{I['lb']},{I['ub']}-{D['g']})", 0.0, ">"),
        ("Solver residual: |EV(g*) - target| (USD m)", f"=ABS({SOL['resid']})", 1e-6, "<="),
        ("Closed-form EV vs explicit schedule: |difference| (USD m)", f"=ABS({SOL['target']}-{D['ev']})", 1e-6, "<="),
        ("Implied value per share ties to market price: |diff| (USD)", f"=ABS({D['px_diff']})", 1e-6, "<="),
        ("Implied equity value ties to market cap: |diff| (USD m)", f"=ABS({D['eq']}-{D['mcap']})", 1e-6, "<="),
        ("Forecast years equal projection columns (N = 10)", f"=ABS({I['n']}-{q(S_DCF)}$M$46)", 0.0, "<="),
        ("Look-ahead: SEC filing date <= valuation date (days of slack)", f"={I['date']}-{SEC['filed']}", 0.0, ">="),
        ("Sensitivity A centre ties to model EV: |diff| (USD m)", f"=ABS({q(S_SEN)}$F$9-{D['ev']})", 1e-6, "<="),
        ("Sensitivity B centre ties to implied growth: |diff|", f"=ABS({q(S_SEN)}$F$19-{D['g']})", 1e-12, "<="),
        ("Sensitivity C centre ties to market price: |diff| (USD)", f"=ABS({q(S_SEN)}$F$29-{I['price']})", 1e-6, "<="),
        ("All inputs populated (blank count on Inputs)", f"=COUNTBLANK({q(S_INP)}$C$6:$C$12)+COUNTBLANK({q(S_INP)}$C$15:$C$20)", 0.0, "<="),
        ("SEC cash-flow items positive where required (CapEx > 0, EBIT > 0, pre-tax > 0)", f"=IF(AND({SEC['capital_expenditures']}>0,{SEC['ebit']}>0,{SEC['income_before_income_taxes']}>0),0,1)", 0.0, "<="),
    ]
    first = 8
    for i, (label, formula, tol, op) in enumerate(chk):
        rr = first + i
        w.put(ws, f"B{rr}", i + 1, align="center"); w.put(ws, f"C{rr}", label)
        w.put(ws, f"D{rr}", formula, "0.000000E+00;-0.000000E+00;0")
        w.put(ws, f"E{rr}", f"{op} {tol:g}", align="center", kind="note")
        w.put(ws, f"G{rr}", f'=IF(D{rr}{op}{tol!r},"Pass","FAIL")', bold=True, align="center")
    last_a = first + len(chk) - 1

    start = last_a + 3
    w.section(ws, start, "B. Python parity: Excel (live formulas) vs Python engine (blue = pasted by run_model.py)", "B", "G")
    w.header(ws, start + 1, {"B": "#", "C": "Quantity", "D": "Python", "E": "Excel (live)", "F": "Abs. difference", "G": "Status"})
    items = parity_items(r)
    for i, (label, py, ref, nf) in enumerate(items):
        rr = start + 2 + i
        w.put(ws, f"B{rr}", i + 1, align="center"); w.put(ws, f"C{rr}", label)
        w.put(ws, f"D{rr}", py, nf, kind="input")
        w.put(ws, f"E{rr}", f"={ref}", nf)
        w.put(ws, f"F{rr}", f"=ABS(D{rr}-E{rr})", "0.00E+00")
        w.put(ws, f"G{rr}", f'=IF(F{rr}<=1E-9*MAX(1,ABS(D{rr})),"Pass","FAIL")', bold=True, align="center")
    last_b = start + 1 + len(items)
    rng = f"G{first}:G{last_a}"
    rng2 = f"G{start+2}:G{last_b}"
    w.put(ws, "D4", f'=IF(COUNTIF({rng},"Pass")+COUNTIF({rng2},"Pass")=ROWS({rng})+ROWS({rng2}),"ALL CHECKS PASS","CHECK ERRORS")',
          bold=True, font_size=11)
    green = w.fmt(font_color="#006100", bg_color="#C6EFCE", bold=True)
    red = w.fmt(font_color="#9C0006", bg_color="#FFC7CE", bold=True)
    for target in (rng, rng2, "D4"):
        ws.conditional_format(target, {"type": "text", "criteria": "begins with", "value": "Pass", "format": green})
        ws.conditional_format(target, {"type": "text", "criteria": "begins with", "value": "ALL", "format": green})
        ws.conditional_format(target, {"type": "text", "criteria": "begins with", "value": "FAIL", "format": red})
        ws.conditional_format(target, {"type": "text", "criteria": "begins with", "value": "CHECK", "format": red})
    ws.freeze_panes(5, 0)


# =========================================================================== #
# Methodology
# =========================================================================== #
def _methodology(w: _Writer, ws, r: ModelResult):
    ws.set_column("A:A", 2); ws.set_column("B:B", 150)
    w.title(ws, "The Coca-Cola Company (NYSE: KO) — Methodology & Limitations", "Aswath Damodaran FCFF framework, run in reverse")
    lines = [
        ("Reverse DCF", True),
        ("Standard DCF: forecast growth -> value.  Reverse DCF: take the market EV as given and solve for the single growth rate g that equates model EV to it:", False),
        ("      EV_market = SUM[t=1..N] FCFF0 (1+g)^t / (1+WACC)^t  +  FCFF0 (1+g)^N (1+gT) / (WACC - gT) / (1+WACC)^N", False),
        ("The solution is found live on the 'Solver' sheet by bisection on the equivalent closed form (geometric series in q = (1+g)/(1+WACC)).", False),
        ("", False),
        ("Definitions", True),
        ("FCFF = EBIT(1 - effective tax rate) + D&A - CapEx - change in NWC, all TTM. Cost of equity = rf + beta x ERP. WACC uses market-value weights; debt includes operating leases.", False),
        ("Market EV = market cap + debt & leases + noncontrolling interests - cash & marketable securities - equity-method investments (EBIT excludes equity income).", False),
        ("Terminal value: Gordon growth on FCFF(N+1); gT is capped below the risk-free rate (Damodaran). Terminal value is discounted with the year-N factor.", False),
        ("", False),
        ("Plausibility (Damodaran)", True),
        ("Compare implied growth with: observed revenue/EBIT CAGRs; fundamental growth = reinvestment rate x ROIC; and the FCFF margin that year N would require.", False),
        ("", False),
        ("Limitations", True),
        ("1) One constant growth rate over N years then perpetuity - no fade, no margin path.  2) Terminal reinvestment is not re-based to g/ROIC.", False),
        ("3) Regression beta (0.58) is low; Damodaran favours bottom-up betas.  4) Book values proxy for equity-method investments and minority interests.", False),
        ("5) Operating leases are from the latest 10-K (not itemised quarterly).  6) TTM change in NWC includes tax-related items and is volatile.", False),
        ("7) Market inputs (price, rf, ERP, beta, cost of debt) are not SEC data and must be refreshed manually.  8) Not investment advice.", False),
    ]
    for i, (text, bold) in enumerate(lines):
        w.put(ws, f"B{4+i}", text, bold=bold, text_wrap=True, font_color=NAVY if bold else "#000000")


# =========================================================================== #
# Summary dashboard
# =========================================================================== #
def _summary(w: _Writer, ws, r: ModelResult):
    ws.set_column("A:A", 2); ws.set_column("B:B", 46); ws.set_column("C:H", 18); ws.set_column("I:I", 3)
    ws.set_column("J:P", 13)
    w.put(ws, "B1", "The Coca-Cola Company (NYSE: KO) — Reverse DCF", bold=True, font_size=16, font_color=NAVY)
    w.put(ws, "B2", f'="Valuation date: "&TEXT({I["date"]},"d mmm yyyy")&"   |   SEC basis: "&{SEC["basis"]}&"   |   USD millions unless stated"', kind="note")
    tiles = [("B", "Share price", f"={I['price']}", NF_USD), ("C", "Market EV", f"={D['ev_target']}", NF_MM),
             ("D", "WACC", f"={D['wacc']}", NF_PCT2), ("E", "Terminal growth", f"={D['tg']}", NF_PCT2),
             ("F", "Implied FCFF growth", f"={D['g']}", NF_PCT2), ("G", "Model status", f"={q(S_CHK)}$D$4", None)]
    for col, label, formula, nf in tiles:
        w.put(ws, f"{col}4", label, bold=True, font_color="#FFFFFF", bg_color=MID_BLUE, align="center")
        w.put(ws, f"{col}5", formula, nf, bold=True, font_size=14 if col != "G" else 9, align="center", bg_color=GREY, font_color=NAVY)
    ws.set_row(4, 26)
    green = w.fmt(font_color="#006100", bg_color="#C6EFCE", bold=True, align="center", font_size=9)
    red = w.fmt(font_color="#9C0006", bg_color="#FFC7CE", bold=True, align="center", font_size=9)
    ws.conditional_format("G5", {"type": "text", "criteria": "begins with", "value": "ALL", "format": green})
    ws.conditional_format("G5", {"type": "text", "criteria": "begins with", "value": "CHECK", "format": red})

    w.section(ws, 7, "Reverse-DCF read-out", "B", "H")
    w.put(ws, "B8", f'="At "&TEXT({I["price"]},"$0.00")&" per share the market is discounting FCFF growth of "&TEXT({D["g"]},"0.00%")&" a year for "&{I["n"]}&" years, then "&TEXT({D["tg"]},"0.00%")&" in perpetuity."', bold=True)
    w.put(ws, "B9", f'="Base-year FCFF (TTM): "&TEXT({D["fcff0"]},"#,##0")&"  ->  year-"&{I["n"]}&" FCFF of "&TEXT({D["fcff0"]}*(1+{D["g"]})^{I["n"]},"#,##0")&"  ("&TEXT((1+{D["g"]})^{I["n"]},"0.00")&"x).  Terminal value = "&TEXT({D["tvshare"]},"0%")&" of EV."')

    w.section(ws, 11, "Plausibility: market-implied growth vs historical fundamentals", "B", "H")
    w.header(ws, 12, {"B": "Benchmark", "C": "Growth", "D": "Implied minus benchmark", "E": "Implied above?"})
    his = q(S_HIS)
    bench = [("Market-implied FCFF growth (years 1-N)", f"={D['g']}"),
             ("Revenue CAGR - 3 years", f"={his}$C$24"), ("Revenue CAGR - 5 years", f"={his}$C$25"),
             ("Revenue CAGR - 6 years", f"={his}$C$26"), ("EBIT CAGR - 3 years", f"={his}$C$27"),
             ("EBIT CAGR - 5 years", f"={his}$C$28"), ("EBIT CAGR - 6 years", f"={his}$C$29"),
             ("Fundamental growth (TTM reinvestment x ROIC)", f"={his}$C$36"),
             ("Fundamental growth (normalised reinvestment x ROIC)", f"={his}$C$38"),
             ("Terminal growth (perpetuity)", f"={D['tg']}")]
    for i, (label, formula) in enumerate(bench):
        rr = 13 + i
        w.put(ws, f"B{rr}", label, bold=(i == 0)); w.put(ws, f"C{rr}", formula, NF_PCT2, bold=(i == 0))
        if i > 0:
            w.put(ws, f"D{rr}", f"=$C$13-C{rr}", '+0.00%;-0.00%'); w.put(ws, f"E{rr}", f'=IF($C$13>C{rr},"Yes","No")', align="center")
    w.put(ws, "B24", f'="Implied growth exceeds "&COUNTIF(E14:E21,"Yes")&" of "&COUNTA(E14:E21)&" historical / fundamental benchmarks."', bold=True)
    w.put(ws, "B25", f'="Required year-"&{I["n"]}&" FCFF margin if revenue grows 5% a year: "&TEXT({his}$F$44,"0.0%")&" vs. NOPAT margin today of "&TEXT({his}$C$47,"0.0%")&" (FCFF margin today "&TEXT({his}$C$46,"0.0%")&")."')
    amber = w.fmt(font_color="#9C5700", bg_color="#FFEB9C")
    ws.conditional_format("E14:E21", {"type": "text", "criteria": "begins with", "value": "Yes", "format": amber})

    # Chart 1 -------------------------------------------------------------
    c1 = w.wb.add_chart({"type": "column"})
    c1.add_series({"name": "Growth", "categories": f"={q(S_SUM)}$B$13:$B$22", "values": f"={q(S_SUM)}$C$13:$C$22",
                   "points": [{"fill": {"color": "#C00000"}}] + [{"fill": {"color": MID_BLUE}}] * 9,
                   "data_labels": {"value": True, "num_format": "0.0%", "font": {"name": "Arial", "size": 8}}, "gap": 60})
    c1.set_title({"name": "Market-implied growth vs benchmarks", "name_font": {"name": "Arial", "size": 11, "bold": True}})
    c1.set_legend({"none": True})
    c1.set_x_axis({"num_font": {"name": "Arial", "size": 7, "rotation": -35}})
    c1.set_y_axis({"num_format": "0%", "num_font": {"name": "Arial", "size": 8}, "major_gridlines": {"visible": True, "line": {"color": "#D9D9D9"}}})
    c1.set_size({"width": 560, "height": 320})
    ws.insert_chart("J4", c1)

    # Sensitivity (implied growth) -----------------------------------------
    w.section(ws, 27, "Sensitivity: market-implied FCFF growth - WACC (rows) vs terminal growth (columns)", "B", "H")
    sen = q(S_SEN)
    w.put(ws, "B28", "WACC  ↓   /   Terminal growth  →", bold=True, align="right")
    for j, col in enumerate("DEFGH"):
        w.put(ws, f"{col}28", f"={sen}{col}16", NF_PCT2, bold=True, bg_color=LIGHT_BLUE, bottom=1, align="right")
    for i in range(5):
        rr = 29 + i
        w.put(ws, f"C{rr}", f"={sen}C{17+i}", NF_PCT2, bold=True, bg_color=LIGHT_BLUE, align="right")
        for col in "DEFGH":
            w.put(ws, f"{col}{rr}", f"={sen}{col}{17+i}", NF_PCT2, bold=(i == 2 and col == "F"),
                  bg_color="#FFF2CC" if (i == 2 and col == "F") else None)
    w.put(ws, "B34", "Value per share if growth is lower than priced in (centre = market price) - see 'Sensitivity' table C.", kind="note")

    # Bridge ---------------------------------------------------------------
    w.section(ws, 36, "Enterprise value bridge: from discounted cash flows to equity value (USD m)", "B", "H")
    hdr = {"B": "Step", "C": "Base (hidden)", "D": "Increase", "E": "Decrease", "F": "Total", "G": "Running"}
    w.header(ws, 37, hdr)
    steps = [
        ("PV of FCFF (years 1-N)", "0", f"={D['sumpv']}", "0", "0", f"=D38"),
        ("PV of terminal value", "=G38", f"={D['pvtv']}", "0", "0", "=G38+D39"),
        ("Enterprise value", "0", "0", "0", f"={D['ev']}", f"={D['ev']}"),
        ("Less: debt & leases", "=G40-E41", "0", f"=-{D['eq_debt']}", "0", "=G40-E41"),
        ("Less: noncontrolling interests", "=G41-E42", "0", f"=-{D['eq_nci']}", "0", "=G41-E42"),
        ("Plus: cash & securities", "=G42", f"={D['eq_cash']}", "0", "0", "=G42+D43"),
        ("Plus: equity-method investments", "=G43", f"={D['eq_eqm']}", "0", "0", "=G43+D44"),
        ("Equity value", "0", "0", "0", f"={D['eq']}", f"={D['eq']}"),
    ]
    for i, (label, base, inc, dec, tot, run) in enumerate(steps):
        rr = 38 + i
        w.put(ws, f"B{rr}", label, bold=label in ("Enterprise value", "Equity value"))
        for col, val in zip("CDEFG", (base, inc, dec, tot, run)):
            w.put(ws, f"{col}{rr}", val if str(val).startswith("=") else float(val), NF_MM)
    c2 = w.wb.add_chart({"type": "column", "subtype": "stacked"})
    cat = f"={q(S_SUM)}$B$38:$B$45"
    c2.add_series({"name": "Base", "categories": cat, "values": f"={q(S_SUM)}$C$38:$C$45", "fill": {"none": True}, "border": {"none": True}})
    c2.add_series({"name": "Increase", "categories": cat, "values": f"={q(S_SUM)}$D$38:$D$45", "fill": {"color": "#548235"}, "gap": 40})
    c2.add_series({"name": "Decrease", "categories": cat, "values": f"={q(S_SUM)}$E$38:$E$45", "fill": {"color": "#C00000"}})
    c2.add_series({"name": "Total", "categories": cat, "values": f"={q(S_SUM)}$F$38:$F$45", "fill": {"color": NAVY},
                   "data_labels": {"value": True, "num_format": "#,##0", "font": {"name": "Arial", "size": 8}}})
    c2.set_title({"name": "Enterprise value bridge (USD m)", "name_font": {"name": "Arial", "size": 11, "bold": True}})
    c2.set_legend({"none": True})
    c2.set_x_axis({"num_font": {"name": "Arial", "size": 7, "rotation": -30}})
    c2.set_y_axis({"num_format": "#,##0", "num_font": {"name": "Arial", "size": 8}, "major_gridlines": {"visible": True, "line": {"color": "#D9D9D9"}}})
    c2.set_size({"width": 560, "height": 320})
    ws.insert_chart("J22", c2)

    # FCFF path -------------------------------------------------------------
    c3 = w.wb.add_chart({"type": "column"})
    c3.add_series({"name": "FCFF", "categories": f"={q(S_DCF)}$C$46:$M$46", "values": f"={q(S_DCF)}$C$48:$M$48", "fill": {"color": MID_BLUE}, "gap": 50})
    c3.add_series({"name": "PV of FCFF", "categories": f"={q(S_DCF)}$C$46:$M$46", "values": f"={q(S_DCF)}$C$51:$M$51", "fill": {"color": "#A6A6A6"}})
    c3.set_title({"name": "Market-implied FCFF path (USD m)", "name_font": {"name": "Arial", "size": 11, "bold": True}})
    c3.set_legend({"position": "bottom", "font": {"name": "Arial", "size": 8}})
    c3.set_x_axis({"name": "Year", "num_font": {"name": "Arial", "size": 8}})
    c3.set_y_axis({"num_format": "#,##0", "num_font": {"name": "Arial", "size": 8}, "major_gridlines": {"visible": True, "line": {"color": "#D9D9D9"}}})
    c3.set_size({"width": 560, "height": 300})
    ws.insert_chart("J39", c3)
    w.put(ws, "B47", "Model integrity: see 'Model Checks'. Blue = hard-coded input, black = formula. Methodology and limitations: see 'Methodology'.", kind="note")
    ws.set_first_sheet()


# =========================================================================== #
# LibreOffice recalculation (to embed cached values) and the public entry point
# =========================================================================== #
def libreoffice_available() -> str | None:
    return shutil.which("soffice") or shutil.which("libreoffice")


def recalculate_values(xlsx_path: Path) -> dict:
    """Open the workbook in headless LibreOffice, recalculate, return {(sheet, addr): value}."""
    import openpyxl
    exe = libreoffice_available()
    if exe is None:
        raise RuntimeError("LibreOffice (soffice) not found.")
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "out"
        out.mkdir()
        profile = Path(tmp) / "profile"
        # Force "always recalculate" for OOXML files (XlsxWriter stores 0 as the cached value).
        user = profile / "user"
        user.mkdir(parents=True)
        (user / "registrymodifications.xcu").write_text(
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<oor:items xmlns:oor="http://openoffice.org/2001/registry" '
            'xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            '<item oor:path="/org.openoffice.Office.Calc/Formula/Load">'
            '<prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item></oor:items>',
            encoding="utf-8")
        subprocess.run([exe, f"-env:UserInstallation=file://{profile}", "--headless", "--calc",
                        "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(out), str(xlsx_path)],
                       check=True, capture_output=True, timeout=300)
        wb = openpyxl.load_workbook(out / Path(xlsx_path).name, data_only=True)
        values = {}
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    if c.value is not None:
                        values[(ws.title, c.coordinate)] = c.value
    return values


def write_model_workbook(result: ModelResult, path: Path) -> tuple[Path, dict | None]:
    """Build the workbook; when LibreOffice is available, embed recalculated cached values."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    if libreoffice_available() is None:
        build_workbook(result, path)
        return path, None
    with tempfile.TemporaryDirectory() as tmp:
        draft = Path(tmp) / path.name
        build_workbook(result, draft)
        cache = recalculate_values(draft)
    build_workbook(result, path, cache)
    return path, cache
