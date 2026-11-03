"""Excel <-> Python parity: recalculate the workbook in LibreOffice and compare."""
import pytest

from src.reverse_dcf import excel_model as xm

pytestmark = pytest.mark.skipif(xm.libreoffice_available() is None, reason="LibreOffice not installed")


@pytest.fixture(scope="module")
def excel_values(result, tmp_path_factory):
    path = tmp_path_factory.mktemp("xl") / "model.xlsx"
    xm.build_workbook(result, path)
    return xm.recalculate_values(path)


def _cell(values, ref):
    sheet, addr = ref.split("!")
    return values[(sheet.strip("'"), addr.replace("$", ""))]


def test_every_parity_item_matches(result, excel_values):
    for label, py, ref, _ in xm.parity_items(result):
        xl = _cell(excel_values, ref)
        assert xl == pytest.approx(py, rel=1e-12, abs=1e-12), label


def test_workbook_self_checks_pass(excel_values):
    assert excel_values[("Model Checks", "D4")] == "ALL CHECKS PASS"


def test_excel_solver_equals_python_bisection(result, excel_values):
    assert excel_values[("Solver", "C13")] == pytest.approx(result.implied_growth_bisection, abs=1e-14)
    assert excel_values[("Solver", "C13")] == pytest.approx(result.implied_growth, abs=1e-12)


def test_full_growth_grid(result, excel_values):
    for i in range(5):
        for j in range(5):
            xl = excel_values[("Sensitivity", f"{'DEFGH'[j]}{17 + i}")]
            assert xl == pytest.approx(float(result.growth_grid.iloc[i, j]), abs=1e-12)


@pytest.mark.parametrize("kwargs", [dict(mid_year=True), dict(include_nonoperating=False)])
def test_switches_stay_in_parity(facts, kwargs, tmp_path):
    from src.reverse_dcf.config import ModelSettings, load_market_inputs
    from src.reverse_dcf.model import run_model
    res = run_model(load_market_inputs(), ModelSettings(**kwargs), facts)
    path = tmp_path / "m.xlsx"
    xm.build_workbook(res, path)
    vals = xm.recalculate_values(path)
    assert vals[("Model Checks", "D4")] == "ALL CHECKS PASS"
    assert vals[("Solver", "C13")] == pytest.approx(res.implied_growth, abs=1e-12)
