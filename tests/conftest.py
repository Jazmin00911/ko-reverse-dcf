from pathlib import Path
import pytest

from src.reverse_dcf.config import ModelSettings, load_market_inputs
from src.reverse_dcf.model import run_model
from src.reverse_dcf.sec import CompanyFacts, RAW_FACTS_PATH

needs_facts = pytest.mark.skipif(not RAW_FACTS_PATH.exists(), reason="SEC company facts not downloaded")


@pytest.fixture(scope="session")
def facts():
    return CompanyFacts.from_path()


@pytest.fixture(scope="session")
def result(facts):
    return run_model(load_market_inputs(), ModelSettings(), facts)
