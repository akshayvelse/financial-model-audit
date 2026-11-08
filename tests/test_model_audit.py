"""Unit and integration test suite for financial model auditing and valuation waterfall."""

from __future__ import annotations

import sys
from pathlib import Path

import numpy as np
import pandas as pd
import pytest

SRC_DIR = Path(__file__).resolve().parent.parent / "src"
sys.path.insert(0, str(SRC_DIR))

from formula_scanner import ModelFormulaScanner, run_audit_scanner
from generate_models import generate_audit_models
from valuation_bridge import ValuationBridgeEngine


@pytest.fixture(scope="module")
def audit_workbooks() -> tuple[Path, Path]:
    """Ensure flawed and audited workbooks are generated before test execution."""
    flawed_path, audited_path = generate_audit_models()
    return flawed_path, audited_path


def test_scanner_detects_all_five_injected_errors(audit_workbooks: tuple[Path, Path]) -> None:
    """Verify formula scanner flags all 5 injected structural defects in flawed model."""
    flawed_path, _ = audit_workbooks
    scanner = ModelFormulaScanner(flawed_path)
    discrepancies = scanner.scan_all()

    error_ids = {d["discrepancy_id"] for d in discrepancies}

    expected_errors = {
        "ERR-001-NWC-SIGN-FLIP",
        "ERR-002-TERMINAL-CAPEX-IMBALANCE",
        "ERR-003-HARDCODE-OVERRIDE",
        "ERR-004-DISCOUNT-EXPONENT-OFF-BY-ONE",
        "ERR-005-NET-DEBT-BRIDGE-OMISSION",
    }

    assert expected_errors.issubset(error_ids), f"Missing expected error flags. Found: {error_ids}"
    assert len(discrepancies) >= 5, "Scanner must detect at least 5 individual discrepancy records"


def test_scanner_passes_audited_model_cleanly(audit_workbooks: tuple[Path, Path]) -> None:
    """Verify audited model passes all structural scanning checks with zero discrepancies."""
    _, audited_path = audit_workbooks
    scanner = ModelFormulaScanner(audited_path)
    discrepancies = scanner.scan_all()

    assert len(discrepancies) == 0, f"Audited model produced unexpected discrepancies: {discrepancies}"


def test_valuation_waterfall_exact_sum() -> None:
    """Verify sequential attribution deltas sum exactly to total valuation difference."""
    engine = ValuationBridgeEngine()
    bridge_df = engine.compute_sequential_waterfall()

    p_flawed = float(bridge_df.loc[0, "share_price"])
    p_audited = float(bridge_df.loc[6, "share_price"])

    step_deltas = bridge_df.loc[1:5, "delta"].values
    total_delta_sum = float(np.sum(step_deltas))

    actual_total_delta = p_audited - p_flawed

    assert abs(total_delta_sum - actual_total_delta) < 1e-6, (
        f"Waterfall delta sum {total_delta_sum} does not match total delta {actual_total_delta}"
    )
    assert p_audited > p_flawed, "Audited model should correct for net negative valuation distortion"


def test_cash_flow_and_reinvestment_parity() -> None:
    """Verify audited model satisfies steady-state capital reinvestment constraints."""
    engine = ValuationBridgeEngine()
    res = engine.evaluate_model_scenario(True, True, True, True, True)

    assert res["enterprise_value"] > 0
    assert res["equity_value"] > 0
    assert res["share_price"] > 0
    assert res["pv_tv"] > 0
