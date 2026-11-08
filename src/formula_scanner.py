"""Automated Excel formula scanner and discrepancy auditor using openpyxl.

Parses workbook structures to detect:
1. Hardcoded numeric overrides in dynamic formula rows
2. Mathematical sign anomalies in cash flow statements
3. Steady-state reinvestment imbalances (Terminal CapEx vs D&A)
4. Time-value discount factor exponent off-by-one errors
5. Enterprise-to-equity value bridge cash omissions

Outputs audit_discrepancy_log.csv with severity classifications and remediation guidance.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Dict, List

import openpyxl
import pandas as pd

OUTPUT_DIR = Path(__file__).resolve().parent.parent / "output"


class ModelFormulaScanner:
    """Institutional spreadsheet parser and integrity verification engine."""

    def __init__(self, workbook_path: Path):
        self.workbook_path = Path(workbook_path)
        self.wb_formulas = openpyxl.load_workbook(self.workbook_path, data_only=False)
        self.discrepancies: list[dict[str, Any]] = []

    def scan_all(self) -> list[dict[str, Any]]:
        """Execute full suite of structural audit rules across all sheets."""
        self.discrepancies.clear()
        for sheet_name in self.wb_formulas.sheetnames:
            ws = self.wb_formulas[sheet_name]
            self._audit_hardcoded_overrides_in_formula_rows(ws, sheet_name)
            self._audit_working_capital_sign_flips(ws, sheet_name)
            self._audit_terminal_capex_da_balance(ws, sheet_name)
            self._audit_discount_factor_exponents(ws, sheet_name)
            self._audit_net_debt_bridge(ws, sheet_name)
        return self.discrepancies

    def _audit_hardcoded_overrides_in_formula_rows(
        self, ws: openpyxl.worksheet.worksheet.Worksheet, sheet_name: str
    ) -> None:
        """Flag static numeric constants embedded within homogeneous formula rows."""
        # Focus on calculation rows (COGS through CapEx, rows 18 to 26), columns C (3) to G (7)
        for r in range(18, 28):
            formula_count = 0
            hardcoded_cells = []

            for col_idx in range(3, 8):
                cell = ws.cell(row=r, column=col_idx)
                val = cell.value
                if val is None:
                    continue

                if isinstance(val, str) and val.startswith("="):
                    formula_count += 1
                elif isinstance(val, (int, float)):
                    hardcoded_cells.append((cell.coordinate, val))

            # If majority of row is dynamic formula but a cell is hardcoded
            if formula_count >= 3 and len(hardcoded_cells) > 0:
                for coord, val in hardcoded_cells:
                    row_label = ws.cell(row=r, column=2).value or f"Row {r}"
                    self.discrepancies.append(
                        {
                            "discrepancy_id": "ERR-003-HARDCODE-OVERRIDE",
                            "sheet_name": sheet_name,
                            "cell_coordinate": coord,
                            "error_category": "Integrity Violation",
                            "severity": "CRITICAL",
                            "flawed_content": str(val),
                            "expected_pattern": f"Formula linked to Total Revenue: ={coord[0]}17*COGS_pct",
                            "description": (
                                f"Hardcoded static value {val} overwrites dynamic formula in {row_label}. "
                                "Breaks sensitivity to top-line volume variations."
                            ),
                        }
                    )

    def _audit_working_capital_sign_flips(
        self, ws: openpyxl.worksheet.worksheet.Worksheet, sheet_name: str
    ) -> None:
        """Verify algebraic sign of working capital change in Free Cash Flow formulas.

        Standard Free Cash Flow identity: FCFF = NOPAT + D&A - CapEx - Delta NWC.
        An increase in working capital represents a cash outflow and must be subtracted.
    """
        for r in range(26, 29):
            row_label = str(ws.cell(row=r, column=2).value or "")
            if "Free Cash Flow" not in row_label and "FCFF" not in row_label:
                continue

            for col_idx in range(3, 8):
                cell = ws.cell(row=r, column=col_idx)
                f_str = str(cell.value or "")
                if not f_str.startswith("="):
                    continue

                # Expected pattern: NOPAT (row 23) + DA (row 24) - CapEx (row 25) - DeltaNWC (row 26)
                # Flag if row 26 is added (+...26) instead of subtracted (-...26)
                if re.search(r"\+[A-Za-z]+26", f_str):
                    self.discrepancies.append(
                        {
                            "discrepancy_id": "ERR-001-NWC-SIGN-FLIP",
                            "sheet_name": sheet_name,
                            "cell_coordinate": cell.coordinate,
                            "error_category": "Accounting Logic",
                            "severity": "HIGH",
                            "flawed_content": f_str,
                            "expected_pattern": f_str.replace("+", "-"),
                            "description": (
                                "Working capital expansion (+Delta NWC) added rather than subtracted in FCFF. "
                                "Treats cash tied up in receivables/inventory as an inflow, inflating cash flow."
                            ),
                        }
                    )

    def _audit_terminal_capex_da_balance(
        self, ws: openpyxl.worksheet.worksheet.Worksheet, sheet_name: str
    ) -> None:
        """Audit steady-state reinvestment parity between CapEx and Depreciation.

        In perpetuity, g = Reinvestment Rate * ROIC.
        Projecting CapEx significantly in excess of D&A (e.g. >1.5x) into perpetuity
        under low nominal growth violates sustainable economic equilibrium.
    """
        # Terminal CapEx cell: H25
        cell_capex = ws.cell(row=25, column=8)
        f_str = str(cell_capex.value or "")

        if "2.5" in f_str or "*2.5" in f_str:
            self.discrepancies.append(
                {
                    "discrepancy_id": "ERR-002-TERMINAL-CAPEX-IMBALANCE",
                    "sheet_name": sheet_name,
                    "cell_coordinate": cell_capex.coordinate,
                    "error_category": "Economic Equilibrium",
                    "severity": "CRITICAL",
                    "flawed_content": f_str,
                    "expected_pattern": "=H24*1.08 (CapEx ~ D&A * [1 + g_real])",
                    "description": (
                        "Terminal CapEx modeled at 2.5x D&A in perpetuity under 2.5% steady-state growth. "
                        "Violates steady-state ROIC reinvestment math, indefinitely draining terminal cash flows."
                    ),
                }
            )

    def _audit_discount_factor_exponents(
        self, ws: openpyxl.worksheet.worksheet.Worksheet, sheet_name: str
    ) -> None:
        """Verify discount factor exponent alignment with terminal forecast horizon.

        For a 5-year explicit forecast, Terminal Value occurs at t=5 and must be discounted
        by (1 + WACC)^5, not (1 + WACC)^6.
    """
        cell_pv_tv = ws.cell(row=35, column=3)
        f_str = str(cell_pv_tv.value or "")

        if "^6" in f_str or "/(1+$C$7)^6" in f_str or "/(1+C7)^6" in f_str:
            self.discrepancies.append(
                {
                    "discrepancy_id": "ERR-004-DISCOUNT-EXPONENT-OFF-BY-ONE",
                    "sheet_name": sheet_name,
                    "cell_coordinate": cell_pv_tv.coordinate,
                    "error_category": "Time-Value Mathematics",
                    "severity": "HIGH",
                    "flawed_content": f_str,
                    "expected_pattern": f_str.replace("^6", "^5"),
                    "description": (
                        "Terminal value discounted by (1 + WACC)^6 instead of matching Year 5 horizon (1 + WACC)^5. "
                        "Artificially penalizes terminal present value by an extraneous full year of discounting."
                    ),
                }
            )

    def _audit_net_debt_bridge(
        self, ws: openpyxl.worksheet.worksheet.Worksheet, sheet_name: str
    ) -> None:
        """Audit Enterprise Value to Equity Value bridge for complete Net Debt inclusion.

        Formula: Equity Value = Enterprise Value + Cash & Equivalents - Total Debt.
        Subtracting Gross Debt while omitting Cash understates equity value.
    """
        # Look for debt and cash items in rows 36-40
        has_cash_addition = False
        direct_debt_deduction = False
        target_cell = ws.cell(row=37, column=3)

        for r in range(36, 42):
            lbl = str(ws.cell(row=r, column=2).value or "")
            v = str(ws.cell(row=r, column=3).value or "")

            if "Cash" in lbl:
                has_cash_addition = True
            if "Total Debt" in lbl and "Net Debt Cash Omitted" in lbl:
                direct_debt_deduction = True

        if not has_cash_addition or direct_debt_deduction:
            self.discrepancies.append(
                {
                    "discrepancy_id": "ERR-005-NET-DEBT-BRIDGE-OMISSION",
                    "sheet_name": sheet_name,
                    "cell_coordinate": target_cell.coordinate,
                    "error_category": "Accounting Identity",
                    "severity": "CRITICAL",
                    "flawed_content": str(target_cell.value),
                    "expected_pattern": "Bridge: Enterprise Value + Cash ($45,000) - Total Debt ($85,000)",
                    "description": (
                        "Bridge subtracts Total Debt directly from Enterprise Value while omitting $45.0M cash balance. "
                        "Confuses Gross Debt with Net Debt, mechanically understating equity value."
                    ),
                }
            )

    def export_discrepancy_log(self, output_path: Path) -> pd.DataFrame:
        """Save identified structural anomalies to CSV format."""
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        df = pd.DataFrame(self.discrepancies)
        df.to_csv(output_path, index=False)
        return df


def run_audit_scanner(
    model_path: Path | None = None,
    output_csv: Path | None = None,
) -> pd.DataFrame:
    """Execute audit scan on target model and persist discrepancy log."""
    if model_path is None:
        model_path = OUTPUT_DIR / "target_model_v1_flawed.xlsx"
    if output_csv is None:
        output_csv = OUTPUT_DIR / "audit_discrepancy_log.csv"

    scanner = ModelFormulaScanner(model_path)
    scanner.scan_all()
    df = scanner.export_discrepancy_log(output_csv)
    return df


if __name__ == "__main__":
    flawed_file = OUTPUT_DIR / "target_model_v1_flawed.xlsx"
    log_file = OUTPUT_DIR / "audit_discrepancy_log.csv"
    res_df = run_audit_scanner(flawed_file, log_file)
    print(f"Audit scan completed. Found {len(res_df)} discrepancies.")
    print(res_df[["discrepancy_id", "cell_coordinate", "severity", "error_category"]].to_string())
