# Financial Model Audit & Integrity Engine

An automated spreadsheet auditing and valuation forensic engine that detects structural modeling errors in financial workbooks using `openpyxl`, exports discrepancy logs, and isolates exact dollar-per-share valuation impacts via sequential waterfall attribution.

## Directory Architecture

```
02_financial_model_audit/
├── README.md
├── audit_findings_report.md
├── src/
│   ├── __init__.py
│   ├── generate_models.py
│   ├── formula_scanner.py
│   └── valuation_bridge.py
└── output/
    ├── target_model_v1_flawed.xlsx
    ├── target_model_v2_audited.xlsx
    ├── audit_discrepancy_log.csv
    └── error_attribution_waterfall.png
```

## Audit Scope & Injected Model Errors

The audit suite evaluates two programmatic Excel models:
1. `output/target_model_v1_flawed.xlsx`: Inherited analyst model embedding 5 structural modeling flaws.
2. `output/target_model_v2_audited.xlsx`: Production-grade audited workbook resolving all errors.

### The 5 Structural Errors Inspected:
1. **Working Capital Sign Flip (`ERR-001`):** Change in Non-Cash Working Capital added rather than subtracted in Free Cash Flow to Firm (FCFF) formulas in Years 2 and 3 (-$1.00/share impact).
2. **Terminal Year CapEx / D&A Imbalance (`ERR-002`):** Terminal CapEx set to 2.5x D&A in perpetuity under 2.5% steady-state growth, violating steady-state ROIC reinvestment parity (+$14.10/share impact).
3. **Hardcoded Override inside Formula Row (`ERR-003`):** Year 4 Cost of Goods Sold overwritten with static numeric constant `185000`, severing dynamic revenue linkage (-$0.26/share impact).
4. **Discount Factor Off-by-One (`ERR-004`):** Terminal Value discounted by $(1 + \text{WACC})^6$ instead of matching the Year 5 terminal horizon $(1 + \text{WACC})^5$ (+$2.62/share impact).
5. **Net Debt Bridge Cash Omission (`ERR-005`):** Gross debt subtracted from Enterprise Value while omitting $45.0M balance sheet cash (+$3.00/share impact).

## Execution

```bash
# Generate flawed (v1) and audited (v2) Excel workbooks
python 02_financial_model_audit/src/generate_models.py

# Run automated formula scanner and export discrepancy log
python 02_financial_model_audit/src/formula_scanner.py

# Compute step-by-step valuation bridge and generate waterfall chart
python 02_financial_model_audit/src/valuation_bridge.py
```

## Summary of Valuation Bridge Attribution

- **Inherited Flawed Share Price:** $20.03
- **Audited Share Price:** $38.48
- **Net Correction:** +$18.45 per share (+92.1% equity uplift)
- **Error Attribution Waterfall Chart:** `output/error_attribution_waterfall.png`
- **Machine-Readable Discrepancy Log:** `output/audit_discrepancy_log.csv`
