# FINANCIAL MODEL AUDIT & FORENSIC VALUATION REPORT

**TARGET ASSET:** Corporate DCF Valuation Model (`target_model_v1_flawed.xlsx`)  
**AUDIT ENGAGEMENT:** Buy-Side Pre-Acquisition Financial Model Integrity Review  
**INITIAL FLAWED VALUATION:** $20.03 per share (Equity Value: $300.4M)  
**AUDITED CORRECTED VALUATION:** $38.48 per share (Equity Value: $577.2M)  
**TOTAL VALUATION DISTORTION:** -$18.45 per share (-47.9% net under-valuation in inherited model)  
**DATE:** November 8, 2026  

---

## 1. EXECUTIVE AUDIT SUMMARY & VALUATION RECONCILIATION

A pre-transaction forensic audit of the inherited analyst financial model revealed 5 material structural errors spanning cash flow accounting, steady-state economic equilibrium, formula integrity, discounting conventions, and enterprise-to-equity value bridge mechanics. 

Cumulatively, these defects suppressed implied equity value by **$276.8M** ($18.45 per share), representing a 92.1% upward correction upon forensic remediation:

| Step / Error Classification | Pre-Fix Price | Post-Fix Price | $/Share Impact | % Equity Distortion | Primary Financial Principle Violated |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **Inherited Flawed Model (V1)** | — | **$20.03** | — | — | Baseline inherited flawed model |
| **Error 1: NWC Sign Flip** | $20.03 | $19.03 | -$1.00 | -5.0% | Cash Flow Statement Sign Convention |
| **Error 2: Terminal CapEx Imbalance** | $19.03 | $33.12 | +$14.10 | +70.4% | Steady-State ROIC Reinvestment Parity |
| **Error 3: Year 4 COGS Override** | $33.12 | $32.86 | -$0.26 | -1.3% | Dynamic Forecast Formula Integrity |
| **Error 4: Discount Exponent Off-by-One** | $32.86 | $35.48 | +$2.62 | +13.1% | Discount Period Temporal Alignment ($t=5$) |
| **Error 5: Net Debt Bridge Omission** | $35.48 | **$38.48** | +$3.00 | +15.0% | Enterprise-to-Equity Identity ($EV - Net Debt$) |
| **Total Net Bridge Delta** | — | — | **+$18.45** | **+92.1%** | Sum of discrete sequential adjustments |

---

## 2. DETAILED FORENSIC DISCREPANCY ANALYSIS

### Finding 1: Working Capital Sign Flip in FCFF Formula
- **Location:** Sheet `DCF Valuation`, Cells `D27` and `E27` (Forecast Years 2 and 3).
- **Flawed Formula:** `=D23+D24-D25+D26` (Adding Change in Non-Cash Working Capital).
- **Corrected Formula:** `=D23+D24-D25-D26` (Subtracting Change in Non-Cash Working Capital).
- **Technical Root Cause:** An expansion in working capital (growth in receivables and inventory outpacing trade payables) represents cash consumed to support incremental revenue. Adding $\Delta\text{NWC}$ mistook an operating cash absorption for a liquidity generation event.
- **Valuation Distortion:** Overstated intermediate cash flows by $4.4M in Year 2 and $4.7M in Year 3. Correcting this reduced intermediate present value by **$15.0M** (-$1.00 per share).

### Finding 2: Terminal Year CapEx / D&A Reinvestment Imbalance
- **Location:** Sheet `DCF Valuation`, Cell `H25` (Terminal Year CapEx).
- **Flawed Formula:** `=H24*2.5` (CapEx modeled at 250.0% of D&A into perpetuity).
- **Corrected Formula:** `=H24*1.08` (CapEx aligned at 108.0% of D&A to support 2.5% terminal nominal growth).
- **Technical Root Cause:** In sustainable equilibrium, the economic reinvestment rate must equal $g / \text{ROIC}$. Assuming CapEx remains at 2.5x D&A in perpetuity while the business grows at only 2.5% implies an uneconomic destruction of capital where net fixed assets expand by 12% annually with zero incremental revenue.
- **Valuation Distortion:** Artificially depressed Terminal Year FCFF from $37.5M to $19.4M, wiping out $258.6M in undiscounted terminal value. Correcting this added **+$14.10 per share** (+70.4% equity uplift).

### Finding 3: Hardcoded Numeric Override in Formula Row
- **Location:** Sheet `DCF Valuation`, Cell `F18` (Year 4 Cost of Goods Sold).
- **Flawed Cell Content:** Hardcoded constant `185000` (overwriting the dynamic column formula).
- **Corrected Formula:** `=F17*0.580` (Linked dynamically to Year 4 Revenue).
- **Technical Root Cause:** An analyst manually plugged a preliminary COGS estimate during prior scenario iterations, severing the link to the 58.0% gross cost assumption. When top-line volume was revised, COGS failed to flex, understating costs by $3.9M.
- **Valuation Distortion:** Overstated Year 4 NOPAT by $3.1M. Restoring dynamic formula integrity corrected the model by **-$0.26 per share**.

### Finding 4: Terminal Value Discount Factor Exponent Off-by-One
- **Location:** Sheet `DCF Valuation`, Cell `C35` (PV of Terminal Value).
- **Flawed Formula:** `=C34/(1+$C$7)^6` (Discounting horizon $t=6$).
- **Corrected Formula:** `=C34/(1+$C$7)^5` (Discounting horizon $t=5$).
- **Technical Root Cause:** Terminal value computed at $t=5$ via the perpetuity formula represents the present value of all cash flows beyond Year 5 as of the end of Year 5. Applying a power of 6 subjected the entire terminal value to an extraneous full year of discounting at 9.50% WACC.
- **Valuation Distortion:** Undervalued terminal present value by 8.7% ($39.3M). Aligning the exponent to Year 5 restored **+$2.62 per share** (+13.1% equity uplift).

### Finding 5: Enterprise Value to Equity Value Bridge Cash Omission
- **Location:** Sheet `DCF Valuation`, Cells `C37:C40`.
- **Flawed Formula:** `Implied Equity = Enterprise Value - Total Debt ($85,000)`.
- **Corrected Formula:** `Implied Equity = Enterprise Value + Cash ($45,000) - Total Debt ($85,000)`.
- **Technical Root Cause:** The bridge confused Total Debt with Net Debt. By subtracting gross debt while omitting the $45.0M in balance sheet cash and marketable securities, the analyst failed to credit equity holders for non-operating liquidity.
- **Valuation Distortion:** Direct dollar-for-dollar understatement of equity value by exactly $45.0M, or **+$3.00 per share** across 15.0M diluted shares.

---

## 3. AUDIT RECOMMENDATIONS & INTERNAL CONTROLS

1. **Automated Workbook Linting:** Embed automated continuous-integration scanning scripts (`formula_scanner.py`) in repository hooks to reject workbooks containing hardcoded constants in homogeneous formula rows.
2. **Reinvestment Parity Checks:** Mandate automated validation enforcing that terminal CapEx-to-D&A multiples satisfy $\frac{\text{CapEx}}{\text{D&A}} \le 1.0 + \frac{g}{\text{Asset Life}}$ prior to final committee presentation.
3. **Formal Balance Sheet Bridge Reconciliation:** Enforce strict 3-way reconciliation linking cash and debt balances directly to balance sheet tabs, prohibiting standalone manual debt deductions.
