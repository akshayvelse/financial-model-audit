"""Valuation impact bridge and error attribution waterfall engine.

Quantifies the discrete dollar-per-share and percentage equity distortion
caused by each of the 5 structural modeling errors, isolating their sequential
attribution and generating an institutional attribution waterfall chart.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Tuple

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

OUTPUT_DIR = Path(__file__).resolve().parent.parent / "output"


class ValuationBridgeEngine:
    """Calculates step-by-step valuation impact from flawed to audited baseline."""

    def __init__(self) -> None:
        # Standardized valuation parameters ($000s)
        self.wacc = 0.0950
        self.g = 0.0250
        self.tax_rate = 0.2100
        self.cash_balance = 45000.0
        self.total_debt = 85000.0
        self.shares_outstanding = 15000.0  # 000s shares

        self.base_revenue = 250000.0
        self.growth_rates = [0.080, 0.075, 0.070, 0.065, 0.060]
        self.cogs_rate = 0.580
        self.opex_rate = 0.220
        self.da_rate = 0.050
        self.capex_rate = 0.060
        self.nwc_rate = 0.015

    def evaluate_model_scenario(
        self,
        fix_nwc_sign: bool = False,
        fix_terminal_capex: bool = False,
        fix_cogs_override: bool = False,
        fix_discount_exponent: bool = False,
        fix_net_debt_cash: bool = False,
    ) -> dict[str, float]:
        """Evaluate DCF enterprise and equity values under specific error states."""
        # Project revenues
        revs = []
        curr_r = self.base_revenue
        for gr in self.growth_rates:
            curr_r *= 1.0 + gr
            revs.append(curr_r)
        revs_arr = np.array(revs)

        # Cost of Goods Sold
        cogs = revs_arr * self.cogs_rate
        if not fix_cogs_override:
            # Error 3: static hardcoded 185000 in Year 4
            cogs[3] = 185000.0

        gp = revs_arr - cogs
        opex = revs_arr * self.opex_rate
        ebit = gp - opex
        tax = ebit * self.tax_rate
        nopat = ebit - tax
        da = revs_arr * self.da_rate
        capex = revs_arr * self.capex_rate
        delta_nwc = revs_arr * self.nwc_rate

        # FCFF
        fcff = nopat + da - capex - delta_nwc
        if not fix_nwc_sign:
            # Error 1: added rather than subtracted in Years 2 and 3
            fcff[1] = nopat[1] + da[1] - capex[1] + delta_nwc[1]
            fcff[2] = nopat[2] + da[2] - capex[2] + delta_nwc[2]

        years = np.array([1, 2, 3, 4, 5])
        dfs = 1.0 / np.power(1.0 + self.wacc, years - 0.5)
        pv_explicit = float(np.sum(fcff * dfs))

        # Terminal Year
        term_rev = revs_arr[-1] * (1.0 + self.g)
        term_cogs = term_rev * self.cogs_rate
        term_gp = term_rev - term_cogs
        term_opex = term_rev * self.opex_rate
        term_ebit = term_gp - term_opex
        term_nopat = term_ebit * (1.0 - self.tax_rate)
        term_da = term_rev * self.da_rate

        if not fix_terminal_capex:
            # Error 2: 2.5x D&A in perpetuity
            term_capex = term_da * 2.5
        else:
            # Audited: 1.08x D&A steady-state replacement
            term_capex = term_da * 1.08

        term_delta_nwc = term_rev * self.nwc_rate
        term_fcff = term_nopat + term_da - term_capex - term_delta_nwc

        tv = term_fcff / (self.wacc - self.g)

        if not fix_discount_exponent:
            # Error 4: Discounted by (1 + WACC)^6
            pv_tv = tv / np.power(1.0 + self.wacc, 6.0)
        else:
            # Audited: Discounted by (1 + WACC)^5
            pv_tv = tv / np.power(1.0 + self.wacc, 5.0)

        ev = pv_explicit + pv_tv

        if not fix_net_debt_cash:
            # Error 5: Gross debt subtracted directly, cash omitted
            equity_val = ev - self.total_debt
        else:
            # Audited: Net debt = Total Debt - Cash
            equity_val = ev + self.cash_balance - self.total_debt

        share_price = equity_val / self.shares_outstanding

        return {
            "pv_explicit": pv_explicit,
            "tv": tv,
            "pv_tv": pv_tv,
            "enterprise_value": ev,
            "equity_value": equity_val,
            "share_price": share_price,
        }

    def compute_sequential_waterfall(self) -> pd.DataFrame:
        """Isolate incremental share price adjustments across all 5 corrections."""
        # Baseline flawed model
        s0 = self.evaluate_model_scenario(False, False, False, False, False)

        # Step 1: Repair Working Capital Sign Flip
        s1 = self.evaluate_model_scenario(True, False, False, False, False)

        # Step 2: Repair Terminal CapEx Imbalance
        s2 = self.evaluate_model_scenario(True, True, False, False, False)

        # Step 3: Repair Year 4 COGS Override
        s3 = self.evaluate_model_scenario(True, True, True, False, False)

        # Step 4: Repair Discount Exponent Off-by-One
        s4 = self.evaluate_model_scenario(True, True, True, True, False)

        # Step 5: Repair Net Debt Bridge Cash Omission (Audited Baseline)
        s5 = self.evaluate_model_scenario(True, True, True, True, True)

        p0 = s0["share_price"]
        p1 = s1["share_price"]
        p2 = s2["share_price"]
        p3 = s3["share_price"]
        p4 = s4["share_price"]
        p5 = s5["share_price"]

        steps = [
            {
                "step_name": "Inherited Model (Flawed V1)",
                "share_price": p0,
                "delta": 0.0,
                "pct_impact": 0.0,
                "category": "Baseline",
            },
            {
                "step_name": "1. Correct Working Capital Sign Flip",
                "share_price": p1,
                "delta": p1 - p0,
                "pct_impact": ((p1 - p0) / p0) * 100.0,
                "category": "Cash Flow Correction",
            },
            {
                "step_name": "2. Balance Terminal CapEx / D&A",
                "share_price": p2,
                "delta": p2 - p1,
                "pct_impact": ((p2 - p1) / p0) * 100.0,
                "category": "Terminal Valuation",
            },
            {
                "step_name": "3. Restore Year 4 COGS Dynamic Formula",
                "share_price": p3,
                "delta": p3 - p2,
                "pct_impact": ((p3 - p2) / p0) * 100.0,
                "category": "Formula Integrity",
            },
            {
                "step_name": "4. Fix Terminal Discount Exponent (t=5)",
                "share_price": p4,
                "delta": p4 - p3,
                "pct_impact": ((p4 - p3) / p0) * 100.0,
                "category": "Discount Timing",
            },
            {
                "step_name": "5. Include Balance Sheet Cash in Bridge",
                "share_price": p5,
                "delta": p5 - p4,
                "pct_impact": ((p5 - p4) / p0) * 100.0,
                "category": "Balance Sheet Bridge",
            },
            {
                "step_name": "Audited Model (Corrected V2)",
                "share_price": p5,
                "delta": p5 - p0,
                "pct_impact": ((p5 - p0) / p0) * 100.0,
                "category": "Final Target",
            },
        ]

        return pd.DataFrame(steps)


def plot_error_attribution_waterfall(
    bridge_df: pd.DataFrame, output_path: Path
) -> None:
    """Render publication-grade valuation waterfall chart."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    labels = [
        "Flawed V1",
        "1. NWC Sign",
        "2. CapEx Imbalance",
        "3. COGS Override",
        "4. Discount Exponent",
        "5. Net Debt Cash",
        "Audited V2",
    ]

    p0 = bridge_df.loc[0, "share_price"]
    d1 = bridge_df.loc[1, "delta"]
    d2 = bridge_df.loc[2, "delta"]
    d3 = bridge_df.loc[3, "delta"]
    d4 = bridge_df.loc[4, "delta"]
    d5 = bridge_df.loc[5, "delta"]
    p_final = bridge_df.loc[6, "share_price"]

    # Waterfall bottom baseline and bar heights
    bottoms = [0.0, p0, p0 + d1, p0 + d1 + d2, p0 + d1 + d2 + d3, p0 + d1 + d2 + d3 + d4, 0.0]
    heights = [p0, d1, d2, d3, d4, d5, p_final]

    # Adjust bottoms if negative delta
    for i in range(1, 6):
        if heights[i] < 0:
            bottoms[i] += heights[i]
            heights[i] = abs(heights[i])

    colors = [
        "#1B365D",  # Navy for baseline
        "#C00000" if d1 < 0 else "#2E75B6",  # Red for negative delta
        "#008000" if d2 > 0 else "#C00000",  # Green for positive delta
        "#008000" if d3 > 0 else "#C00000",
        "#008000" if d4 > 0 else "#C00000",
        "#008000" if d5 > 0 else "#C00000",
        "#1B365D",  # Navy for final
    ]

    fig, ax = plt.subplots(figsize=(11, 6), dpi=300)

    bars = ax.bar(labels, heights, bottom=bottoms, color=colors, width=0.55, edgecolor="black", linewidth=0.8)

    # Add numeric callouts
    for idx, (b, h, d) in enumerate(zip(bottoms, heights, [p0, d1, d2, d3, d4, d5, p_final])):
        top_y = b + h
        if idx in [0, 6]:
            txt = f"${d:.2f}"
            y_pos = top_y + 0.8
        else:
            sign = "+" if d > 0 else "-"
            txt = f"{sign}${abs(d):.2f}"
            y_pos = top_y + 0.8

        ax.text(idx, y_pos, txt, ha="center", va="bottom", fontsize=9, fontweight="bold")

    ax.set_ylabel("Implied Target Share Price (USD)", fontsize=10, fontweight="bold")
    ax.set_title(
        "Financial Model Audit: Valuation Error Attribution Waterfall\nBridge from Flawed V1 ($) to Audited V2 ($)",
        fontsize=12,
        fontweight="bold",
        pad=15,
    )
    ax.set_ylim(0, max(p0, p_final) * 1.25)
    ax.grid(axis="y", linestyle="--", alpha=0.4)

    plt.setp(ax.get_xticklabels(), rotation=20, ha="right", fontsize=9, fontweight="bold")

    fig.tight_layout()
    fig.savefig(output_path, dpi=300)
    plt.close(fig)


def run_valuation_bridge() -> tuple[pd.DataFrame, Path]:
    """Execute valuation bridge calculation and export attribution chart."""
    engine = ValuationBridgeEngine()
    df = engine.compute_sequential_waterfall()
    img_path = OUTPUT_DIR / "error_attribution_waterfall.png"
    plot_error_attribution_waterfall(df, img_path)
    return df, img_path


if __name__ == "__main__":
    waterfall_df, chart_file = run_valuation_bridge()
    print("=== FINANCIAL MODEL AUDIT VALUATION BRIDGE ===")
    print(waterfall_df[["step_name", "share_price", "delta", "pct_impact"]].to_string(index=False))
    print(f"\nWaterfall chart saved to: {chart_file}")
