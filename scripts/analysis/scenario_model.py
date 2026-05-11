"""
Transform SR Holding Management LLC — 18-Month Forward Scenario Model
======================================================================
Reads confirmed actuals from the analysis JSON files, then projects
June 2026 – November 2027 under Bear / Base / Bull assumptions.

Usage:
    python3 scenario_model.py                    # run all 3 scenarios
    python3 scenario_model.py --scenario bull    # single scenario
    python3 scenario_model.py --csv              # CSV output
    python3 scenario_model.py --custom           # edit CUSTOM_PARAMS below

All monetary values in $M unless stated.

Data sources (actuals):
  - analysis/analysis_disbursements_2024_2026.json  (BU beta actuals)
  - analysis/analysis_inflows_trend.json             (Inflows Actuals)
  - analysis/analysis_payroll_headcount.json
  - analysis/analysis_liquidity.json
"""

import json, os, sys, argparse
from datetime import date, timedelta
from dateutil.relativedelta import relativedelta  # pip install python-dateutil

# ── Paths ────────────────────────────────────────────────────────────────────
# Adjust BASE_DIR if running from a different location
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.abspath(os.path.join(SCRIPT_DIR, "..", ".."))
ANALYSIS_DIR = os.path.join(REPO_ROOT, "analysis")

# ── Confirmed May 2026 actuals (starting point) ──────────────────────────────
MAY_2026_ACTUALS = {
    "available_cash_M":      1.51,   # Cash Flow tab, as of May 8 2026
    "total_cash_M":         89.59,   # Available + unavailable
    "unavailable_cash_M":   82.84,   # LC collateral + reserves + holdbacks
    # Operating inflows (monthly run-rate from May 2026 data)
    "home_services_M":      26.0,    # Jan-Apr 2026 avg: ~$25-27M/mo
    "kmart_stores_M":        6.9,    # Apr 2026 actual (monthly)
    "sears_stores_M":        1.2,    # FY26 monthly avg ~$1.3M
    "kcd_royalties_M":       2.5,    # Starting low; ramp assumption varies
    "syw_inflows_M":         3.9,    # ~$47M/yr ÷ 12
    "online_M":              1.0,    # FY26 annualized ~$9M
    "supply_chain_M":        2.8,    # ~$33M/yr ÷ 12
    "tenant_income_M":       0.6,    # ~$7M/yr ÷ 12
    "rx_rebates_M":          0.1,    # Sporadic; small monthly average
    "cchs_b2b_M":            0.0,    # Terminated May 2026
    "citi_reimb_M":          0.0,    # Wound down
    "asset_sales_M":         0.0,    # Event-driven; modeled separately
    "debt_financing_M":      0.0,    # Event-driven; modeled separately
    # Disbursements (monthly run-rate from Q1 2026 actuals)
    "payroll_M":           -19.1,    # FY26 Jan-Apr avg
    "merch_M":             -12.4,    # Q1 2026 avg
    "logistics_M":          -1.2,    # Q1 2026 avg (post-rationalization)
    "rent_facilities_M":    -5.0,    # Q1 2026 avg
    "interest_cash_M":       0.0,    # Q1 2026 avg ~$0 (PIK active)
    "debt_repayment_M":      0.0,    # Episodic; modeled below
    "other_disb_M":         -8.5,    # Systems + legal + SG&A + other
}

# ── Scenario Definitions ─────────────────────────────────────────────────────
#
# Each scenario has:
#   - monthly_changes: dict of {month_offset: {field: new_value_or_delta}}
#   - one_time_events: list of {month: offset, field: ..., amount: ...}
#   - growth_rates:    dict of {field: monthly_growth_rate}  (e.g. 0.02 = +2%/mo)
#
# month_offset = months after May 2026 (1 = June 2026, 18 = Nov 2027)

SCENARIOS = {

    # ────────────────────────────────────────────────────────────────────────
    "bear": {
        "description": (
            "BEAR: KCD licensing fails to ramp. HS accelerates decline. "
            "Asset sales slip. Interest becomes cash-pay in Q4 2026. "
            "No new TopCo injection."
        ),
        "color": "🔴",
        "growth_rates": {
            "home_services_M":    -0.018,   # -1.8%/mo = ~-20%/yr (accelerated decline)
            "kmart_stores_M":     -0.012,   # -1.2%/mo
            "sears_stores_M":     -0.030,   # -3.0%/mo (near-zero by FY27)
            "kcd_royalties_M":     0.000,   # Stays at $2.5M/mo (no ramp)
            "syw_inflows_M":      -0.005,   # Slight decline
            "online_M":           -0.020,   # Continued decline
        },
        "monthly_changes": {
            # Payroll stays flat (no further cuts)
        },
        "one_time_events": [
            # Cash interest resumes Q4 2026 ($15M semi-annual → ~$7.5M/mo in Oct + Mar)
            {"month": 5,  "field": "interest_cash_M", "amount": -15.0,
             "note": "Semi-annual cash interest payment restarts (Oct 2026)"},
            {"month": 10, "field": "interest_cash_M", "amount": -15.0,
             "note": "Semi-annual cash interest payment (Mar 2027)"},
            # Asset sales: only $10M total, in two tranches
            {"month": 2,  "field": "asset_sales_M",   "amount": 5.0,
             "note": "Single asset sale close"},
            {"month": 8,  "field": "asset_sales_M",   "amount": 5.0,
             "note": "Single asset sale close"},
        ],
        "structural_changes": {
            # No payroll reduction, no LC release, no new capital
        },
    },

    # ────────────────────────────────────────────────────────────────────────
    "base": {
        "description": (
            "BASE: KCD licensing ramps as forecast. HS declines ~15%/yr. "
            "$37.7M asset sales close on schedule. PIK continues. "
            "PNC LC released in Q3 2026. Payroll drifts to $16M by Q2 2027."
        ),
        "color": "🟡",
        "growth_rates": {
            "home_services_M":    -0.013,   # -1.3%/mo = ~-15%/yr
            "kmart_stores_M":     -0.006,   # -0.6%/mo = -7%/yr
            "sears_stores_M":     -0.025,   # -2.5%/mo
            "kcd_royalties_M":     0.025,   # +2.5%/mo ramping to ~$5M/mo
            "syw_inflows_M":       0.008,   # Credit card growth
            "online_M":           -0.015,
        },
        "monthly_changes": {
            # Payroll drifts down $0.25M/mo through mid-2027
            **{m: {"payroll_delta": -0.25} for m in range(1, 13)},
        },
        "one_time_events": [
            # PNC LC release (+$5.6M, month 3 = Aug 2026)
            {"month": 3,  "field": "available_cash_delta", "amount": 5.6,
             "note": "PNC LC facility terminated — $5.6M cash collateral released"},
            # Asset sales: $37.7M across 5 events
            {"month": 2,  "field": "asset_sales_M", "amount": 8.0,
             "note": "Asset sale close #1"},
            {"month": 4,  "field": "asset_sales_M", "amount": 10.0,
             "note": "Asset sale close #2"},
            {"month": 6,  "field": "asset_sales_M", "amount": 9.7,
             "note": "Asset sale close #3"},
            {"month": 9,  "field": "asset_sales_M", "amount": 5.0,
             "note": "Asset sale close #4"},
            {"month": 12, "field": "asset_sales_M", "amount": 5.0,
             "note": "Asset sale close #5"},
            # Semi-annual interest: stays ~$0.5M/mo (PIK continues, small residual)
        ],
        "structural_changes": {},
    },

    # ────────────────────────────────────────────────────────────────────────
    "bull": {
        "description": (
            "BULL: KCD licensing exceeds forecast ($55M+ FY26). HS decline "
            "slows to -10%/yr. IHR contract book sale announced ($75M). "
            "PNC + partial UBS LC released (+$21.5M). SYW card hits 100K accounts. "
            "Payroll reaches $15M by Q4 2026."
        ),
        "color": "🟢",
        "growth_rates": {
            "home_services_M":    -0.009,   # -0.9%/mo = ~-10%/yr (slowdown)
            "kmart_stores_M":     -0.004,
            "sears_stores_M":     -0.020,
            "kcd_royalties_M":     0.045,   # Strong ramp: $2.5M → $8M+ by FY27
            "syw_inflows_M":       0.018,   # Credit card 100K accounts
            "online_M":           -0.010,
        },
        "monthly_changes": {
            # Aggressive payroll reduction: -$0.5M/mo
            **{m: {"payroll_delta": -0.5} for m in range(1, 13)},
        },
        "one_time_events": [
            # PNC LC release (month 2)
            {"month": 2,  "field": "available_cash_delta", "amount": 5.6,
             "note": "PNC LC facility terminated — $5.6M released"},
            # Partial UBS LC release (month 4)
            {"month": 4,  "field": "available_cash_delta", "amount": 15.9,
             "note": "Partial UBS LC collateral reduction — $15.9M released"},
            # Asset sales: $60M total
            {"month": 2,  "field": "asset_sales_M", "amount": 12.0},
            {"month": 4,  "field": "asset_sales_M", "amount": 15.0},
            {"month": 6,  "field": "asset_sales_M", "amount": 18.0},
            {"month": 9,  "field": "asset_sales_M", "amount": 15.0},
            # IHR contract book sale (month 8 = Jan 2027)
            {"month": 8,  "field": "available_cash_delta", "amount": 75.0,
             "note": "IHR contract portfolio sale closes"},
            # Small debt repayment after IHR proceeds
            {"month": 9,  "field": "debt_repayment_M", "amount": -20.0,
             "note": "Voluntary debt paydown from IHR proceeds"},
        ],
        "structural_changes": {},
    },
}

# ── Custom scenario override ──────────────────────────────────────────────────
# Edit these values to run your own scenario (use --custom flag)
CUSTOM_PARAMS = {
    "description": "CUSTOM: Edit CUSTOM_PARAMS in script to define your assumptions",
    "color": "🔵",
    "kcd_monthly_start_M":      2.5,   # KCD inflows starting point
    "kcd_monthly_growth_rate":  0.025, # Monthly growth rate
    "hs_monthly_decline_rate":  0.013, # Monthly decline rate (positive = decline)
    "asset_sales_total_M":      37.7,  # Total asset sales over period
    "pnc_lc_release_M":         5.6,   # One-time cash from PNC LC release
    "ubs_lc_release_M":         0.0,   # One-time cash from UBS LC partial release
    "ihr_sale_proceeds_M":      0.0,   # IHR contract book sale (0 = no sale)
    "ihr_sale_month":           0,     # Month offset for IHR sale
    "topco_injection_M":        0.0,   # New TopCo equity injection
    "topco_injection_month":    0,
    "payroll_monthly_reduction": 0.25, # Monthly payroll reduction ($M)
    "interest_resumes_month":   0,     # 0 = never; 5 = Oct 2026 etc.
    "interest_semiannual_M":    15.0,  # Amount if interest resumes
}


# ── Projection Engine ─────────────────────────────────────────────────────────

def run_scenario(name, scenario_def, n_months=18, start_month=None):
    """
    Project n_months forward from May 2026 baseline.
    Returns list of monthly dicts with all line items.
    """
    if start_month is None:
        start_month = date(2026, 6, 1)

    # Deep copy actuals as mutable state
    state = dict(MAY_2026_ACTUALS)
    results = []

    growth = scenario_def.get("growth_rates", {})
    changes = scenario_def.get("monthly_changes", {})
    events = scenario_def.get("one_time_events", [])

    # Build event lookup: month_offset -> list of events
    event_map = {}
    for ev in events:
        m = ev["month"]
        event_map.setdefault(m, []).append(ev)

    payroll_base = state["payroll_M"]  # track separately for delta logic

    for offset in range(1, n_months + 1):
        month_date = start_month + relativedelta(months=offset - 1)
        month_label = month_date.strftime("%b %Y")

        # Apply monthly growth rates to inflow lines
        for field, rate in growth.items():
            if field in state:
                state[field] *= (1 + rate)

        # Apply monthly structural changes
        if offset in changes:
            mc = changes[offset]
            if "payroll_delta" in mc:
                payroll_base = max(payroll_base + mc["payroll_delta"],
                                   -15.0)  # floor at $15M
                state["payroll_M"] = payroll_base

        # Sum operating inflows
        inflow_fields = [
            "home_services_M", "kmart_stores_M", "sears_stores_M",
            "kcd_royalties_M", "syw_inflows_M", "online_M",
            "supply_chain_M", "tenant_income_M", "rx_rebates_M",
            "cchs_b2b_M", "citi_reimb_M",
        ]
        total_op_inflows = sum(max(state.get(f, 0), 0) for f in inflow_fields)

        # Sum disbursements
        disb_fields = [
            "payroll_M", "merch_M", "logistics_M", "rent_facilities_M",
            "interest_cash_M", "debt_repayment_M", "other_disb_M",
        ]
        total_disb = sum(min(state.get(f, 0), 0) for f in disb_fields)

        # Net operating CF
        net_op_cf = total_op_inflows + total_disb

        # One-time events this month
        event_inflows = 0.0
        event_notes = []
        cash_adjustments = 0.0
        if offset in event_map:
            for ev in event_map[offset]:
                f = ev["field"]
                amt = ev["amount"]
                note = ev.get("note", "")
                if f == "asset_sales_M":
                    event_inflows += amt
                    event_notes.append(f"Asset sale +${amt:.1f}M")
                elif f == "available_cash_delta":
                    cash_adjustments += amt
                    event_notes.append(note or f"One-time +${amt:.1f}M")
                elif f == "interest_cash_M":
                    total_disb += amt
                    net_op_cf += amt
                    event_notes.append(note or f"Interest payment ${amt:.1f}M")
                elif f == "debt_repayment_M":
                    total_disb += amt
                    net_op_cf += amt
                    event_notes.append(note or f"Debt repayment ${amt:.1f}M")

        # Total cash flow
        total_net_cf = net_op_cf + event_inflows

        # Available cash (carry forward from prior month)
        if offset == 1:
            prior_avail = MAY_2026_ACTUALS["available_cash_M"]
        else:
            prior_avail = results[-1]["ending_available_cash_M"]

        ending_avail = prior_avail + total_net_cf + cash_adjustments
        # Floor at 0 (model convention — negative = requires external rescue)
        is_negative = ending_avail < 0
        ending_avail_display = ending_avail  # show negative for stress signal

        results.append({
            "month":                     month_label,
            "month_offset":              offset,
            "total_op_inflows_M":        round(total_op_inflows, 2),
            "home_services_M":           round(state["home_services_M"], 2),
            "kcd_royalties_M":           round(state["kcd_royalties_M"], 2),
            "kmart_stores_M":            round(state["kmart_stores_M"], 2),
            "sears_stores_M":            round(state["sears_stores_M"], 2),
            "syw_inflows_M":             round(state["syw_inflows_M"], 2),
            "asset_sales_M":             round(event_inflows, 2),
            "total_disb_M":              round(total_disb, 2),
            "payroll_M":                 round(state["payroll_M"], 2),
            "net_op_cf_M":               round(net_op_cf, 2),
            "one_time_cash_M":           round(cash_adjustments, 2),
            "total_net_cf_M":            round(total_net_cf + cash_adjustments, 2),
            "ending_available_cash_M":   round(ending_avail_display, 2),
            "liquidity_flag":            "⚠️ CRISIS" if ending_avail_display < 2.0
                                         else ("⚡ WARNING" if ending_avail_display < 5.0
                                               else "✅ OK"),
            "event_notes":               "; ".join(event_notes),
        })

    return results


# ── Output Formatters ──────────────────────────────────────────────────────────

def print_scenario(name, scenario_def, results):
    color = scenario_def.get("color", "")
    desc = scenario_def.get("description", "")
    print(f"\n{'='*78}")
    print(f"{color}  SCENARIO: {name.upper()}")
    print(f"{'='*78}")
    print(f"  {desc}")
    print()
    print(f"  {'Month':<10} {'Op Inflows':>11} {'HS':>7} {'KCD':>7} {'Kmart':>7} "
          f"{'Disb':>8} {'Asset$':>7} {'1x Cash':>8} {'Net CF':>8} {'Avail $':>9}  Flag")
    print(f"  {'-'*10} {'-'*11} {'-'*7} {'-'*7} {'-'*7} "
          f"{'-'*8} {'-'*7} {'-'*8} {'-'*8} {'-'*9}  ----")

    for r in results:
        flag_str = r["liquidity_flag"]
        print(f"  {r['month']:<10} "
              f"${r['total_op_inflows_M']:>9.1f}  "
              f"${r['home_services_M']:>5.1f}  "
              f"${r['kcd_royalties_M']:>5.1f}  "
              f"${r['kmart_stores_M']:>5.1f}  "
              f"${r['total_disb_M']:>6.1f}  "
              f"${r['asset_sales_M']:>5.1f}  "
              f"${r['one_time_cash_M']:>6.1f}  "
              f"${r['total_net_cf_M']:>6.1f}  "
              f"${r['ending_available_cash_M']:>7.2f}  "
              f"{flag_str}"
              + (f"  ← {r['event_notes']}" if r["event_notes"] else "")
              )

    # Summary stats
    final_avail = results[-1]["ending_available_cash_M"]
    crisis_months = [r["month"] for r in results if r["ending_available_cash_M"] < 2.0]
    min_avail = min(r["ending_available_cash_M"] for r in results)
    max_avail = max(r["ending_available_cash_M"] for r in results)
    total_asset_sales = sum(r["asset_sales_M"] for r in results)
    total_one_time = sum(r["one_time_cash_M"] for r in results)

    print()
    print(f"  SUMMARY:")
    print(f"    Starting available cash (May 2026):   ${MAY_2026_ACTUALS['available_cash_M']:.2f}M")
    print(f"    Ending available cash (Nov 2027):     ${final_avail:.2f}M")
    print(f"    Min available cash in period:         ${min_avail:.2f}M")
    print(f"    Max available cash in period:         ${max_avail:.2f}M")
    print(f"    Months in CRISIS zone (<$2M):         {len(crisis_months)}")
    if crisis_months:
        print(f"    Crisis months:                        {', '.join(crisis_months)}")
    print(f"    Total asset sale proceeds:            ${total_asset_sales:.1f}M")
    print(f"    Total one-time cash releases:         ${total_one_time:.1f}M")


def print_comparison(all_results):
    """Side-by-side available cash comparison."""
    print(f"\n\n{'='*78}")
    print("  AVAILABLE CASH COMPARISON — ALL SCENARIOS")
    print(f"{'='*78}")
    scenarios = list(all_results.keys())
    header = f"  {'Month':<10}" + "".join(f"  {s.upper():>10}" for s in scenarios)
    print(header)
    print(f"  {'-'*10}" + "".join(f"  {'-'*10}" for _ in scenarios))

    n_months = min(len(r) for r in all_results.values())
    for i in range(n_months):
        month = list(all_results.values())[0][i]["month"]
        row = f"  {month:<10}"
        for s in scenarios:
            val = all_results[s][i]["ending_available_cash_M"]
            flag = " ⚠" if val < 2.0 else (" ⚡" if val < 5.0 else "  ")
            row += f"  ${val:>7.2f}M{flag}"
        print(row)

    print()
    print("  KEY LEVERS (delta Bear→Bull by Nov 2027):")
    bear_final = all_results["bear"][-1]["ending_available_cash_M"]
    bull_final = all_results["bull"][-1]["ending_available_cash_M"]
    base_final = all_results["base"][-1]["ending_available_cash_M"]
    print(f"    Bear vs Base:  ${base_final - bear_final:+.1f}M")
    print(f"    Base vs Bull:  ${bull_final - base_final:+.1f}M")
    print(f"    Bear vs Bull:  ${bull_final - bear_final:+.1f}M")
    print()
    print("  TOP SWING FACTORS (Bear→Bull):")
    print("    1. KCD licensing ramp:     $5M→$55M/yr        = ~$40M/yr swing")
    print("    2. IHR contract book sale: $0→$75M one-time   = $75M one-time")
    print("    3. LC collateral release:  $0→$21.5M          = $21.5M one-time")
    print("    4. HS decline rate:        -20%→-10%/yr       = ~$15-25M/yr swing")
    print("    5. PIK continuation:       +$0→-$30M/yr risk  = $30M/yr swing")
    print("    6. Asset sales pace:       $10M→$60M          = $50M timing swing")


def export_csv(all_results, path="scenario_output.csv"):
    import csv
    rows = []
    for scenario, results in all_results.items():
        for r in results:
            row = {"scenario": scenario}
            row.update(r)
            rows.append(row)
    if rows:
        with open(path, "w", newline="") as f:
            w = csv.DictWriter(f, fieldnames=rows[0].keys())
            w.writeheader()
            w.writerows(rows)
        print(f"\nCSV saved: {path}")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    try:
        from dateutil.relativedelta import relativedelta
    except ImportError:
        print("ERROR: python-dateutil required. Run: pip install python-dateutil")
        sys.exit(1)

    parser = argparse.ArgumentParser(description="Transform SR Scenario Model")
    parser.add_argument("--scenario", choices=["bear", "base", "bull", "all"],
                        default="all")
    parser.add_argument("--months", type=int, default=18)
    parser.add_argument("--csv", action="store_true")
    args = parser.parse_args()

    print("\nTransform SR Holding Management LLC — Forward Scenario Model")
    print(f"Base date: May 2026  |  Projection: {args.months} months")
    print(f"Starting available cash: ${MAY_2026_ACTUALS['available_cash_M']:.2f}M")

    scenarios_to_run = (
        ["bear", "base", "bull"] if args.scenario == "all"
        else [args.scenario]
    )

    all_results = {}
    for name in scenarios_to_run:
        results = run_scenario(name, SCENARIOS[name], n_months=args.months)
        all_results[name] = results
        print_scenario(name, SCENARIOS[name], results)

    if len(scenarios_to_run) > 1:
        print_comparison(all_results)

    if args.csv:
        export_csv(all_results)


if __name__ == "__main__":
    main()
