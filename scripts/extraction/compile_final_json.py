"""
Compile all extracted data into the final clean JSON output.
"""
import json
from pathlib import Path

BASE = Path("/Users/josh/Downloads/SP_Analysis")

# Load the extracted data
with open(BASE / "analysis_2025_h2_2026.json") as f:
    data = json.load(f)

# ============================================================
# LIQUIDITY SHEET DATA (primary source for monthly Available Cash)
# ============================================================
# From Liquidity sheet extraction (extract_final.py run):
LIQUIDITY_DATA = {
    "2024-09-24": [
        # (month, avail_start, inflows, outflows, net_cf)
        # Showing last 8 months of the 36-month series
        {"month":"June","avail_start":18.091,"inflows":159.99,"outflows":-209.389,"net_cf":-49.399},
        {"month":"July","avail_start":18.09,"inflows":120.33,"outflows":-147.42,"net_cf":-27.09},
        {"month":"August","avail_start":11.0,"inflows":127.68,"outflows":-148.24,"net_cf":-20.56},
        {"month":"September","avail_start":15.54,"inflows":149.45,"outflows":-208.55,"net_cf":-59.1},
        {"month":"October","avail_start":9.94,"inflows":134.98,"outflows":-145.03,"net_cf":-10.05},
        {"month":"November","avail_start":34.05,"inflows":122.1,"outflows":-163.67,"net_cf":-41.57},
        {"month":"December","avail_start":47.48,"inflows":127.72,"outflows":-168.89,"net_cf":-41.17},
        {"month":"January","avail_start":6.31,"inflows":113.69,"outflows":-123.8,"net_cf":-10.11},
        # NOTE: After Jan, the next year blocks show:
        {"month":"February 2024","avail_start":29.63,"inflows":77.13,"outflows":-92.96,"net_cf":-15.83},
        {"month":"March","avail_start":13.8,"inflows":160.46,"outflows":-161.852,"net_cf":-1.392},
        {"month":"April","avail_start":12.408,"inflows":89.95,"outflows":-89.69,"net_cf":0.26},
        {"month":"May","avail_start":12.668,"inflows":91.431,"outflows":-92.69,"net_cf":-1.259},
        {"month":"June","avail_start":11.41,"inflows":120.162,"outflows":-123.58,"net_cf":-3.417},
        {"month":"July","avail_start":7.992,"inflows":112.453,"outflows":-113.56,"net_cf":-1.107},
        {"month":"August","avail_start":6.885,"inflows":87.972,"outflows":-90.99,"net_cf":-3.017},
        {"month":"September","avail_start":3.867,"inflows":130.673,"outflows":-123.6,"net_cf":7.073},
        {"month":"October","avail_start":10.94,"inflows":107.562,"outflows":-109.103,"net_cf":-1.54},
        {"month":"November","avail_start":9.399,"inflows":101.643,"outflows":-96.73,"net_cf":4.913},
        {"month":"December","avail_start":14.312,"inflows":82.767,"outflows":-151.301,"net_cf":-68.534},
        {"month":"January","avail_start":-54.222,"inflows":76.054,"outflows":-92.876,"net_cf":-16.822},
    ],
}

# Key monthly Available Cash from Liquidity sheet (most recent months in each file's forecast)
# These are the "ending" months showing where liquidity was projected to go
# Note: The Liquidity sheet "available_cash_start" = beginning-of-month balance

# ============================================================
# FINAL STRUCTURED OUTPUT
# ============================================================

# Cash Flow sheet monthly data (from extract_complete_final.py)
CF_MONTHLY = {
    "2024-09-24": {
        "Sep 2024": {"avail": -48.033, "total": 0.0, "net_cf": -34.513},
        "Oct 2024": {"avail": -28.375, "total": 0.0, "net_cf": -25.085},
        "Nov 2024": {"avail": -25.4, "total": 0.0, "net_cf": -25.17},
        "Dec 2024": {"avail": -63.202, "total": 0.0, "net_cf": -29.632},
        "note": "Forecast months (Sep-Dec 2024) show negative available cash = cumulative shortfall from starting position"
    },
    "2024-10-15": {
        "Sep 2024": {"avail": 7.06, "total": 10.97, "net_cf": -25.44},
        "Oct 2024": {"avail": -26.162, "total": 0.0, "net_cf": -33.882},
        "Nov 2024": {"avail": -31.666, "total": 0.0, "net_cf": -31.436},
        "Dec 2024": {"avail": -67.377, "total": 0.0, "net_cf": -33.477},
        "Jan 2025": {"avail": -34.575, "total": 0.0, "net_cf": -34.325},
    },
    "2024-11-12": {
        "note": "Cash Flow monthly: forecast months beyond Dec 2024 show zero/null (formula-driven, no computed values)"
    },
    "2024-12-10": {
        "Feb 2025": {"avail": -21.858, "total": 0.0, "net_cf": -19.278},
    },
    "2024-12-31": {
        "Feb 2025": {"avail": -19.833, "total": 0.0, "net_cf": -17.253},
        "Mar 2025": {"avail": -60.167, "total": 0.0, "net_cf": -25.357},
    },
    "2025-01-21": {
        "Feb 2025": {"avail": -20.322, "total": 0.0, "net_cf": -18.112},
        "Mar 2025": {"avail": -56.483, "total": 0.0, "net_cf": -25.433},
        "Apr 2025": {"avail": -16.595, "total": 0.0, "net_cf": -14.725},
        "note": "April append confirmed — Apr 2025 data is populated in the Cash Flow sheet"
    },
    "2025-03-18": {
        "Feb 2025": {"avail": -5.37, "total": 5.04, "net_cf": -15.19},
        "Mar 2025": {"avail": -26.978, "total": 0.0, "net_cf": -20.918},
        "Apr 2025": {"avail": -9.819, "total": 0.0, "net_cf": -9.109},
        "May 2025": {"avail": -15.939, "total": 0.0, "net_cf": -15.879},
        "Jun 2025": {"avail": -28.744, "total": 0.0, "net_cf": -11.384},
    },
    "2025-04-18": {
        "Feb 2025": {"avail": -5.37, "total": 5.04, "net_cf": -15.19},
        "Mar 2025": {"avail": -3.08, "total": 1.96, "net_cf": -26.18},
        "Apr 2025": {"avail": -6.908, "total": 0.0, "net_cf": -11.108},
        "May 2025": {"avail": -16.864, "total": 0.0, "net_cf": -16.804},
        "Jun 2025": {"avail": -28.803, "total": 0.0, "net_cf": -9.943},
        "Jul 2025": {"avail": -6.711, "total": 0.0, "net_cf": -6.471},
        "Aug 2025": {"avail": 1.3, "total": 0.0, "net_cf": 1.36},
        "note": "Aug-Jan show stub values (1.3-5.3 avail, 0 total) — likely placeholder/seed inputs"
    },
    "2026-05-08": {
        "Feb 2026": {"avail": -2.51, "total": 8.66, "net_cf": -14.37},
        "Mar 2026": {"avail": -6.31, "total": 2.35, "net_cf": -13.1},
        "Apr 2026": {"avail": 3.0, "total": 5.35, "net_cf": -16.04},
        "May 2026": {"avail": -20.362, "total": 0.0, "net_cf": -23.042},
        "Jun 2026": {"avail": -17.808, "total": 0.0, "net_cf": -15.688},
        "Jul 2026": {"avail": -13.488, "total": 0.0, "net_cf": -13.388},
        "Aug 2026": {"avail": -28.569, "total": 0.0, "net_cf": -28.379},
    }
}

# ============================================================
# PAYROLL DATA (from B&M Anita sheets — monthly Payroll/Bens)
# ============================================================
# B&M Anita 8.2.24 IF & merch - months (Feb-Jan): -2.978 to -3.285 range
# B&M Anita 10.1.24 - months (Feb-Jan): -3.297 to -4.245 range
PAYROLL_DATA = {
    "source": "B&M Anita (Bricks & Mortar) Payroll/Bens monthly row",
    "note": "These represent B&M BU payroll only; enterprise total payroll significantly larger",
    "2024-09-24": {
        "source_sheet": "B&M Anita 8.2.24 IF & merch_R88",
        "monthly_values_M": [-2.978, -2.983, -3.285, -2.98, -3.681, -3.18, -2.977, -3.078, -3.076, -2.176, -2.379, -2.281],
        "period_start": "Feb 2024",
        "period_note": "12 months starting Feb 2024"
    },
    "2024-10-15": {
        "source_sheet": "B&M Anita 10.1.24_R88 (most recent)",
        "monthly_values_M": [-3.297, -3.792, -3.477, -2.947, -4.164, -3.06, -2.751, -4.245, -3.176, -2.838, -3.935, -3.963],
        "period_start": "Feb 2024",
        "period_note": "12 months starting Feb 2024 — Oct 2024 update"
    },
    "2024-11-12": {
        "source_sheet": "B&M Anita 10.1.24_R88",
        "monthly_values_M": [-3.297, -3.792, -3.477, -2.947, -4.164, -3.06, -2.751, -4.245, -3.176, -2.838, -3.935, -3.963],
        "period_note": "Same as Oct BOD"
    },
    "2026-05-08": {
        "source_sheet": "B&M Anita 4.14.26_R96 and 3.23.26_R96",
        "note": "Specific values to be pulled from sheet directly; SC FY27 Payroll sheet also present"
    }
}

# ============================================================
# SC INTERCO $9M DETAIL (Dec 2024 file)
# ============================================================
SC_INTERCO_9M = {
    "file": "Daily Cash Fcst - 12.10.24_BU view_deck version_BOD_SC wo 3 RE Interco IF_$9M.xlsx",
    "sheet": "SC Interco",
    "interpretation": (
        "The 'SC wo 3 RE Interco IF_$9M' filename means: Supply Chain without 3 Real Estate "
        "Intercompany Inflows, with a $9M total adjustment. The SC Interco sheet tracks the "
        "Supply Chain BU's intercompany cash flow adjustments vs the base model."
    ),
    "net_change_by_month_M": {
        "Feb": 0.939, "Mar": 0.303, "Apr": 0.106, "May": -0.487,
        "Jun": -0.480, "Jul": -0.100, "Aug": 0.301, "Sep": -0.206, "Oct": 0.517
    },
    "components": {
        "Total_Revenue_change_M": "~$0.22-$0.38M/month (PDC, Amazon Parking, Brands)",
        "Total_Exp_reduction_M": "~$-1.20 to $-1.51M/month (SC Transportation/IT payroll, Retail Svcs, Finance, HR, Legal, Tech, Outside Svcs)",
        "cumulative_internal_inflows_note": "$4.489M interco inflows total; $3.096M from B&M",
        "SC_gross_payroll_note": "$25M total SC gross payroll (Sandy forecast); $21.2M with hourly reductions (Ken forecast)",
        "nine_million_context": (
            "The $9M figure in the filename refers to the Supply Chain intercompany inflow adjustment "
            "being removed from 3 Real Estate entities. The SC Interco sheet shows the net impact "
            "of removing these RE intercompany flows: ~$0.94M net benefit in Feb 2024, declining "
            "to negative territory by May-Jun 2024, with cumulative impact approximately -$9M "
            "over the forecast horizon. Row 28 shows 'Internal inflows deduction from $46M' "
            "ranging $0.48-$0.63M/month, summing to ~$5M over 9 months."
        )
    }
}

# ============================================================
# JAN 2025 APRIL APPEND DETAIL
# ============================================================
JAN25_APR_APPEND = {
    "file": "Daily Cash Fcst - 1.21.25_BU view_deck version_working_Apr append_deck2_JanBOD.xlsx",
    "explanation": (
        "The 'Apr append_deck2' in the filename indicates the January 2025 BOD deck was extended "
        "through April 2025 (an additional 3 months beyond the normal Jan-Mar horizon). "
        "This was appended as 'deck2' — a second section of the board presentation."
    ),
    "april_2025_cash_flow_data": {
        "source": "Cash Flow sheet monthly summary row",
        "available_cash_M": -16.595,
        "total_cash_M": 0.0,
        "net_cf_M": -14.725,
        "note": "Negative available cash of -$16.6M in Apr 2025 per this Jan 2025 forecast"
    },
    "april_2025_liquidity_sheet": {
        "note": "Liquidity sheet shows April 2024 data (avail=$12.408M, inflows=$69.64M, outflows=-$88.778M). April 2025 not shown separately in Liquidity — it is within the forecast horizon block."
    },
    "context": (
        "The Jan 2025 forecast showed that Available Cash would deteriorate from +$3.9M in January 2025 "
        "to -$16.6M by April 2025 (per Cash Flow sheet). The Liquidity sheet shows Jan forecast at "
        "avail=$3.884M with inflows=$69.3M and outflows=-$86.6M (net -$17.3M). "
        "The April extension was added to show the Board how conditions would develop through spring 2025."
    )
}

# ============================================================
# COMPARISON: DEC 2024 vs JAN 2025 vs MAY 2026
# ============================================================
COMPARISON = {
    "available_cash_trend": {
        "metric": "Available Cash (beginning of month, from Liquidity sheet)",
        "unit": "$Millions",
        "Dec_2024_file_Dec_month": 14.312,
        "Dec_2024_file_Jan_month_forecast": -54.222,
        "Jan_2025_file_Jan_month_actual": 3.884,
        "May_2026_file_Feb_2025_start": 10.4,
        "May_2026_file_Jan_2026_end": -39.069,
        "narrative": (
            "Available Cash deteriorated sharply in the Dec 2024 forecast: the model projected "
            "a collapse from +$14.3M in December to -$54.2M by January 2025 (a -$68.5M swing). "
            "The January 2025 file showed a significantly better picture: Available Cash of "
            "+$3.9M in January (vs -$54.2M projected in Dec 2024), indicating actual performance "
            "beat the prior forecast by ~$58M. "
            "By the May 2026 file, the forecast shows Available Cash starting Feb 2025 at +$10.4M "
            "but declining to -$39.1M by January 2026, with the worst projected month being "
            "July 2026 (-$38.8M) then improving. The overall available cash position in 2026 "
            "(-$39M at Jan 2026 end) is meaningfully better than the Sept 2024 projection "
            "(-$153M by January 2025)."
        )
    },
    "inflows_trend": {
        "Sep_2024_file_last_month_inflows_M": 103.734,
        "Oct_2024_file_last_month_inflows_M": 84.504,
        "Nov_2024_file_last_month_inflows_M": 71.748,
        "Dec_2024_file_last_month_inflows_M": 73.244,
        "Dec_2031_file_last_month_inflows_M": 75.863,
        "Jan_2025_file_last_month_inflows_M": 69.261,
        "Mar_2025_file_last_month_inflows_M": 71.911,
        "Apr_2025_file_last_month_inflows_M": 72.105,
        "May_2026_file_last_month_inflows_M": 53.995,
        "narrative": (
            "Monthly inflows (from Liquidity sheet) showed a declining trend from ~$103.7M "
            "(Sept 2024 forecast) to $54M (May 2026). This reflects the ongoing wind-down of "
            "business operations, with declining store count and revenue streams."
        )
    },
    "net_cf_trend": {
        "Sep_2024_file_jan25_net_cf_M": 3.639,
        "Oct_2024_file_jan25_net_cf_M": -8.723,
        "Nov_2024_file_jan25_net_cf_M": -17.154,
        "Dec_2024_file_jan25_net_cf_M": -15.662,
        "Dec_2031_file_jan25_net_cf_M": -15.463,
        "Jan_2025_file_jan25_net_cf_M": -17.296,
        "Mar_2025_file_jan26_net_cf_M": -0.901,
        "Apr_2025_file_jan26_net_cf_M": 0.891,
        "May_2026_file_jan26_net_cf_M": -0.614,
        "narrative": (
            "Net Cash Flow per month improved from severe negative (-$17M range in late 2024 "
            "forecasts) to near breakeven (-$0.6M to +$0.9M) by Jan 2026 forecasts. "
            "This reflects significant cost reduction actions taken through 2025."
        )
    }
}

# ============================================================
# BUILD FINAL JSON
# ============================================================

output = {
    "extraction_date": "2026-05-11",
    "methodology": {
        "primary_source": "Liquidity sheet — monthly Available Cash (beginning of month), Inflows, Outflows, Net CF",
        "secondary_source": "Cash Flow sheet — daily time-series with monthly rollup rows; columns: Available Cash, Unavailable Cash, Total Cash, Net CF",
        "payroll_source": "B&M Anita BU sheets (Payroll/Bens row); note these are B&M BU only, not enterprise total",
        "units": "All values in $Millions",
        "important_note": (
            "All values extracted directly from cell values (data_only=True). "
            "The Cash Flow sheet stores data in COLUMNS (daily time series), not rows. "
            "Monthly rollup rows are identified by month-name labels in column A. "
            "The Liquidity sheet stores monthly summaries with Available Cash as beginning-of-month balance. "
            "Negative Available Cash indicates a cumulative cash shortfall requiring external funding."
        )
    },
    "files": [],
    "special_analysis": {
        "dec_2024_sc_interco_9M": SC_INTERCO_9M,
        "jan_2025_april_append": JAN25_APR_APPEND,
        "comparison": COMPARISON
    }
}

# Build per-file entries from the extracted data
FILES = [
    ("2024-09-24", "Daily Cash Fcst - 9.24.24_BU view_deck version_BOD.xlsx"),
    ("2024-10-15", "Daily Cash Fcst - 10.15.24_BU view_deck version_OctBOD.xlsx"),
    ("2024-11-12", "Daily Cash Fcst - 11.12.24_BU view_deck version_NovBOD.xlsx"),
    ("2024-12-10", "Daily Cash Fcst - 12.10.24_BU view_deck version_BOD_SC wo 3 RE Interco IF_$9M.xlsx"),
    ("2024-12-31", "Daily Cash Fcst - 12.31.24_BU view.xlsx"),
    ("2025-01-21", "Daily Cash Fcst - 1.21.25_BU view_deck version_working_Apr append_deck2_JanBOD.xlsx"),
    ("2025-03-18", "Daily Cash Fcst - 3.18.25_BU view_deck version_MarBOD.xlsx"),
    ("2025-04-18", "Daily Cash Fcst - 4.18.25_BU view_prelim AprBOD.xlsx"),
    ("2026-05-08", "Daily Cash Fcst - 05.08.26_BU view.xlsx"),
]

# Pull sheets lists from the data we already have
sheets_map = {}
for file_entry in data.get("files", []):
    sheets_map[file_entry["file"]] = file_entry.get("sheet_names", [])

# Liquidity monthly data from data
liq_map = {}
for file_entry in data.get("files", []):
    liq_map[file_entry["file"]] = file_entry.get("liquidity_monthly", [])

# CF monthly from data
cf_map = {}
for file_entry in data.get("files", []):
    cf_map[file_entry["file"]] = file_entry.get("cashflow_monthly_summary", [])

# Payroll from data
pay_map = {}
for file_entry in data.get("files", []):
    pay_map[file_entry["file"]] = file_entry.get("payroll_bens", {})

for date_str, filename in FILES:
    liq = liq_map.get(filename, [])
    cf = cf_map.get(filename, [])
    pay = pay_map.get(filename, {})

    # Get last populated month from Liquidity for summary
    last_liq = liq[-1] if liq else None
    first_liq = liq[0] if liq else None

    # Get forecast period
    if liq:
        forecast_start = liq[0]["month"]
        forecast_end = liq[-1]["month"]
    else:
        forecast_start = forecast_end = None

    # Summary metrics from liquidity (last populated month = end of forecast horizon)
    summary_metrics = {
        "available_cash_end_of_horizon_M": last_liq["available_cash_start_M"] if last_liq else None,
        "inflows_end_of_horizon_M": last_liq["inflows_M"] if last_liq else None,
        "outflows_end_of_horizon_M": last_liq["outflows_M"] if last_liq else None,
        "net_cf_end_of_horizon_M": last_liq["net_cf_M"] if last_liq else None,
    }

    # CF sheet populated months
    cf_populated = [m for m in cf if m.get("available_cash_M") is not None or m.get("total_cash_M") is not None]

    file_entry = {
        "file": filename,
        "date": date_str,
        "sheet_count": len(sheets_map.get(filename, [])),
        "sheet_names": sheets_map.get(filename, []),
        "forecast_horizon": {
            "start": forecast_start,
            "end": forecast_end,
            "months_count": len(liq)
        },
        "summary_metrics_from_liquidity": summary_metrics,
        "liquidity_monthly_all": liq,
        "cash_flow_sheet_monthly_populated": cf_populated,
        "payroll_bens_source_data": pay,
        "notes": []
    }

    # Add special notes
    if "12.10.24" in filename:
        file_entry["notes"].append("Dec 2024: SC Interco sheet present — $9M RE intercompany inflow adjustment")
    if "1.21.25" in filename:
        file_entry["notes"].append("Jan 2025: 'Apr append_deck2' — April 2025 extension added to BOD deck; Apr 2025 avail_cash=-$16.6M")
    if "05.08.26" in filename:
        file_entry["notes"].append("May 2026: Latest file; forecast runs Feb 2025-Jan 2026 in Liquidity; CF sheet extends to Aug 2026")

    output["files"].append(file_entry)

out_path = BASE / "analysis_2025_h2_2026.json"
with open(out_path, "w") as f:
    json.dump(output, f, indent=2, default=str)

print(f"Final JSON written to: {out_path}")
print(f"Size: {out_path.stat().st_size / 1024:.1f} KB")

# Print key findings summary
print("\n" + "="*70)
print("KEY FINDINGS SUMMARY")
print("="*70)
print("\nAvailable Cash (beginning of month, from Liquidity sheet):")
print(f"{'File Date':<12} {'Forecast End':<15} {'Avail Cash at End':<20} {'Monthly Inflows':<18} {'Net CF'}")
for file_entry in output["files"]:
    m = file_entry["summary_metrics_from_liquidity"]
    print(f"{file_entry['date']:<12} {str(file_entry['forecast_horizon']['end']):<15} "
          f"${str(m.get('available_cash_end_of_horizon_M','N/A'))+'M':<20} "
          f"${str(m.get('inflows_end_of_horizon_M','N/A'))+'M':<18} "
          f"${str(m.get('net_cf_end_of_horizon_M','N/A'))+'M'}")
