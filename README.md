# Transform SR — Daily Cash Forecast Analysis Suite

**Prepared for:** Shane (via Josh Stephens)  
**Date:** May 11, 2026  
**Company:** Transform SR Holding Management LLC (Post-Sears Holdings, ESL Investments)

This repository contains everything produced by a 28-agent AI analysis of Transform's complete Daily Cash Forecast SharePoint repository — 50 BOD/deck files, 619 MB of source data, covering September 2022 through May 8, 2026.

---

## 🗂️ Repository Structure

```
For-Shane/
├── README.md                          ← You are here
├── MASTER_INSIGHTS_DOCUMENT.md        ← Full 3.5-year company analysis (40KB, 700 lines)
├── CFO_RECOMMENDATIONS.md             ← Business-by-business strategic recommendations
│
├── analysis/                          ← 27 structured JSON files from agent extraction
│   ├── analysis_liquidity.json        ← Complete liquidity/available cash timeline
│   ├── analysis_payroll_headcount.json← Full monthly payroll actuals + BU breakdown
│   ├── analysis_disbursements_2024_2026.json ← Disbursements by category, 29 months
│   ├── analysis_inflows_trend.json    ← 45-month inflows timeline, all categories
│   ├── analysis_trapped_cash_evolution.json  ← LC collateral & trapped cash history
│   ├── analysis_home_warranty.json    ← Assurant/HelloSuper/CCHS deep dive
│   ├── analysis_syw_online_rx.json    ← SYW credit card, Online decline, Rx rebates
│   ├── analysis_post_reductions.json  ← Feb 2026 workforce restructuring analysis
│   ├── analysis_automation_architecture.json ← Full automation pipeline design
│   ├── analysis_HS_KCD.json           ← NOTE: 1.8MB file not in repo (too large)
│   │                                    → run scripts/extraction/extract_hs_kcd.py
│   └── [20 more analysis files...]
│
├── scripts/
│   ├── sharepoint/                    ← Accessing Keith Kim's SharePoint drive
│   │   ├── batch_download.py          ← Batch download via authenticated browser session
│   │   ├── download_sp_files.py       ← Playwright-based download (requires auth)
│   │   ├── fetch_listing.py           ← SharePoint folder listing via REST API
│   │   └── ms_graph_poll.py           ← MS Graph API polling for file metadata
│   │
│   ├── extraction/                    ← Reading and extracting data from .xlsx files
│   │   ├── extract_payroll.py         ← Full payroll history from Payroll tab
│   │   ├── extract_disbursements.py   ← Disbursements from BU beta (Jan 2024–May 2026)
│   │   ├── extract_inflows_master.py  ← Complete inflows from Inflows Actuals
│   │   ├── extract_hs_kcd.py          ← Home Services & KCD longitudinal extraction
│   │   ├── extract_retail_data.py     ← Sears/Kmart store-level inflow extraction
│   │   ├── extract_cash_metrics.py    ← Available cash, total cash, unavailable cash
│   │   ├── scan_bu_beta.py            ← BU beta sheet structure scanner
│   │   └── [15 more extraction scripts]
│   │
│   ├── build/                         ← Building/generating Excel and Word files
│   │   ├── build_netsuite_file.py     ← Build Daily Cash Fcst from NetSuite MCP data
│   │   ├── build_10tab.py             ← Rebuild 10-visible-tab workbook structure
│   │   ├── build_3statement.py        ← Three-statement model builder
│   │   ├── build_word_doc.py          ← Convert MASTER_INSIGHTS_DOCUMENT.md → .docx
│   │   ├── build_may8_fcst.py         ← May 8, 2026 forecast file builder
│   │   ├── forecast.py / forecast2-5.py ← Iterative forecast builders
│   │   └── [8 more build scripts]
│   │
│   └── analysis/                      ← CFO reports and variance analysis
│       ├── cfo_report_ns.py           ← CFO Variance Report (NS vs team file), .docx
│       ├── cfo_report.py              ← CFO report generator v1
│       ├── compare_cash_flow.py       ← Two-file cash flow comparison
│       └── generate_arch_doc.py       ← Generate automation architecture document
│
└── data/
    ├── files_js_array.txt             ← Complete SharePoint file manifest (51 files + IDs)
    ├── sp_item_ids.json               ← All 695 SharePoint items with driveId + itemId
    └── download_targets.json          ← 51 strategic BOD files with full metadata
```

---

## 📋 Key Findings Summary

### The Company in 5 Numbers
| Metric | Sep 2022 | May 2026 | Change |
|--------|---------|---------|--------|
| Available cash | $9.9M | **$1.5M** | -85% (peak $72.6M Jan'23) |
| Monthly payroll | $48.4M | $19.5M | **-60%** |
| Home Services net (annual) | $629M | $291M | **-54%** |
| Total monthly inflows | $149M | $28M | **-81%** |
| Sears store inflows | $18.2M/mo | $0.5M/mo | **-97%** |

### Business Unit Status
| BU | Status | Key Metric |
|----|--------|-----------|
| Home Services | Declining run-off | $629M→$291M net/yr |
| KCD (Kenmore/Craftsman/DieHard) | **Forecast ramp-up** | $41.6M FY2026 (was $5M) |
| Kmart Stores | Winding down | $83M FY2026 forecast |
| Sears Stores | Essentially done | $15M FY2026 forecast |
| SYW Loyalty | Growing | $47M+ w/ new credit card |
| Supply Chain | Declining support function | -$2.3M payroll post-reductions |
| Online/E-commerce | Declining | $11M→$7M/yr |
| Home Warranty | Transitioning | Assurant→HelloSuper |

### The 8 Liquidity Crises
Every single one was resolved by a last-minute capital event. Crisis #8 (Dec 2025–present, sub-$2M available cash) is **unresolved as of May 8, 2026**.

Largest rescue: **October 2023 — $173.19M term loan injection** (available cash jumped $137M in one day).

Total capital raised since 2022: **$627M** ($370M equity in 2022; $488M debt in 2024–2025).

### Cost Structure Transformation
Total monthly disbursements fell from **$105.6M** (Q4 2024) to **$72.3M** (Q1 2026) — **$33.3M/month saved = ~$400M annualized**.

---

## 🔧 How to Use This on Claude

### Talking to the analysis files

Load one or more JSON files and ask questions:
```
Read /path/to/analysis_liquidity.json and tell me when available cash 
fell below $5M for the first time and what caused it.
```

```
Compare analysis_payroll_headcount.json with analysis_disbursements_2024_2026.json 
and identify which quarters had the steepest combined payroll + logistics cuts.
```

### Running the extraction scripts

All extraction scripts require:
- Python 3.9+
- `openpyxl` (`pip install openpyxl`)
- `python-docx` (`pip install python-docx`) for Word output
- Source `.xlsx` files in `/Users/josh/Downloads/SP_Analysis/` (or update `BASE_DIR`)

```bash
# Example: regenerate payroll analysis
python3 scripts/extraction/extract_payroll.py

# Example: rebuild the Word document from the markdown
python3 scripts/build/build_word_doc.py
```

### Regenerating the Daily Cash Forecast from NetSuite

`scripts/build/build_netsuite_file.py` uses the NetSuite MCP to pull live data and build a complete Daily Cash Forecast `.xlsx`. Requires:
- NetSuite MCP configured in Claude Code (`mcp__f5e53216...` tools)
- Transform SR Holding Management LLC subsidiary IDs loaded

```
Key caveat: NetSuite only covers ~20% of actual cash flows by value.
The remaining 80% requires: ADP SFTP (payroll), JPMorgan BAI2 (bank),
MERCH feed (merchandise), and manual HS projections.
```

---

## 📊 SharePoint Access

**Drive:** Keith Kim's SharePoint at searshc-my.sharepoint.com  
**Drive ID:** `b!5YH6v6JhxUClPPX4EPRoLOvOPIl2jWtDquaa19bpJwPVFivQQJUyQKC4fgUz3Xf-`  
**Path:** `/personal/keith_kim_transformco_com/Documents/Cash forecast - BU/Daily Excel Files`  
**Total files in folder:** 695  
**Strategic files downloaded:** 51 (BOD + deck versions, one per month)

### To access SharePoint files programmatically:
1. Authenticate as `Joshua.stephens@transformco.com` via browser (Playwright MCP or headless)
2. Use the v2.0 API with driveId + itemId from `data/sp_item_ids.json`:
   ```
   GET https://searshc-my.sharepoint.com/_api/v2.0/drives/{driveId}/items/{itemId}/content
   ```
3. Or use the browser fetch API pattern in `scripts/sharepoint/batch_download.py`

---

## 📁 The Source Files (not in this repo — too large)

The 51 downloaded `.xlsx` files (~619 MB total) live at:
```
/Users/josh/Downloads/SP_Analysis/*.xlsx
```

They follow this naming convention:
```
Daily Cash Fcst - [M.DD.YY]_BU view[_label]_[BOD name].xlsx
```

File structure (10 visible tabs + 8 hidden):
| Tab | Contents |
|-----|---------|
| Cash Flow | Daily cash flow: inflows, disbursements, available cash |
| Inflows Detail | Daily inflows by BU (all history + forecast) |
| Inflows Actuals | Daily actual inflows (the authoritative source) |
| Inflows Forecasting | Weekly forward projections by BU |
| Disbursement Detail | Daily disbursements by category |
| Trapped Cash | Unavailable/restricted cash breakdown |
| Liquidity | Monthly summary + financing sources |
| BU beta | Monthly BU-level P&L (505 rows × 2,695 cols in May 2026) |
| RE Fcst Summary | Real estate revenue forecast |
| [Hidden tabs] | Payroll, MERCH, SYW_CF, SC IF, KCD backup, SHS Cash Changes, etc. |

---

## 🏗️ Automation Architecture (Design Spec)

The full automation pipeline design is in `analysis/analysis_automation_architecture.json`.

**Summary:** A 4-layer pipeline to produce the Daily Cash Forecast automatically in <5 minutes:

1. **Collection** — NetSuite MCP + ADP SFTP + JPMorgan BAI2 + MERCH feed
2. **Transformation** — Variance engine, rolling 4-week projector, debt/rent schedule builder
3. **Generation** — openpyxl Excel builder (preserves formulas, updates dynamic cells)
4. **QA** — 7 automated quality gates (date continuity, BU totals, cash bridge, CF tie-out, etc.)

Current manual effort: ~4 hours/day → Target: 30 minutes/day (exception review only).

---

## 💬 Context for Shane's Claude Instance

This analysis was produced by Josh Stephens (jstephens12083@gmail.com) using Claude Code with 28 parallel analysis agents deployed across the full SharePoint repository of Transform SR's Daily Cash Forecast files.

**The company context:**
- Transform SR = post-Sears Holdings bankruptcy entity (ESL Investments, Ed Lampert)
- "Midco" in internal documents = Transform SR Holding Management LLC
- "TopCo" = Transform Holdco LLC (parent, controls equity raises)
- Keith Kim = Treasury analyst who maintains the Daily Cash Forecast
- Josh Stephens = Finance team member analyzing the data

**Model file structure tip for Shane:**
The BU beta sheet is the master P&L engine. The label column moves over time:
- 2022–2023 files: label column = column 21 (U)
- 2025–2026 files: label column = column 61 (BI)

Always use `data_only=True, read_only=True` in openpyxl to get computed values (not formulas) without loading all the formula calculation engine — these files are 7–15 MB each and have 233K+ formula cells.

---

*Generated May 11, 2026. All figures sourced directly from spreadsheet cell values.*
