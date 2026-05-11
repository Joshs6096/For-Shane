"""
Build the longitudinal retail table directly from Historicals/BU beta data.
This is faster because we know exactly where the data lives.
"""
import os, json, openpyxl
from datetime import datetime

BASE_DIR = "/Users/josh/Downloads/SP_Analysis"

BOD_FILES = [
    ("9.19.22_Sep BOD",   "9.19.22_BU view_Sep BOD"),
    ("12.7.22_DecBOD",    "12.7.22_BU view_DecBOD"),
    ("1.17.23_JanBOD",    "1.17.23_BU view_JanBOD"),
    ("4.25.23_AprBOD",    "4.25.23_BU view_April 27 BOD"),
    ("7.25.23_JulBOD",    "7.25.23_BU view_JulBOD"),
    ("11.14.23_NovBOD",   "11.14.23_BU view_adUPDATE"),
    ("2.20.24_BOD",       "2.20.24_BU view_BOD"),
    ("6.18.24_JunBOD",    "6.18.24_BU view_JunBOD"),
    ("10.15.24_OctBOD",   "10.15.24_BU view_deck version_OctBOD"),
    ("12.31.24_BU view",  "12.31.24_BU view"),
    ("3.18.25_MarBOD",    "3.18.25_BU view_deck version_MarBOD"),
    ("7.22.25_JulBOD",    "7.22.25_BU view_deck version_condensed_JulBOD"),
    ("12.15.25_DecBOD",   "12.15.25_BU view_DecBOD"),
    ("01.16.26_JanBOD",   "01.16.26_BU view new format_asset sales_JanBOD"),
    ("04.21.26_AprBOD",   "04.21.26_ BU version_deck version_AprBOD"),
    ("05.08.26_BU view",  "05.08.26_BU view"),
]

def find_file(partial_name):
    for fn in os.listdir(BASE_DIR):
        if partial_name.lower() in fn.lower() and fn.endswith(".xlsx"):
            return os.path.join(BASE_DIR, fn)
    return None

def safe(v):
    if isinstance(v, (int, float)):
        return round(float(v), 4)
    return None

MONTH_ABBREVS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12
}

def parse_month_col_header(v):
    """Parse strings like 'Feb 2022', 'Mar 2023', 'Jan 2026' into YYYY-MM."""
    if v is None: return None
    s = str(v).strip()
    parts = s.split()
    if len(parts) == 2:
        mon = parts[0][:3].lower()
        year_s = parts[1]
        if mon in MONTH_ABBREVS:
            try:
                year = int(year_s)
                return f"{year}-{MONTH_ABBREVS[mon]:02d}"
            except: pass
    return None


def extract_historicals_view(ws):
    """
    Historicals view KK: Row 3 = month headers (Feb 2022, Mar 2022, ...),
    Row 4 = Sears Stores, Row 5 = Kmart Stores
    Returns: {YYYY-MM: {sears: x, kmart: x}}
    """
    # Find header row with month labels
    month_header_row = None
    sears_data_row = None
    kmart_data_row = None

    for r in range(1, 20):
        for c in range(1, 5):
            v = ws.cell(row=r, column=c).value
            if v is None: continue
            vn = str(v).lower().strip()
            if "sears store" in vn:
                sears_data_row = r
            if "kmart store" in vn:
                kmart_data_row = r
        # Check for month labels in this row
        month_count = 0
        for c in range(2, 40):
            v = ws.cell(row=r, column=c).value
            if v and isinstance(v, str):
                if parse_month_col_header(v):
                    month_count += 1
        if month_count >= 3:
            month_header_row = r

    if not (sears_data_row and kmart_data_row and month_header_row):
        return None

    # Map column -> YYYY-MM
    col_to_month = {}
    for c in range(2, min((ws.max_column or 50)+1, 100)):
        v = ws.cell(row=month_header_row, column=c).value
        ym = parse_month_col_header(v)
        if ym:
            col_to_month[c] = ym

    data = {}
    for c, ym in col_to_month.items():
        s = safe(ws.cell(row=sears_data_row, column=c).value)
        k = safe(ws.cell(row=kmart_data_row, column=c).value)
        if s is not None or k is not None:
            data[ym] = {"sears_M": s, "kmart_M": k}

    return data


def extract_bu_beta_monthly(ws):
    """
    BU beta: Row 3 = month labels (Feb 2022, ...) but offset right
    Sears Stores and Kmart Stores appear in rows 6 or 7
    Returns: {YYYY-MM: {sears: x, kmart: x}}
    """
    month_header_row = None
    sears_data_row = None
    kmart_data_row = None
    sears_label_col = None

    for r in range(1, 30):
        for c in range(1, 30):
            v = ws.cell(row=r, column=c).value
            if v is None: continue
            vn = str(v).lower().strip()
            if "sears store" in vn and sears_data_row is None:
                sears_data_row = r
                sears_label_col = c
            if "kmart store" in vn and kmart_data_row is None:
                kmart_data_row = r

    if not (sears_data_row and kmart_data_row):
        return None

    # Find month header row: scan rows before sears_data_row for month labels
    for r in range(1, sears_data_row):
        month_count = 0
        for c in range(sears_label_col, min((ws.max_column or 100)+1, 200)):
            v = ws.cell(row=r, column=c).value
            if v and isinstance(v, str) and parse_month_col_header(v):
                month_count += 1
        if month_count >= 3:
            month_header_row = r
            break

    if not month_header_row:
        return None

    # Map column -> YYYY-MM
    col_to_month = {}
    for c in range(sears_label_col, min((ws.max_column or 100)+1, 200)):
        v = ws.cell(row=month_header_row, column=c).value
        ym = parse_month_col_header(v)
        if ym:
            col_to_month[c] = ym

    data = {}
    for c, ym in col_to_month.items():
        s = safe(ws.cell(row=sears_data_row, column=c).value)
        k = safe(ws.cell(row=kmart_data_row, column=c).value)
        if s is not None or k is not None:
            data[ym] = {"sears_M": s, "kmart_M": k}

    return data


def extract_bu_beta_annual(ws):
    """Extract annual (FY) totals from BU beta sheet."""
    sears_data_row = None
    kmart_data_row = None
    sears_label_col = None

    for r in range(1, 30):
        for c in range(1, 30):
            v = ws.cell(row=r, column=c).value
            if v is None: continue
            vn = str(v).lower().strip()
            if "sears store" in vn and sears_data_row is None:
                sears_data_row = r
                sears_label_col = c
            if "kmart store" in vn and kmart_data_row is None:
                kmart_data_row = r

    if not (sears_data_row and kmart_data_row):
        return None

    # Find FY total columns: look for FY22, FY23, FY24, FY25, FY26 labels
    fy_cols = {}
    for r in range(1, sears_data_row):
        for c in range(1, min((ws.max_column or 100)+1, 100)):
            v = ws.cell(row=r, column=c).value
            if v and isinstance(v, str):
                vn = v.strip().upper()
                if vn.startswith("FY") and len(vn) == 4:
                    try:
                        yr = int(vn[2:])
                        fy_cols[c] = f"FY{yr}"
                    except: pass

    annual = {}
    for c, fy_label in fy_cols.items():
        s = safe(ws.cell(row=sears_data_row, column=c).value)
        k = safe(ws.cell(row=kmart_data_row, column=c).value)
        if s is not None or k is not None:
            annual[fy_label] = {"sears_M": s, "kmart_M": k}

    return annual


def extract_store_counts_from_sheet(ws, sheet_name):
    """Scan for store count rows."""
    STORE_KW = ["store count", "# stores", "# of stores", "number of stores",
                "open stores", "operating stores", "b&m stores", "sears stores #",
                "kmart stores #"]
    results = []
    for r in range(1, min((ws.max_row or 300)+1, 400)):
        for c in range(1, 8):
            v = ws.cell(row=r, column=c).value
            if v is None: continue
            vn = str(v).lower().strip()
            if any(kw in vn for kw in STORE_KW):
                vals = {}
                for cc in range(c, min((ws.max_column or 100)+1, 100)):
                    vv = ws.cell(row=r, column=cc).value
                    if isinstance(vv, (int, float)) and vv > 0:
                        vals[cc] = vv
                results.append({"sheet": sheet_name, "row": r, "label": str(v), "values": vals})
    return results


def process_file(label, partial_name):
    filepath = find_file(partial_name)
    if not filepath:
        print(f"  NOT FOUND: {partial_name}", flush=True)
        return {"label": label, "error": "not found"}

    print(f"\nProcessing {label}: {os.path.basename(filepath)}", flush=True)

    result = {
        "label": label,
        "filename": os.path.basename(filepath),
        "monthly_data": {},          # YYYY-MM -> {sears_M, kmart_M}
        "annual_data": {},           # FY22 -> {sears_M, kmart_M}
        "store_counts": [],
        "territory": {},
        "source_sheets": [],
        "errors": [],
    }

    try:
        wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
        sheets = wb.sheetnames

        for sn in sheets:
            sn_lower = sn.lower()
            try:
                ws = wb[sn]

                # Historicals view
                if "historica" in sn_lower:
                    data = extract_historicals_view(ws)
                    if data:
                        result["monthly_data"].update(data)
                        result["source_sheets"].append(f"Historicals({len(data)} months)")
                        print(f"  Historicals: {len(data)} months", flush=True)

                # BU beta
                elif "bu beta" in sn_lower:
                    data = extract_bu_beta_monthly(ws)
                    if data:
                        # Only add months not already in result
                        for ym, d in data.items():
                            if ym not in result["monthly_data"]:
                                result["monthly_data"][ym] = d
                        result["source_sheets"].append(f"BU beta({len(data)} months)")
                        print(f"  BU beta monthly: {len(data)} months", flush=True)

                    annual = extract_bu_beta_annual(ws)
                    if annual:
                        result["annual_data"].update(annual)
                        print(f"  BU beta annual: {list(annual.keys())}", flush=True)

                # Inflows Actuals (daily) - use for recent months not in historical
                elif "inflows actuals" in sn_lower or "inflow actuals" in sn_lower:
                    # Extract daily data
                    date_col = 1
                    sears_col = kmart_col = guam_col = usvi_col = None
                    for r in range(1, 5):
                        for c in range(1, 60):
                            v = ws.cell(row=r, column=c).value
                            if v is None: continue
                            vn = str(v).lower().strip()
                            if "sears store" in vn and sears_col is None:
                                sears_col = c
                                header_row = r
                            if "kmart store" in vn and kmart_col is None:
                                kmart_col = c
                            if "guam" in vn and guam_col is None:
                                guam_col = c
                            if "usvi" in vn and usvi_col is None:
                                usvi_col = c
                    if sears_col and kmart_col:
                        from collections import defaultdict
                        monthly = defaultdict(lambda: {"sears_M": 0.0, "kmart_M": 0.0, "guam_M": 0.0, "usvi_M": 0.0})
                        for r in range((header_row or 1)+1, min((ws.max_row or 500)+1, 2500)):
                            dv = ws.cell(row=r, column=1).value
                            if isinstance(dv, datetime):
                                ym = dv.strftime("%Y-%m")
                            elif hasattr(dv, "year"):
                                ym = f"{dv.year}-{dv.month:02d}"
                            else:
                                continue
                            sv = ws.cell(row=r, column=sears_col).value
                            kv = ws.cell(row=r, column=kmart_col).value
                            if isinstance(sv, (int, float)):
                                monthly[ym]["sears_M"] += sv
                            if isinstance(kv, (int, float)):
                                monthly[ym]["kmart_M"] += kv
                            if guam_col:
                                gv = ws.cell(row=r, column=guam_col).value
                                if isinstance(gv, (int, float)):
                                    monthly[ym]["guam_M"] += gv
                            if usvi_col:
                                uv = ws.cell(row=r, column=usvi_col).value
                                if isinstance(uv, (int, float)):
                                    monthly[ym]["usvi_M"] += uv
                        for ym, d in monthly.items():
                            if ym not in result["monthly_data"]:
                                result["monthly_data"][ym] = {
                                    "sears_M": round(d["sears_M"], 4),
                                    "kmart_M": round(d["kmart_M"], 4)
                                }
                            if d.get("guam_M") or d.get("usvi_M"):
                                result["territory"][ym] = {
                                    "guam_M": round(d["guam_M"], 4),
                                    "usvi_M": round(d["usvi_M"], 4)
                                }
                        result["source_sheets"].append(f"Inflows Actuals({len(monthly)} months)")
                        print(f"  Inflows Actuals: {len(monthly)} months", flush=True)

                # Store count scan on all sheets
                sc = extract_store_counts_from_sheet(ws, sn)
                if sc:
                    result["store_counts"].extend(sc)
                    print(f"  Store counts in '{sn}': {[x['label'] for x in sc]}", flush=True)

            except Exception as e:
                result["errors"].append(f"{sn}: {e}")

        wb.close()

    except Exception as e:
        result["errors"].append(f"File: {e}")
        print(f"  ERROR: {e}", flush=True)

    # Sort monthly data
    result["monthly_data"] = dict(sorted(result["monthly_data"].items()))
    return result


def main():
    results = []
    for label, partial in BOD_FILES:
        r = process_file(label, partial)
        results.append(r)

    # Build longitudinal summary
    # Collect all months across all files
    all_months = set()
    for r in results:
        all_months.update(r.get("monthly_data", {}).keys())
    all_months = sorted(all_months)

    longitudinal = []
    for r in results:
        md = r.get("monthly_data", {})
        ad = r.get("annual_data", {})
        row = {
            "bod_label": r["label"],
            "filename": r.get("filename", ""),
            "source_sheets": r.get("source_sheets", []),
            "monthly_inflows": md,
            "annual_inflows": ad,
            "store_counts": r.get("store_counts", []),
            "territory": r.get("territory", {}),
            "months_with_data": len(md),
            "errors": r.get("errors", [])[:3],
        }
        longitudinal.append(row)

    # Build a clean summary table for each BOD showing latest known monthly actuals
    print("\n\n" + "="*100, flush=True)
    print("LONGITUDINAL SUMMARY: Sears & Kmart Monthly Inflows ($M)", flush=True)
    print("="*100, flush=True)
    print(f"{'BOD':28s} | {'Source':22s} | {'Sears Last Month':18s} | {'Kmart Last Month':18s} | {'#Months':7s}", flush=True)
    print("-"*100, flush=True)

    for row in longitudinal:
        md = row["monthly_inflows"]
        if md:
            last_month = sorted(md.keys())[-1]
            last_sears = md[last_month].get("sears_M", "n/a")
            last_kmart = md[last_month].get("kmart_M", "n/a")
            src = ", ".join(row["source_sheets"])[:22]
            print(f"{row['bod_label']:28s} | {src:22s} | {str(last_sears):>10s} ({last_month}) | {str(last_kmart):>10s} ({last_month}) | {row['months_with_data']:7d}", flush=True)
        else:
            print(f"{row['bod_label']:28s} | no data", flush=True)

    # Print annual totals
    print("\n\nANNUAL TOTALS ($M) from BU beta FY columns:", flush=True)
    for row in longitudinal:
        ad = row["annual_inflows"]
        if ad:
            annual_str = ", ".join([f"{fy}: S={d['sears_M']} K={d['kmart_M']}" for fy, d in sorted(ad.items())])
            print(f"  {row['bod_label']:28s}: {annual_str[:120]}", flush=True)

    # Save
    output = {
        "generated": "2026-05-11",
        "description": "Sears and Kmart retail inflow longitudinal analysis across 16 BOD files",
        "units": "Millions USD unless noted",
        "methodology": {
            "primary_source": "Historicals view KK sheet (monthly aggregates), BU beta sheet",
            "secondary_source": "Inflows Actuals sheet (daily data aggregated to monthly)",
            "note": "Sears Stores col2 and Kmart Stores col3 in Inflows Actuals. Data_only=True used."
        },
        "longitudinal_table": longitudinal,
        "all_months_in_dataset": all_months,
    }

    out_path = os.path.join(BASE_DIR, "analysis_retail_stores.json")
    with open(out_path, "w") as f:
        json.dump(output, f, indent=2, default=str)
    print(f"\nSaved: {out_path}", flush=True)


if __name__ == "__main__":
    main()
