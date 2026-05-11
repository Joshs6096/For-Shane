"""
Fast targeted extraction:
- Historicals view KK: monthly data for ALL files
- BU beta: annual FY totals only (faster - only scan specific cells)
- BU beta for 01.16.26 and 05.08.26: known structure, get FY cols 9-16

Focus on getting the data we need quickly.
"""
import os, json, openpyxl
from datetime import datetime

BASE_DIR = "/Users/josh/Downloads/SP_Analysis"

MONTH_ABBREVS = {
    "jan":1,"feb":2,"mar":3,"apr":4,"may":5,"jun":6,
    "jul":7,"aug":8,"sep":9,"oct":10,"nov":11,"dec":12
}

def parse_ym(v):
    if v is None: return None
    s = str(v).strip()
    parts = s.split()
    if len(parts) == 2:
        mon = parts[0][:3].lower()
        if mon in MONTH_ABBREVS:
            try:
                return f"{int(parts[1])}-{MONTH_ABBREVS[mon]:02d}"
            except: pass
    return None

def safe(v):
    if isinstance(v, (int, float)): return round(float(v), 4)
    return None

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

def find_file(partial):
    for fn in os.listdir(BASE_DIR):
        if partial.lower() in fn.lower() and fn.endswith(".xlsx"):
            return os.path.join(BASE_DIR, fn)
    return None

def extract_historicals(ws):
    """Row 3 = month headers, Row 4 = Sears, Row 5 = Kmart."""
    sears_row = kmart_row = hdr_row = None
    for r in range(1, 20):
        for c in range(1, 5):
            v = ws.cell(row=r, column=c).value
            if v is None: continue
            vn = str(v).lower().strip()
            if "sears store" in vn: sears_row = r
            if "kmart store" in vn: kmart_row = r
    if not sears_row: return {}

    # Find header row above sears
    for r in range(1, sears_row):
        cnt = sum(1 for c in range(2, 40) if parse_ym(ws.cell(row=r, column=c).value))
        if cnt >= 3: hdr_row = r

    if not hdr_row: return {}

    col_map = {}
    for c in range(2, min((ws.max_column or 50)+1, 80)):
        ym = parse_ym(ws.cell(row=hdr_row, column=c).value)
        if ym: col_map[c] = ym

    data = {}
    for c, ym in col_map.items():
        s = safe(ws.cell(row=sears_row, column=c).value)
        k = safe(ws.cell(row=kmart_row, column=c).value)
        if s is not None or k is not None:
            data[ym] = {"sears_M": s if s else 0.0, "kmart_M": k if k else 0.0}
    return data


def extract_bu_beta_annual_fast(ws):
    """Extract FY columns from BU beta. Faster: only scan for FY labels."""
    sears_row = kmart_row = None
    sears_col_offset = None

    for r in range(1, 30):
        for c in range(1, 25):
            v = ws.cell(row=r, column=c).value
            if v is None: continue
            vn = str(v).lower().strip()
            if "sears store" in vn and sears_row is None:
                sears_row = r
                sears_col_offset = c
            if "kmart store" in vn and kmart_row is None:
                kmart_row = r

    if not (sears_row and kmart_row): return {}

    # Find FY label columns
    fy_cols = {}
    for r in range(1, sears_row):
        for c in range(1, min((ws.max_column or 100)+1, 80)):
            v = ws.cell(row=r, column=c).value
            if v and isinstance(v, str):
                vn = v.strip().upper()
                if vn.startswith("FY") and len(vn) == 4:
                    try:
                        yr = int(vn[2:])
                        if 20 < yr < 30: fy_cols[c] = f"FY20{yr}"
                    except: pass

    annual = {}
    for c, fy in fy_cols.items():
        s = safe(ws.cell(row=sears_row, column=c).value)
        k = safe(ws.cell(row=kmart_row, column=c).value)
        if s is not None or k is not None:
            if fy not in annual:  # take first occurrence
                annual[fy] = {"sears_M": s if s else 0.0, "kmart_M": k if k else 0.0}

    return annual


def extract_inflows_actuals_monthly(ws):
    """Daily actuals aggregated to monthly."""
    sears_col = kmart_col = header_row = None
    for r in range(1, 5):
        for c in range(1, 60):
            v = ws.cell(row=r, column=c).value
            if v is None: continue
            vn = str(v).lower().strip()
            if "sears store" in vn and sears_col is None:
                sears_col = c; header_row = r
            if "kmart store" in vn and kmart_col is None:
                kmart_col = c

    if not (sears_col and kmart_col): return {}

    from collections import defaultdict
    monthly = defaultdict(lambda: {"sears_M": 0.0, "kmart_M": 0.0})
    for r in range((header_row or 1)+1, min((ws.max_row or 1000)+1, 3000)):
        dv = ws.cell(row=r, column=1).value
        if isinstance(dv, datetime):
            ym = dv.strftime("%Y-%m")
        elif hasattr(dv, "year"):
            ym = f"{dv.year}-{dv.month:02d}"
        else: continue
        sv = ws.cell(row=r, column=sears_col).value
        kv = ws.cell(row=r, column=kmart_col).value
        if isinstance(sv, (int, float)): monthly[ym]["sears_M"] += sv
        if isinstance(kv, (int, float)): monthly[ym]["kmart_M"] += kv

    return {ym: {"sears_M": round(d["sears_M"],4), "kmart_M": round(d["kmart_M"],4)}
            for ym, d in monthly.items()}


def process(label, partial):
    fp = find_file(partial)
    if not fp:
        print(f"  NOT FOUND: {partial}", flush=True)
        return {"label": label, "error": "not found", "monthly": {}, "annual": {}}

    print(f"\n[{label}] {os.path.basename(fp)}", flush=True)
    r = {"label": label, "filename": os.path.basename(fp), "monthly": {}, "annual": {}, "territory": {}, "store_counts": []}

    try:
        wb = openpyxl.load_workbook(fp, data_only=True, read_only=True)
        sheets = wb.sheetnames

        for sn in sheets:
            sn_l = sn.lower()
            try:
                ws = wb[sn]
                if "historica" in sn_l:
                    data = extract_historicals(ws)
                    if data:
                        r["monthly"].update(data)
                        print(f"  Historicals: {len(data)} months [{min(data)} to {max(data)}]", flush=True)

                elif "bu beta" in sn_l and "adj" not in sn_l:
                    ann = extract_bu_beta_annual_fast(ws)
                    if ann:
                        r["annual"].update(ann)
                        print(f"  BU beta annual: {sorted(ann.keys())}", flush=True)

                elif "inflows actuals" in sn_l:
                    md = extract_inflows_actuals_monthly(ws)
                    if md:
                        # Only use months not already captured
                        added = 0
                        for ym, d in md.items():
                            if ym not in r["monthly"]:
                                r["monthly"][ym] = d
                                added += 1
                        if added:
                            print(f"  Inflows Actuals: added {added} months", flush=True)

                # Store counts
                if "b&m" in sn_l or "anita" in sn_l or "sc if" in sn_l or "sc b&m" in sn_l:
                    SKWS = ["store count", "# stores", "open store", "b&m stores", "sears stores", "kmart stores"]
                    for rr in range(1, min((ws.max_row or 50)+1, 100)):
                        for cc in range(1, 6):
                            v = ws.cell(row=rr, column=cc).value
                            if v and any(k in str(v).lower() for k in SKWS):
                                vals = {}
                                for ccc in range(1, min((ws.max_column or 50)+1, 80)):
                                    vv = ws.cell(row=rr, column=ccc).value
                                    if isinstance(vv, (int, float)) and 0 < vv < 200:
                                        vals[ccc] = vv
                                if vals:
                                    r["store_counts"].append({"sheet": sn, "label": str(v), "row": rr, "values": vals})
                                    print(f"  Store count in '{sn}': {v} -> {vals}", flush=True)

            except Exception as e:
                print(f"  Sheet '{sn}' error: {e}", flush=True)

        wb.close()
    except Exception as e:
        print(f"  FILE ERROR: {e}", flush=True)
        r["error"] = str(e)

    r["monthly"] = dict(sorted(r["monthly"].items()))
    return r


def main():
    all_results = []
    for label, partial in BOD_FILES:
        res = process(label, partial)
        all_results.append(res)

    # Build the longitudinal summary table
    print("\n\n" + "="*120, flush=True)
    print("FULL MONTHLY INFLOWS ($M) — SEARS AND KMART ACTUALS + FORECASTS", flush=True)
    print("="*120, flush=True)

    # Collect all unique months
    all_months = sorted(set(ym for r in all_results for ym in r.get("monthly", {}).keys()))

    # Print header
    print(f"\n{'Month':<10}", end="", flush=True)
    for r in all_results:
        lbl = r['label'][:15]
        print(f"  {'S':>7} {'K':>7}", end="", flush=True)
    print(flush=True)
    print(f"{'':10}", end="", flush=True)
    for r in all_results:
        lbl = r['label'][:15]
        print(f"  {'Sears':>7} {'Kmart':>7}", end="", flush=True)
    print(flush=True)
    print("-"*120, flush=True)

    for ym in all_months:
        print(f"{ym:<10}", end="", flush=True)
        for r in all_results:
            d = r.get("monthly", {}).get(ym, {})
            s = d.get("sears_M", "")
            k = d.get("kmart_M", "")
            sv = f"{s:.2f}" if isinstance(s, float) else ""
            kv = f"{k:.2f}" if isinstance(k, float) else ""
            print(f"  {sv:>7} {kv:>7}", end="", flush=True)
        print(flush=True)

    print("\n\nANNUAL FY TOTALS FROM BU BETA:", flush=True)
    for r in all_results:
        if r.get("annual"):
            items = ", ".join(f"{fy}:S={d['sears_M']:.1f}M K={d['kmart_M']:.1f}M"
                             for fy, d in sorted(r["annual"].items()))
            print(f"  {r['label']:30s}: {items}", flush=True)

    print("\n\nSTORE COUNT DATA:", flush=True)
    for r in all_results:
        if r.get("store_counts"):
            for sc in r["store_counts"]:
                print(f"  {r['label']:28s} | {sc['sheet']:25s} | {sc['label']:40s} | {sc['values']}", flush=True)

    # Save JSON
    output = {
        "generated": "2026-05-11",
        "units": "Millions USD",
        "description": "Sears and Kmart retail inflows — longitudinal analysis across 16 BOD files (Feb 2022 – May 2026)",
        "methodology": "Historicals view KK (monthly aggregates) + BU beta (FY totals + monthly) + Inflows Actuals (daily aggregated to monthly). data_only=True.",
        "longitudinal_table": [
            {
                "bod_label": r["label"],
                "filename": r.get("filename", ""),
                "monthly_inflows_M": r.get("monthly", {}),
                "annual_fy_totals_M": r.get("annual", {}),
                "store_count_data": r.get("store_counts", []),
                "territory": r.get("territory", {}),
                "months_with_data": len(r.get("monthly", {})),
            }
            for r in all_results
        ],
        "all_months": all_months,
    }

    out = os.path.join(BASE_DIR, "analysis_retail_stores.json")
    with open(out, "w") as f:
        json.dump(output, f, indent=2, default=str)
    print(f"\nSaved: {out}", flush=True)


if __name__ == "__main__":
    main()
