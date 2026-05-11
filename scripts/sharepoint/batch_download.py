"""
Batch download strategic SharePoint files using itemId + driveId.
Downloads via the browser's authenticated session using base64 encoding.
Run this AFTER the browser session is authenticated.
"""
import asyncio, json, base64, os
from pathlib import Path
from playwright.async_api import async_playwright

OUT_DIR = Path("/Users/josh/Downloads/SP_Analysis")
OUT_DIR.mkdir(exist_ok=True)

# Strategic files to download (BOD + deck versions, one per month per year)
STRATEGIC_NAMES = [
    # 2022
    "Daily Cash Fcst - 9.19.22_BU view_Sep BOD.xlsx",
    "Daily Cash Fcst - 10.18.22_BU view_BOD.xlsx",
    "Daily Cash Fcst - 11.16.22_BU view_NovBOD.xlsx",
    "Daily Cash Fcst - 12.7.22_BU view_DecBOD Treasury_liquidity update final.xlsx",
    "Daily Cash Fcst - 12.30.22_BU view.xlsx",
    # 2023
    "Daily Cash Fcst - 1.17.23_BU view_JanBOD.xlsx",
    "Daily Cash Fcst - 1.31.23_BU view_Bohemia Adj New Methodology_Use for FY23a.xlsx",
    "Daily Cash Fcst - 2.27.23_BU view_BOD_3.2.23_liquidity update.xlsx",
    "Daily Cash Fcst - 3.31.23_BU view.xlsx",
    "Daily Cash Fcst - 4.25.23_BU view_April 27 BOD.xlsx",
    "Daily Cash Fcst - 5.26.23_BU view_24 adj_MayBOD_Liquidity & May actuals.xlsx",
    "Daily Cash Fcst - 6.20.23_BU view_JunBOD.xlsx",
    "Daily Cash Fcst - 7.25.23_BU view_JulBOD_7.27.23.xlsx",
    "Daily Cash Fcst - 8.22.23_BU view_AugBOD.xlsx",
    "Daily Cash Fcst - 9.19.23_BU view_SepBOD.xlsx",
    "Daily Cash Fcst - 10.24.23_BU view_HS retail interco_OctBOD.xlsx",
    "Daily Cash Fcst - 11.14.23_BU view_adUPDATE_11.14.23_BUPNL_HW_Assurant Nov gross_NovBOD.xlsx",
    # 2024
    "Daily Cash Fcst - 1.23.24_BU view_BOD.xlsx",
    "Daily Cash Fcst - 2.20.24_BU view_BOD.xlsx",
    "Daily Cash Fcst - 3.19.24_BU view_MarBOD.xlsx",
    "Daily Cash Fcst - 4.30.24_BU view_final_online adj.xlsx",
    "Daily Cash Fcst - 5.7.24_BU view_MayBOD_5.9.24.xlsx",
    "Daily Cash Fcst - 6.18.24_BU view_JunBOD.xlsx",
    "Daily Cash Fcst - 7.23.24_deck_JulBOD_FY25.xlsx",
    "Daily Cash Fcst - 8.19.24_BU view_AugBOD.xlsx",
    "Daily Cash Fcst - 9.24.24_BU view_deck version_BOD.xlsx",
    "Daily Cash Fcst - 10.15.24_BU view_deck version_OctBOD.xlsx",
    "Daily Cash Fcst - 11.12.24_BU view_deck version_NovBOD.xlsx",
    "Daily Cash Fcst - 12.10.24_BU view_deck version_BOD_SC wo 3 RE Interco IF_$9M.xlsx",
    "Daily Cash Fcst - 12.31.24_BU view.xlsx",
    # 2025
    "Daily Cash Fcst - 1.21.25_BU view_deck version_working_Apr append_deck2_JanBOD.xlsx",
    "Daily Cash Fcst - 2.18.25_BU view_deck version_FebBOD.xlsx",
    "Daily Cash Fcst - 3.18.25_BU view_deck version_MarBOD.xlsx",
    "Daily Cash Fcst - 4.18.25_BU view_prelim AprBOD.xlsx",
    "Daily Cash Fcst - 5.20.25_BU view_deck version_MayBOD.xlsx",
    "Daily Cash Fcst - 6.17.25_BU view_deck version_JunBOD.xlsx",
    "Daily Cash Fcst - 7.22.25_BU view_deck version_condensed_JulBOD.xlsx",
    "Daily Cash Fcst - 8.26.25_BU view_final deck version.xlsx",
    "Daily Cash Fcst - 9.30.25_BU view_deck version.xlsx",
    "Daily Cash Fcst - 10.21.25_BU view_deck version_OctBOD.xlsx",
    "Daily Cash Fcst - 11.18.25_BU view_deck version_NovBOD no meeting.xlsx",
    "Daily Cash Fcst - 12.15.25_BU view_DecBOD.xlsx",
    "Daily Cash Fcst - 12.30.25_BU view_deck version final.xlsx",
    # 2026
    "Daily Cash Fcst - 01.16.26_BU view new format_asset sales_JanBOD.xlsx",
    "Daily Cash Fcst - 02.17.26_BU view_deck version.xlsx",
    "Daily Cash Fcst - 03.03.26_post-reductions_deck version_MarBOD.xlsx",
    "Daily Cash Fcst - 03.31.26_BU view_deck version.xlsx",
    "Daily Cash Fcst - 04.21.26_ BU version_deck version_AprBOD.xlsx",
    "Daily Cash Fcst - 05.05.26_BU view_deck version.xlsx",
    "Daily Cash Fcst - 05.08.26_BU view.xlsx",
    "cash file details.xlsx",  # already done but include for completeness
]

DRIVE_ID = "b!5YH6v6JhxUClPPX4EPRoLOvOPIl2jWtDquaa19bpJwPVFivQQJUyQKC4fgUz3Xf-"

# Load item ID map
with open("/Users/josh/Documents/Finance Bots/sp_item_ids.json") as f:
    all_items = json.load(f)
item_map = {item['name']: item for item in all_items}

print(f"Item map loaded: {len(item_map)} files")
targets = [(n, item_map[n]) for n in STRATEGIC_NAMES if n in item_map]
missing = [n for n in STRATEGIC_NAMES if n not in item_map]
print(f"Targets found: {len(targets)}, Missing: {len(missing)}")
if missing:
    print("Missing:", missing[:5])

async def download_file_via_browser(page, item, out_path):
    """Download a file using the authenticated browser session."""
    drive_id = item['driveId']
    item_id = item['itemId']
    content_url = f"https://searshc-my.sharepoint.com/_api/v2.0/drives/{drive_id}/items/{item_id}/content"

    js = f"""
    async () => {{
      const resp = await fetch('{content_url}', {{ credentials: 'include' }});
      if (!resp.ok) return {{ error: resp.status }};
      const blob = await resp.blob();
      return new Promise(resolve => {{
        const reader = new FileReader();
        reader.onload = () => resolve({{ size: blob.size, b64: reader.result.split(',')[1] }});
        reader.readAsDataURL(blob);
      }});
    }}
    """
    result = await page.evaluate(js)

    if 'error' in result:
        return False, f"HTTP {result['error']}"

    content = base64.b64decode(result['b64'])
    with open(out_path, 'wb') as f:
        f.write(content)
    return True, len(content)

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context()
        page = await context.new_page()

        # Load SharePoint with existing session
        print("Loading SharePoint session...")
        await page.goto(
            "https://searshc-my.sharepoint.com/shared?listurl=https%3A%2F%2Fsearshc%2Dmy%2Esharepoint%2Ecom%2Fpersonal%2Fkeith%5Fkim%5Ftransformco%5Fcom%2FDocuments&id=%2Fpersonal%2Fkeith%5Fkim%5Ftransformco%5Fcom%2FDocuments%2FCash%20forecast%20%2D%20BU%2FDaily%20Excel%20Files&shareLink=1&ga=1",
            timeout=30000
        )
        await page.wait_for_load_state('networkidle', timeout=20000)
        title = await page.title()
        print(f"Session: {title}")

        downloaded = []; failed = []; skipped = []

        for i, (name, item) in enumerate(targets):
            out_path = OUT_DIR / name
            if out_path.exists() and out_path.stat().st_size > 10000:
                print(f"  [{i+1:2}/{len(targets)}] SKIP: {name[:55]}")
                skipped.append(name)
                continue

            print(f"  [{i+1:2}/{len(targets)}] Downloading: {name[:55]} ({item['size']/1e6:.1f}MB)")
            ok, info = await download_file_via_browser(page, item, out_path)
            if ok:
                print(f"    ✓ {info:,} bytes")
                downloaded.append(name)
            else:
                print(f"    ✗ {info}")
                failed.append((name, info))

        await browser.close()

        print(f"\n{'='*60}")
        print(f"Done: {len(downloaded)} downloaded, {len(skipped)} skipped, {len(failed)} failed")
        if failed:
            print("FAILED files:")
            for n, e in failed: print(f"  {n}: {e}")

        # List what we have
        files = sorted(OUT_DIR.glob("*.xlsx"))
        print(f"\nFiles in {OUT_DIR}: {len(files)}")
        total_mb = sum(f.stat().st_size for f in files) / 1e6
        print(f"Total size: {total_mb:.1f} MB")

if __name__ == '__main__':
    asyncio.run(main())
