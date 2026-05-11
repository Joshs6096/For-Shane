"""
Download strategic cash forecast files from SharePoint.
Uses Playwright to authenticate via existing browser session context.
Saves to /Users/josh/Downloads/SP_Analysis/
"""
import asyncio, os, re
from pathlib import Path
from playwright.async_api import async_playwright

OUT_DIR = Path("/Users/josh/Downloads/SP_Analysis")
OUT_DIR.mkdir(exist_ok=True)

BASE_URL = "https://searshc-my.sharepoint.com"
SP_PATH  = "/personal/keith_kim_transformco_com/Documents/Cash forecast - BU/Daily Excel Files"

# Strategic selection: BOD + deck versions, one per month per year
# Prioritize: BOD > deck version > BU view
STRATEGIC_FILES = [
    # ── 2022 ───────────────────────────────────────────────────────────────
    "Daily Cash Fcst - 9.19.22_BU view_Sep BOD.xlsx",
    "Daily Cash Fcst - 10.18.22_BU view_BOD.xlsx",
    "Daily Cash Fcst - 11.16.22_BU view_NovBOD.xlsx",
    "Daily Cash Fcst - 12.7.22_BU view_DecBOD Treasury_liquidity update final.xlsx",
    "Daily Cash Fcst - 12.30.22_BU view.xlsx",
    # ── 2023 ───────────────────────────────────────────────────────────────
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
    "Daily Cash Fcst - 12.30.22_BU view.xlsx",
    # ── 2024 ───────────────────────────────────────────────────────────────
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
    # ── 2025 ───────────────────────────────────────────────────────────────
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
    # ── 2026 ───────────────────────────────────────────────────────────────
    "Daily Cash Fcst - 01.16.26_BU view new format_asset sales_JanBOD.xlsx",
    "Daily Cash Fcst - 02.17.26_BU view_deck version.xlsx",
    "Daily Cash Fcst - 03.03.26_post-reductions_deck version_MarBOD.xlsx",
    "Daily Cash Fcst - 03.31.26_BU view_deck version.xlsx",
    "Daily Cash Fcst - 04.21.26_ BU version_deck version_AprBOD.xlsx",
    "Daily Cash Fcst - 05.05.26_BU view_deck version.xlsx",
    "Daily Cash Fcst - 05.08.26_BU view.xlsx",
    # Special: cash file details (metadata/mapping)
    "cash file details.xlsx",
]

def make_download_url(filename):
    from urllib.parse import quote
    path = f"{SP_PATH}/{filename}"
    return f"{BASE_URL}/_layouts/15/download.aspx?SourceUrl={quote(path)}"

async def download_all():
    async with async_playwright() as p:
        # Launch browser and navigate to SharePoint to establish session
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(accept_downloads=True)
        page = await context.new_page()

        print("Authenticating with SharePoint...")
        await page.goto(f"{BASE_URL}/personal/keith_kim_transformco_com/_layouts/15/onedrive.aspx?id=/personal/keith_kim_transformco_com/Documents/Cash forecast - BU/Daily Excel Files&ga=1")
        await page.wait_for_load_state('networkidle', timeout=30000)

        title = await page.title()
        print(f"  Auth result: {title}")

        downloaded = []
        skipped = []
        failed = []

        for i, filename in enumerate(STRATEGIC_FILES):
            out_path = OUT_DIR / filename
            if out_path.exists():
                print(f"  [{i+1}/{len(STRATEGIC_FILES)}] SKIP (exists): {filename[:60]}")
                skipped.append(filename)
                continue

            url = make_download_url(filename)
            print(f"  [{i+1}/{len(STRATEGIC_FILES)}] Downloading: {filename[:60]}")

            try:
                async with page.expect_download(timeout=120000) as dl_info:
                    await page.goto(url)

                dl = await dl_info.value
                await dl.save_as(str(out_path))
                size = out_path.stat().st_size
                print(f"    ✓ Saved: {size/1024/1024:.1f} MB")
                downloaded.append(filename)

            except Exception as e:
                print(f"    ✗ FAILED: {e}")
                failed.append((filename, str(e)))
                # Navigate back to folder to maintain session
                await page.goto(f"{BASE_URL}/personal/keith_kim_transformco_com/_layouts/15/onedrive.aspx?id=/personal/keith_kim_transformco_com/Documents/Cash forecast - BU/Daily Excel Files&ga=1")
                await page.wait_for_load_state('networkidle', timeout=15000)

        await browser.close()

        print(f"\n{'='*60}")
        print(f"COMPLETE: {len(downloaded)} downloaded, {len(skipped)} skipped, {len(failed)} failed")
        if failed:
            print("FAILED:")
            for f, e in failed: print(f"  {f}: {e}")

        return downloaded, skipped, failed

if __name__ == '__main__':
    asyncio.run(download_all())
