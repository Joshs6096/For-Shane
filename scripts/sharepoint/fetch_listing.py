import sys, json
from playwright.sync_api import sync_playwright

url = "https://www.ebay.com/itm/117104155092"
output_dir = "/tmp"

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    ctx = browser.new_context(storage_state="/Users/josh/.openclaw/workspace/.ebay_storage_state.json")
    page = ctx.new_page()
    page.goto(url, wait_until="domcontentloaded", timeout=30000)
    page.wait_for_timeout(2000)
    
    # Get title + main image
    title = page.title()
    print(f"TITLE: {title}")
    
    # Try to find image URLs from the carousel
    images = page.evaluate("""() => {
        const imgs = Array.from(document.querySelectorAll('img'));
        return imgs
            .map(i => i.src || i.dataset.src)
            .filter(s => s && (s.includes('ebayimg.com') || s.includes('i.ebayimg.com')))
            .filter(s => s.includes('s-l') && !s.includes('s-l64') && !s.includes('s-l140'))
            .slice(0, 10);
    }""")
    print(f"IMAGES_FOUND: {len(images)}")
    for i, img in enumerate(images[:6]):
        print(f"  IMG{i}: {img}")
    
    # Get description text
    desc = page.evaluate("""() => {
        const el = document.querySelector('#viTabs_0_is') || document.querySelector('.x-item-description-child');
        return el ? el.innerText.slice(0, 500) : '';
    }""")
    print(f"DESC: {desc[:400]}")
    
    # Screenshot the photos section
    page.screenshot(path=f"{output_dir}/scruffy_norfed_full.png", full_page=False)
    print(f"SCREENSHOT: {output_dir}/scruffy_norfed_full.png")
    
    browser.close()
