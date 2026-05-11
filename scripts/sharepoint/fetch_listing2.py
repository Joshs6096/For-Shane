import sys
from playwright.sync_api import sync_playwright

url = "https://www.ebay.com/itm/117104155092"

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    ctx = browser.new_context(storage_state="/Users/josh/.openclaw/workspace/.ebay_storage_state.json")
    page = ctx.new_page()
    page.goto(url, wait_until="domcontentloaded", timeout=30000)
    page.wait_for_timeout(3000)
    
    # Click on each thumbnail to load all images, then grab them
    # The main image is typically in a specific selector
    images = page.evaluate("""() => {
        // Try common eBay selectors for the main listing carousel
        const selectors = [
            'img.ux-image-carousel-item__image',
            '.ux-image-carousel-item img',
            'img[data-zoom-src]',
            'div.ux-image-carousel img',
            '.image-treatment img'
        ];
        const out = new Set();
        for (const sel of selectors) {
            document.querySelectorAll(sel).forEach(img => {
                const src = img.dataset.zoomSrc || img.dataset.src || img.src;
                if (src && src.includes('ebayimg.com') && src.includes('s-l')) {
                    out.add(src);
                }
            });
        }
        return [...out].slice(0, 12);
    }""")
    print(f"LISTING IMAGES: {len(images)}")
    for i, img in enumerate(images):
        print(f"  IMG{i}: {img}")
    
    # Also screenshot the photos area specifically
    try:
        photo_el = page.locator('.ux-image-carousel-container, .image-treatment').first
        if photo_el:
            photo_el.screenshot(path="/tmp/listing_photos.png")
            print("PHOTOS_SCREENSHOT: /tmp/listing_photos.png")
    except Exception as e:
        print(f"screenshot err: {e}")
    
    browser.close()
