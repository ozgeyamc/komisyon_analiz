from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page()
    page.goto("https://www.denizbank.com/urun-ve-hizmet-ucretleri", timeout=90000, wait_until="domcontentloaded")
    page.wait_for_timeout(5000)
    print("Gercek URL:", page.url)
    print("Baslik:", page.title())
    links = page.evaluate("""
        () => Array.from(document.querySelectorAll('a[href]'))
            .map(a => a.href)
            .filter(h => h.includes('ucret') || h.includes('tarife') || h.includes('fiyat'))
    """)
    for l in links[:30]:
        print("Link:", l)
    browser.close()
