from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    context = browser.new_context(
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        locale="tr-TR",
        timezone_id="Europe/Istanbul",
    )
    context.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    page = context.new_page()
    page.goto("https://www.isbank.com.tr/urun-ve-hizmet-ucretleri", timeout=90000, wait_until="domcontentloaded")
    page.wait_for_timeout(15000)

    # header id'li TÜM elementleri bul
    result = page.evaluate("""
        () => {
            const all = document.querySelectorAll('[id*="header"]');
            return Array.from(all).map(el => ({
                id: el.id,
                tag: el.tagName,
                text: el.innerText ? el.innerText.substring(0, 80) : '',
                parentTag: el.parentElement ? el.parentElement.tagName : '',
                parentId: el.parentElement ? el.parentElement.id : '',
                parentClass: el.parentElement ? el.parentElement.className.substring(0, 80) : '',
            }));
        }
    """)
    for r in result:
        print(r)

    browser.close()
