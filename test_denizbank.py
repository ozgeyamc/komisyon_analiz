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

    def handle_response(response):
        url = response.url
        ct = response.headers.get("content-type", "")
        # sadece html/json/text dönenler
        if any(k in ct for k in ["json", "text", "xml"]):
            try:
                body = response.body()
                print(f"\n[{response.status}] {url}")
                print(f"  CT: {ct}")
                print(f"  Body (ilk 200): {body[:200]}")
            except Exception as e:
                print(f"  [hata] {e}")

    page.on("response", handle_response)

    page.goto("https://www.isbank.com.tr/urun-ve-hizmet-ucretleri", timeout=90000, wait_until="domcontentloaded")
    page.wait_for_timeout(10000)
    browser.close()

print("\nBitti.")
