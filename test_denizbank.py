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

    for i in range(1, 10):
        result = page.evaluate(f"""
            () => {{
                const el = document.getElementById('h{i}');
                if (!el) return 'YOK';
                return {{
                    html: el.innerHTML.substring(0, 800),
                    childCount: el.children.length,
                    tableCount: el.querySelectorAll('table').length,
                    text: el.innerText.substring(0, 300),
                }};
            }}
        """)
        print(f"\n=== #h{i} ===")
        print(result)

    browser.close()
