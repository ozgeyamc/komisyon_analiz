from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup

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

    # header1'in parent div'ini ve içindeki tüm HTML'i al
    for i in range(1, 5):
        try:
            inner = page.evaluate(f"""
                () => {{
                    const el = document.getElementById('header{i}');
                    if (!el) return 'YOK';
                    // parent birkaç seviye yukarı çık
                    let parent = el.parentElement;
                    for (let j = 0; j < 4; j++) {{
                        if (parent && parent.parentElement) parent = parent.parentElement;
                    }}
                    return parent ? parent.innerHTML.substring(0, 1000) : 'parent yok';
                }}
            """)
            print(f"\n=== header{i} parent HTML (ilk 1000) ===")
            print(inner)
        except Exception as e:
            print(f"header{i} hata: {e}")

    browser.close()
