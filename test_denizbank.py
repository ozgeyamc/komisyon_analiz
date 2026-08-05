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
    html = page.content()
    browser.close()

soup = BeautifulSoup(html, "lxml")

# h1 içindeki ilk UHU_item_icerikC bloğunun tüm class'larını ve içeriğini göster
h1 = soup.find(id="h1")
if h1:
    bloklar = h1.find_all(class_="UHU_item_icerikC")
    print(f"h1 içinde {len(bloklar)} icerikC bloğu var\n")
    for i, blok in enumerate(bloklar[:3]):
        print(f"--- Blok {i+1} ---")
        # Tüm child elemanları class + metin ile göster
        for child in blok.find_all(True):
            cls = child.get("class", [])
            text = child.get_text(strip=True)[:100]
            if text:
                print(f"  <{child.name} class={cls}>: {text}")
        print()
