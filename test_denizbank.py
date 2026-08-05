from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup
import re

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
    page.wait_for_timeout(10000)
    html = page.content()
    browser.close()

soup = BeautifulSoup(html, "lxml")

# 1. Script tagları içinde JSON/veri var mı?
scripts = soup.find_all("script")
print(f"Toplam script tag: {len(scripts)}")
for i, s in enumerate(scripts):
    text = s.get_text()
    if any(k in text for k in ["ucret", "masraf", "komisyon", "tarife", "fee", "price"]):
        print(f"\n[Script {i+1}] (ilk 500):\n{text[:500]}")

# 2. data- attribute içinde veri var mı?
data_els = soup.find_all(attrs={"data-url": True})
print(f"\ndata-url attribute'lu eleman: {len(data_els)}")
for el in data_els[:10]:
    print(f"  {el.name}: {el.get('data-url')}")

data_src = soup.find_all(attrs={"data-src": True})
print(f"\ndata-src attribute'lu eleman: {len(data_src)}")
for el in data_src[:10]:
    print(f"  {el.name}: {el.get('data-src')}")

# 3. iframe var mı?
iframes = soup.find_all("iframe")
print(f"\niframe sayısı: {len(iframes)}")
for f in iframes:
    print(f"  src: {f.get('src')}")
