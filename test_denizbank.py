from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    context = browser.new_context(
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        viewport={"width": 1280, "height": 800},
        locale="tr-TR",
        timezone_id="Europe/Istanbul",
    )
    context.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

    page = context.new_page()
    page.goto("https://www.isbank.com.tr/urun-ve-hizmet-ucretleri", timeout=90000, wait_until="networkidle")
    page.wait_for_timeout(8000)
    print("URL:", page.url)
    print("Başlık:", page.title())
    html = page.content()
    browser.close()

soup = BeautifulSoup(html, "lxml")
print("Page title tag:", soup.title.get_text() if soup.title else "YOK")
tables = soup.find_all("table")
print(f"Toplam tablo: {len(tables)}")

for i, table in enumerate(tables[:5]):
    rows = table.find_all("tr")
    text = table.get_text()[:200].strip().replace("\n", " ")
    print(f"\nTablo {i+1} ({len(rows)} satır): {text}")

# tablo yoksa div yapısına bak
if not tables:
    divs = soup.find_all("div", class_=True)
    classes = set()
    for d in divs[:300]:
        for c in d.get("class", []):
            classes.add(c)
    print(f"\nDiv class'ları: {sorted(classes)[:60]}")
    body = soup.find("body")
    if body:
        print(f"\nSayfa metni (ilk 500):\n{body.get_text()[:500]}")
