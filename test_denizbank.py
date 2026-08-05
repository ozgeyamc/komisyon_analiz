from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page()
    page.goto("https://www.isbank.com.tr/urun-ve-hizmet-ucretleri", timeout=90000, wait_until="domcontentloaded")
    page.wait_for_timeout(5000)
    print("Gerçek URL:", page.url)
    print("Başlık:", page.title())
    html = page.content()
    browser.close()

soup = BeautifulSoup(html, "lxml")
tables = soup.find_all("table")
print(f"\nToplam tablo: {len(tables)}")

for i, table in enumerate(tables):
    rows = table.find_all("tr")
    text = table.get_text()[:200].strip().replace("\n", " ")
    print(f"\nTablo {i+1} ({len(rows)} satır): {text}")
