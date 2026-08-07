from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page(user_agent="Mozilla/5.0")
    page.goto("https://www.ziraatbank.com.tr/tr/urun-ve-hizmet-ucretleri", timeout=60000, wait_until="networkidle")
    page.wait_for_timeout(3000)
    html = page.content()
    browser.close()

soup = BeautifulSoup(html, "lxml")
tables = soup.find_all("table")

print(f"Toplam tablo: {len(tables)}")
for i, table in enumerate(tables[:3]):
    rows = table.find_all("tr")
    print(f"\nTablo {i+1} — son kolon ham değerler (ilk 3 satır):")
    for row in rows[:4]:
        cells = row.find_all(["th","td"])
        if cells:
            print(f"  son hücre: {repr(cells[-1].get_text(strip=False)[:80])}")
