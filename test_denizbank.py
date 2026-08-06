from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup

GARANTI_URL = "https://www.garantibbva.com.tr/urun-ve-hizmet-ucretleri"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page(user_agent=HEADERS["User-Agent"])
    page.goto(GARANTI_URL, timeout=60000, wait_until="networkidle")
    page.wait_for_timeout(3000)
    html = page.content()
    browser.close()

soup = BeautifulSoup(html, "lxml")
tables = soup.find_all("table")
print(f"Toplam {len(tables)} tablo bulundu\n")

for i, table in enumerate(tables[:5]):
    thead = table.find("thead")
    if thead:
        hr = thead.find("tr")
        if hr:
            headers = [c.get_text(strip=True) for c in hr.find_all(["th", "td"])]
            print(f"Tablo {i+1} başlıkları: {headers}")
    rows = table.find_all("tr")
    for row in rows[1:3]:
        cells = [c.get_text(strip=True)[:60] for c in row.find_all(["th", "td"])]
        print(f"  Satır: {cells}")
    print()
