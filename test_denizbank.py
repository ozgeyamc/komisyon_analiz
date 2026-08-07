from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup
import re

DATE_PATTERN = re.compile(
    r"G[üu]ncellenme\s*Tarihi\s*:\s*[\s\xa0]*(\d{2}[./]\d{2}[./]\d{4}(?:[\s\xa0]+\d{2}:\d{2})?)",
    re.IGNORECASE,
)
DATE_PATTERN_ITIBAR = re.compile(
    r"(\d{2}[./]\d{2}[./]\d{4})\s+tarihi\s+itibar",
    re.IGNORECASE,
)

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page(user_agent="Mozilla/5.0")
    page.goto("https://www.garantibbva.com.tr/urun-ve-hizmet-ucretleri", timeout=60000, wait_until="networkidle")
    page.wait_for_timeout(3000)
    html = page.content()
    browser.close()

soup = BeautifulSoup(html, "lxml")
bos = []
for table in soup.find_all("table"):
    for row in table.find_all("tr")[1:]:
        cells = row.find_all(["th","td"])
        if not cells: continue
        aciklama = cells[-1].get_text(strip=False)
        if not DATE_PATTERN.search(aciklama) and not DATE_PATTERN_ITIBAR.search(aciklama):
            bos.append(repr(aciklama[:200]))

print(f"Tarih bulunamayan satır sayısı: {len(bos)}")
# Benzersiz pattern'leri göster
unique = list(dict.fromkeys(bos))
for b in unique[:20]:
    print(" ", b)
