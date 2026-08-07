from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup
import re

DATE_PATTERN = re.compile(
    r"G[üu]ncellenme\s*Tarihi\s*:\s*[\s\xa0]*(\d{2}[./]\d{2}[./]\d{4}(?:[\s\xa0]+\d{2}:\d{2})?)",
    re.IGNORECASE,
)

# ── GARANTİ ──────────────────────────────────────────────
print("\n=== GARANTİ — boş kalan satırlar ===")
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
        if not DATE_PATTERN.search(aciklama):
            bos.append(repr(aciklama[:120]))

print(f"Tarih bulunamayan satır sayısı: {len(bos)}")
for b in bos[:10]:
    print(" ", b)

# ── DENİZBANK ────────────────────────────────────────────
print("\n=== DENİZBANK — ilk 5 tablonun tam yapısı ===")
with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page(user_agent="Mozilla/5.0")
    page.goto("https://www.denizbank.com/urun-ve-hizmet-ucretleri", timeout=90000, wait_until="domcontentloaded")
    page.wait_for_timeout(5000)
    page.evaluate("""
        () => document.querySelectorAll('.tab-pane').forEach(el => {
            el.classList.add('active','show');
            el.style.display='block';
        })
    """)
    page.wait_for_timeout(2000)
    html = page.content()
    browser.close()

soup = BeautifulSoup(html, "lxml")
tb1 = soup.find(id="tb-1")
if tb1:
    tables = tb1.find_all("table")
    print(f"tb-1 içinde {len(tables)} tablo")
    for i, t in enumerate(tables[:5]):
        rows = t.find_all("tr")
        print(f"\n  Tablo {i+1} ({len(rows)} satır):")
        for row in rows[:4]:
            cells = [repr(c.get_text(strip=False)[:60]) for c in row.find_all(["th","td"])]
            print(f"    {cells}")
