from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page()
    page.goto("https://www.denizbank.com/urun-ve-hizmet-ucretleri", timeout=90000, wait_until="domcontentloaded")
    page.wait_for_timeout(5000)

    # nav-0 dan nav-6 ya kadar tüm tab'ları tıkla
    for i in range(7):
        try:
            page.click(f"a[href='#nav-{i}']", timeout=3000)
            page.wait_for_timeout(1000)
            print(f"[tab nav-{i}] tiklandi")
        except Exception as e:
            print(f"[tab nav-{i}] tiklanamadi: {e}")

    page.wait_for_timeout(2000)
    html = page.content()
    browser.close()

soup = BeautifulSoup(html, "lxml")

# tb-1 den tb-9 a kadar her bolumun basligini ve tablo sayisini logla
for i in range(1, 10):
    bolum = soup.find(id=f"tb-{i}")
    if bolum:
        baslik = bolum.find(["h1","h2","h3","h4","h5"])
        tablolar = bolum.find_all("table")
        print(f"\n[tb-{i}] Baslik: {baslik.get_text(strip=True) if baslik else 'YOK'} | Tablo sayisi: {len(tablolar)}")
        for j, t in enumerate(tablolar):
            print(f"  Tablo {j+1} (ilk 200): {t.get_text()[:200]}")
