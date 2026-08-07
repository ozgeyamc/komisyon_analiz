from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup

def debug_banka(url, banka_adi, extra_js=None):
    print(f"\n{'='*60}")
    print(f"BANKA: {banka_adi}")
    print(f"{'='*60}")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        )
        page.goto(url, timeout=90000, wait_until="domcontentloaded")
        page.wait_for_timeout(5000)
        if extra_js:
            page.evaluate(extra_js)
            page.wait_for_timeout(2000)
        html = page.content()
        browser.close()

    soup = BeautifulSoup(html, "lxml")
    tables = soup.find_all("table")
    print(f"Toplam tablo: {len(tables)}")

    for i, table in enumerate(tables[:3]):
        thead = table.find("thead")
        if thead:
            hr = thead.find("tr")
            if hr:
                headers = [c.get_text(strip=True) for c in hr.find_all(["th","td"])]
                print(f"\nTablo {i+1} başlıkları: {headers}")
        rows = table.find_all("tr")
        print(f"İlk 3 satır son kolon ham metin:")
        for row in rows[1:4]:
            cells = row.find_all(["th","td"])
            if cells:
                print(f"  {repr(cells[-1].get_text(strip=False))}")

debug_banka(
    "https://www.garantibbva.com.tr/urun-ve-hizmet-ucretleri",
    "GARANTİ"
)

debug_banka(
    "https://www.halkbank.com.tr/tr/urun-ve-hizmet-ucretleri/kredi-kartlari-ve-banka-kartlari",
    "HALKBANK"
)

debug_banka(
    "https://www.denizbank.com/urun-ve-hizmet-ucretleri",
    "DENİZBANK",
    extra_js="""
        () => {
            document.querySelectorAll('.tab-pane').forEach(el => {
                el.classList.add('active', 'show');
                el.style.display = 'block';
            });
        }
    """
)
