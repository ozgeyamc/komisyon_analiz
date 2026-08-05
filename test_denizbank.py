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

    # header1..header10 id'li elementleri kontrol et
    for i in range(1, 15):
        el = page.query_selector(f"#header{i}")
        if el:
            text = el.inner_text()[:200].replace("\n", " ")
            print(f"#header{i}: {text}")

    # Tablo sayısını kontrol et
    tables = page.query_selector_all("table")
    print(f"\nToplam tablo: {len(tables)}")

    # div içinde ücret verisi ara
    html = page.content()
    browser.close()

soup = BeautifulSoup(html, "lxml")

# isbank.table.css kullanan div/section bul
for tag in ["div", "section", "article"]:
    els = soup.find_all(tag, class_=lambda c: c and any(k in " ".join(c) for k in ["table", "ucret", "fee", "price", "tarife", "komisyon"]))
    if els:
        print(f"\n{tag} eleman sayısı: {len(els)}")
        for el in els[:3]:
            print(f"  class={el.get('class')} id={el.get('id')}")
            print(f"  metin (ilk 200): {el.get_text()[:200].strip()}")
