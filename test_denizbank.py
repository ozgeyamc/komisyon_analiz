from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page()
    page.goto("https://www.teb.com.tr/urun-ve-hizmet-ucretleri/", timeout=90000, wait_until="domcontentloaded")
    page.wait_for_timeout(5000)

    for selector in [
        "button[aria-expanded='false']",
        ".accordion-button.collapsed",
        "[data-bs-toggle='collapse']",
        "li[role='tab']",
        ".nav-link",
    ]:
        try:
            elements = page.query_selector_all(selector)
            for el in elements:
                try:
                    el.click(timeout=1500)
                    page.wait_for_timeout(200)
                except Exception:
                    continue
        except Exception:
            continue

    page.wait_for_timeout(3000)
    html = page.content()
    browser.close()

soup = BeautifulSoup(html, "lxml")
tables = soup.find_all("table")
print(f"Toplam tablo: {len(tables)}")

for i, table in enumerate(tables):
    text = table.get_text()[:200].strip().replace("\n", " ")
    rows = table.find_all("tr")
    print(f"\nTablo {i+1} ({len(rows)} satır): {text}")
