from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup
import json

api_responses = []

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    context = browser.new_context(
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        locale="tr-TR",
        timezone_id="Europe/Istanbul",
    )
    context.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

    page = context.new_page()

    # Tüm network isteklerini yakala
    def handle_response(response):
        url = response.url
        if any(k in url for k in ["ucret", "fee", "commission", "price", "tarife", "api", "json", "data"]):
            try:
                body = response.body()
                print(f"\n[API] {response.status} {url}")
                print(f"Content-Type: {response.headers.get('content-type','')}")
                print(f"İlk 300 byte: {body[:300]}")
            except Exception as e:
                print(f"[API hata] {url}: {e}")

    page.on("response", handle_response)

    page.goto("https://www.isbank.com.tr/urun-ve-hizmet-ucretleri", timeout=90000, wait_until="domcontentloaded")
    page.wait_for_timeout(10000)
    browser.close()

print("\nBitti.")
