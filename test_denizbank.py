import requests
from bs4 import BeautifulSoup

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "tr-TR,tr;q=0.9,en-US;q=0.8,en;q=0.7",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "none",
    "Sec-Fetch-User": "?1",
}

session = requests.Session()
resp = session.get("https://www.isbank.com.tr/urun-ve-hizmet-ucretleri", headers=headers, timeout=30)
print("Status:", resp.status_code)
print("Başlık:", resp.url)

soup = BeautifulSoup(resp.text, "lxml")
print("Page title:", soup.title.get_text() if soup.title else "YOK")

tables = soup.find_all("table")
print(f"Toplam tablo: {len(tables)}")
for i, table in enumerate(tables[:5]):
    rows = table.find_all("tr")
    text = table.get_text()[:200].strip().replace("\n", " ")
    print(f"\nTablo {i+1} ({len(rows)} satır): {text}")
