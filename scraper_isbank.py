"""
Türkiye İş Bankası "Ürün ve Hizmet Ücretleri" scraper modülü.
Bot koruması nedeniyle isbank.core.js'den endpoint tespit edilerek çekilir.
"""

import re
import sys
import requests
from dataclasses import dataclass
from typing import List

ISBANK_URL = "https://www.isbank.com.tr/urun-ve-hizmet-ucretleri"
ISBANK_CORE_JS = "https://www.isbank.com.tr/StaticFiles/Isbank/scripts/isbank.core.js"

DATE_PATTERN = re.compile(
    r"G[üu]ncell[ei]nme\s*Tarihi\s*:\s*(\d{2}\.\d{2}\.\d{4}(?:\s+\d{2}:\d{2})?)",
    re.IGNORECASE,
)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Referer": "https://www.isbank.com.tr/urun-ve-hizmet-ucretleri",
    "Accept-Language": "tr-TR,tr;q=0.9",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}


@dataclass
class UcretSatiri:
    kategori: str
    masraf: str
    asgari_tutar: str = ""
    asgari_oran: str = ""
    azami_tutar: str = ""
    azami_oran: str = ""
    aciklama: str = ""
    site_guncelleme_tarihi: str = ""


class ScraperError(Exception):
    pass


def _parse_aciklama(raw_aciklama: str):
    raw_aciklama = raw_aciklama.strip()
    match = DATE_PATTERN.search(raw_aciklama)
    tarih = match.group(1) if match else ""
    temiz_aciklama = DATE_PATTERN.sub("", raw_aciklama).strip(" .")
    return temiz_aciklama, tarih


def _normalize(val: str) -> str:
    return val.strip().replace("\xa0", " ").replace("\u200b", "").strip()


def _find_category_title(el, fallback: str) -> str:
    parent = el.parent
    depth = 0
    while parent is not None and depth < 10:
        for sibling in parent.find_all_previous(["h1", "h2", "h3", "h4", "h5"], limit=3):
            text = _normalize(sibling.get_text())
            if len(text) > 5 and text not in ["Müşteri Ol", "Ara", "Kapat", "Menü", "Ana Sayfa"]:
                return text
        parent = parent.parent
        depth += 1
    return fallback


def _extract_from_table(table, kategori: str) -> List[UcretSatiri]:
    from bs4 import BeautifulSoup
    satirlar = []
    thead = table.find("thead")
    tbody = table.find("tbody")

    header_texts = []
    if thead:
        hr = thead.find("tr")
        if hr:
            header_texts = [_normalize(c.get_text(strip=True)).lower() for c in hr.find_all(["th", "td"])]

    if tbody:
        data_rows = tbody.find_all("tr")
        if not header_texts and data_rows:
            header_texts = [_normalize(c.get_text(strip=True)).lower() for c in data_rows[0].find_all(["th", "td"])]
            data_rows = data_rows[1:]
    else:
        all_rows = table.find_all("tr")
        if not all_rows:
            return satirlar
        if not header_texts:
            header_texts = [_normalize(c.get_text(strip=True)).lower() for c in all_rows[0].find_all(["th", "td"])]
        data_rows = all_rows[1:]

    def find_col(keywords):
        for i, h in enumerate(header_texts):
            if all(k in h for k in keywords):
                return i
        return -1

    col_masraf    = find_col(["masraf"])
    col_asg_tutar = find_col(["asgari", "tutar"])
    col_asg_oran  = find_col(["asgari", "oran"])
    col_azm_tutar = find_col(["azami", "tutar"])
    col_azm_oran  = find_col(["azami", "oran"])
    col_aciklama  = find_col(["açıklama"])
    if col_aciklama == -1:
        col_aciklama = find_col(["aciklama"])

    if col_masraf == -1:
        col_masraf = 0; col_asg_tutar = 1; col_asg_oran = 2
        col_azm_tutar = 3; col_azm_oran = 4; col_aciklama = 5

    for row in data_rows:
        cells = row.find_all(["th", "td"])
        if not cells or len(cells) < 2:
            continue
        values = [_normalize(c.get_text(strip=True)) for c in cells]

        def get(idx):
            return values[idx] if 0 <= idx < len(values) else ""

        masraf = get(col_masraf)
        if not masraf:
            continue

        temiz_aciklama, site_tarihi = _parse_aciklama(get(col_aciklama))
        satirlar.append(UcretSatiri(
            kategori=kategori, masraf=masraf,
            asgari_tutar=get(col_asg_tutar), asgari_oran=get(col_asg_oran),
            azami_tutar=get(col_azm_tutar), azami_oran=get(col_azm_oran),
            aciklama=temiz_aciklama, site_guncelleme_tarihi=site_tarihi,
        ))
    return satirlar


def _find_endpoints_in_js() -> List[str]:
    """isbank.core.js içindeki endpoint URL'lerini bul."""
    try:
        r = requests.get(ISBANK_CORE_JS, headers=HEADERS, timeout=30)
        r.raise_for_status()
        js_content = r.text

        # URL pattern'lerini ara
        urls = re.findall(r'["\']([^"\']*(?:ucret|price|fee|tarim|servis|Service|getPage|GetPage|content|Content)[^"\']*)["\']', js_content, re.IGNORECASE)
        print(f"[isbank] isbank.core.js'de bulunan URL pattern'leri:", file=sys.stderr)
        for u in urls[:20]:
            print(f"  {u}", file=sys.stderr)
        return urls
    except Exception as e:
        print(f"[isbank] isbank.core.js okunamadı: {e}", file=sys.stderr)
        return []


def scrape_isbank(url: str = ISBANK_URL) -> List[UcretSatiri]:
    from playwright.sync_api import sync_playwright
    from bs4 import BeautifulSoup

    print(f"[isbank] {url} adresinden veri çekiliyor...", file=sys.stderr)

    # JS'den endpoint bulmayı dene
    _find_endpoints_in_js()

    captured_html_responses = {}

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        try:
            context = browser.new_context(
                user_agent=HEADERS["User-Agent"],
                extra_http_headers={
                    "Accept-Language": "tr-TR,tr;q=0.9",
                    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                }
            )
            page = context.new_page()

            # HTML response'larını body ile birlikte yakala
            def handle_response(response):
                try:
                    ctype = response.headers.get("content-type", "")
                    rurl = response.url
                    if "text/html" in ctype and "TSPD" not in rurl and "google" not in rurl:
                        captured_html_responses[rurl] = response
                except Exception:
                    pass

            page.on("response", handle_response)

            page.goto(url, timeout=120000, wait_until="domcontentloaded")
            page.wait_for_timeout(15000)

            # HTML response'larının içeriğini logla
            print(f"[isbank] HTML response'ları:", file=sys.stderr)
            for rurl, resp in captured_html_responses.items():
                try:
                    body = resp.text()
                    soup_r = BeautifulSoup(body, "lxml")
                    tables_r = soup_r.find_all("table")
                    text_preview = soup_r.get_text()[:200].replace("\n", " ")
                    print(f"  {rurl[:80]} — {len(tables_r)} tablo — {text_preview[:100]}", file=sys.stderr)
                except Exception as e:
                    print(f"  {rurl[:80]} — okunamadı: {e}", file=sys.stderr)

            # Sayfadaki tüm <script> içeriklerinde endpoint ara
            inline_scripts = page.evaluate("""
                () => Array.from(document.querySelectorAll('script:not([src])')).map(s => s.textContent).join('\\n')
            """)
            # ucret/servis içeren satırları bul
            for line in inline_scripts.split("\n"):
                if any(k in line.lower() for k in ["ucret", "servis", "getpage", "service", "ajax", "fetch", "url"]):
                    print(f"  [inline-js] {line.strip()[:150]}", file=sys.stderr)

            html = page.content()
        finally:
            browser.close()

    soup = BeautifulSoup(html, "lxml")
    tables = soup.find_all("table")
    print(f"[isbank] Toplam {len(tables)} <table> bulundu", file=sys.stderr)

    if not tables:
        raise ScraperError("İş Bankası sayfasında hiç <table> bulunamadı.")

    tum_satirlar = []
    for table in tables:
        kategori = _find_category_title(table, "Genel")
        rows = _extract_from_table(table, kategori)
        tum_satirlar.extend(rows)

    if not tum_satirlar:
        raise ScraperError("İş Bankası sayfasında tablo var ama hiç veri satırı çekilemedi.")

    print(f"[isbank] Toplam {len(tum_satirlar)} satır bulundu.", file=sys.stderr)
    return tum_satirlar


if __name__ == "__main__":
    veriler = scrape_isbank()
    for v in veriler[:5]:
        print(v)
    print(f"Toplam {len(veriler)} satır bulundu.")
