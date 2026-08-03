"""
Türkiye İş Bankası "Ürün ve Hizmet Ücretleri" sayfasını çeken scraper modülü.
"""

import re
import sys
from dataclasses import dataclass
from typing import List

ISBANK_URL = "https://www.isbank.com.tr/urun-ve-hizmet-ucretleri"

DATE_PATTERN = re.compile(
    r"G[üu]ncell[ei]nme\s*Tarihi\s*:\s*(\d{2}\.\d{2}\.\d{4}(?:\s+\d{2}:\d{2})?)",
    re.IGNORECASE,
)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
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


def scrape_isbank(url: str = ISBANK_URL) -> List[UcretSatiri]:
    from playwright.sync_api import sync_playwright
    from bs4 import BeautifulSoup

    print(f"[isbank] {url} adresinden veri çekiliyor...", file=sys.stderr)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        try:
            page = browser.new_page(user_agent=HEADERS["User-Agent"])

            # Network isteklerini yakala — API endpoint'i bulmak için
            api_urls = []
            def handle_request(request):
                if any(x in request.url for x in ["json", "api", "service", "ucret", "price", "fee", "data"]):
                    api_urls.append(request.url)
            page.on("request", handle_request)

            page.goto(url, timeout=120000, wait_until="domcontentloaded")
            page.wait_for_timeout(10000)

            # Yakalanan API URL'lerini logla
            print(f"[isbank] Yakalanan API istekleri ({len(api_urls)} adet):", file=sys.stderr)
            for u in api_urls[:20]:
                print(f"  {u}", file=sys.stderr)

            # Tab'ları bul ve tıkla
            tab_selectors = [
                ".tabContentC",
                "[class*='tab']",
                "[class*='Tab']",
                "[class*='Cover_M']",
                "button[aria-expanded='false']",
                ".accordion-button.collapsed",
                "[data-bs-toggle='collapse']",
                "[class*='accordion']",
                "li[role='tab']",
                ".nav-link",
            ]
            for selector in tab_selectors:
                try:
                    elements = page.query_selector_all(selector)
                    for el in elements:
                        try:
                            el.click(timeout=1500)
                            page.wait_for_timeout(500)
                        except Exception:
                            continue
                except Exception:
                    continue

            page.wait_for_timeout(5000)
            html = page.content()

            # Sayfanın içeriğini logla
            soup_debug = BeautifulSoup(html, "lxml")
            body = soup_debug.find("body")
            if body:
                text = body.get_text(separator="\n", strip=True)
                # Ücret/masraf içeren satırları bul
                lines = [l for l in text.split("\n") if any(k in l.lower() for k in ["masraf", "ücret", "komisyon", "tl", "%"])]
                print(f"[isbank] Ücret içeren satırlar ({len(lines)} adet):", file=sys.stderr)
                for l in lines[:20]:
                    print(f"  {l}", file=sys.stderr)

            # Tüm div class'larını logla
            divs = soup_debug.find_all("div", class_=True)
            classes = set()
            for d in divs[:500]:
                for c in d.get("class", []):
                    classes.add(c)
            print(f"[isbank] Tüm div class'ları: {sorted(classes)}", file=sys.stderr)

        finally:
            browser.close()

    soup = BeautifulSoup(html, "lxml")
    tables = soup.find_all("table")
    print(f"[isbank] Toplam {len(tables)} <table> bulundu", file=sys.stderr)

    if not tables:
        raise ScraperError("İş Bankası sayfasında hiç <table> bulunamadı — log'dan yapıyı inceleyin.")

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
