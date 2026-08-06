"""
Yapı Kredi "Bireysel Ürün ve Hizmet Ücretleri" sayfasını çeken scraper modülü.
"""

import re
import sys
from dataclasses import dataclass
from typing import List

YAPIKREDI_URL = "https://www.yapikredi.com.tr/bireysel-bankacilik/hesaplama-araclari/bireysel-urun-ve-hizmet-ucretleri"

DATE_PATTERN = re.compile(
    r"G[üu]ncellenme\s*Tarihi\s*:\s*(\d{2}\.\d{2}\.\d{4}(?:\s+\d{2}:\d{2})?)",
    re.IGNORECASE,
)

DATE_PATTERN_TR = re.compile(
    r"(?:son\s+)?g[üu]ncellenme\s+tarihi\s*:?\s*"
    r"(\d{1,2})\s+(ocak|şubat|mart|nisan|mayıs|haziran|temmuz|ağustos|eylül|ekim|kasım|aralık)\s+(\d{4})",
    re.IGNORECASE,
)

TURKCE_AYLAR = {
    "ocak": "01", "şubat": "02", "mart": "03", "nisan": "04",
    "mayıs": "05", "haziran": "06", "temmuz": "07", "ağustos": "08",
    "eylül": "09", "ekim": "10", "kasım": "11", "aralık": "12",
}

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
    if match:
        tarih = match.group(1)
        temiz_aciklama = DATE_PATTERN.sub("", raw_aciklama).strip(" .")
        return temiz_aciklama, tarih

    match_tr = DATE_PATTERN_TR.search(raw_aciklama)
    if match_tr:
        gun = match_tr.group(1).zfill(2)
        ay = TURKCE_AYLAR.get(match_tr.group(2).lower(), "00")
        yil = match_tr.group(3)
        tarih = f"{gun}.{ay}.{yil}"
        temiz_aciklama = DATE_PATTERN_TR.sub("", raw_aciklama).strip(" .")
        return temiz_aciklama, tarih

    return raw_aciklama, ""


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
    col_tarih     = find_col(["güncelleme"])
    if col_tarih == -1:
        col_tarih = find_col(["guncelleme"])
    if col_masraf == -1:
        col_masraf = find_col(["işlem"])
    if col_masraf == -1:
        col_masraf = find_col(["ücret"])
    if col_asg_tutar == -1:
        col_asg_tutar = find_col(["tutar"])
    if col_azm_tutar == -1:
        col_azm_tutar = find_col(["azami"])
    if col_asg_oran == -1:
        col_asg_oran = find_col(["oran"])

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

        site_tarihi = get(col_tarih) if col_tarih >= 0 else ""
        temiz_aciklama, aciklama_tarihi = _parse_aciklama(get(col_aciklama))
        if not site_tarihi:
            site_tarihi = aciklama_tarihi

        satirlar.append(UcretSatiri(
            kategori=kategori, masraf=masraf,
            asgari_tutar=get(col_asg_tutar), asgari_oran=get(col_asg_oran),
            azami_tutar=get(col_azm_tutar), azami_oran=get(col_azm_oran),
            aciklama=temiz_aciklama, site_guncelleme_tarihi=site_tarihi,
        ))
    return satirlar


def scrape_yapikredi(url: str = YAPIKREDI_URL) -> List[UcretSatiri]:
    from playwright.sync_api import sync_playwright
    from bs4 import BeautifulSoup

    print(f"[yapikredi] {url} adresinden veri çekiliyor...", file=sys.stderr)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        try:
            context = browser.new_context(
                user_agent=HEADERS["User-Agent"],
                viewport={"width": 1280, "height": 800},
                locale="tr-TR",
                extra_http_headers={
                    "Accept-Language": "tr-TR,tr;q=0.9,en;q=0.8",
                    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
                    "Accept-Encoding": "gzip, deflate, br",
                }
            )
            page = context.new_page()
            page.goto(url, timeout=120000, wait_until="domcontentloaded")
            page.wait_for_timeout(8000)

            table_count = page.evaluate("() => document.querySelectorAll('table').length")
            print(f"[yapikredi] JS ile {table_count} <table> bulundu", file=sys.stderr)

            html = page.content()
        finally:
            browser.close()

    soup = BeautifulSoup(html, "lxml")
    tables = soup.find_all("table")
    print(f"[yapikredi] Toplam {len(tables)} <table> bulundu", file=sys.stderr)

    if not tables:
        raise ScraperError("Yapı Kredi sayfasında hiç <table> bulunamadı.")

    tum_satirlar = []
    for table in tables:
        kategori = _find_category_title(table, "Genel")
        rows = _extract_from_table(table, kategori)
        tum_satirlar.extend(rows)

    if not tum_satirlar:
        raise ScraperError("Yapı Kredi sayfasında tablo var ama hiç veri satırı çekilemedi.")

    print(f"[yapikredi] Toplam {len(tum_satirlar)} satır bulundu.", file=sys.stderr)
    return tum_satirlar


if __name__ == "__main__":
    veriler = scrape_yapikredi()
    for v in veriler[:5]:
        print(v)
    print(f"Toplam {len(veriler)} satır bulundu.")
