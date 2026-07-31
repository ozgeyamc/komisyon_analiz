"""
Halkbank "Ürün ve Hizmet Ücretleri" sayfasını çeken scraper modülü.
"""

import re
import sys
from dataclasses import dataclass
from typing import List, Optional

import requests
from bs4 import BeautifulSoup

HALKBANK_URL = "https://www.halkbank.com.tr/tr/urun-ve-hizmet-ucretleri"

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

# Halkbank 9 alt kategori sayfası
HALKBANK_ALT_SAYFALAR = [
    ("Kredi Kartları ve Banka Kartları", "https://www.halkbank.com.tr/tr/urun-ve-hizmet-ucretleri/kredi-kartlari-ve-banka-kartlari"),
    ("Havale EFT FAST", "https://www.halkbank.com.tr/tr/urun-ve-hizmet-ucretleri/havale-eft-fast"),
    ("Krediler", "https://www.halkbank.com.tr/tr/urun-ve-hizmet-ucretleri/krediler"),
    ("Mevduat İşlemleri", "https://www.halkbank.com.tr/tr/urun-ve-hizmet-ucretleri/mevduat-islemleri"),
    ("Menkul Kıymet İşlemleri", "https://www.halkbank.com.tr/tr/urun-ve-hizmet-ucretleri/menkul-kiymet-islemleri"),
    ("Dış Ticaret İşlemleri", "https://www.halkbank.com.tr/tr/urun-ve-hizmet-ucretleri/dis-ticaret-islemleri"),
    ("Çekler ve Senetler", "https://www.halkbank.com.tr/tr/urun-ve-hizmet-ucretleri/cekler-ve-senetler"),
    ("Sigorta İşlemleri", "https://www.halkbank.com.tr/tr/urun-ve-hizmet-ucretleri/sigorta-islemleri"),
    ("Diğer İşlemler", "https://www.halkbank.com.tr/tr/urun-ve-hizmet-ucretleri/diger-islemler"),
]


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


def _find_category_title(table, sayfa_kategorisi: str) -> str:
    el = table.parent
    depth = 0
    while el is not None and depth < 8:
        for sibling in el.find_all_previous(["h1", "h2", "h3", "h4", "h5", "button"], limit=3):
            text = _normalize(sibling.get_text())
            if len(text) > 5 and text not in ["Müşteri Ol", "Ara", "Kapat", "Menü"]:
                return text
        el = el.parent
        depth += 1
    return sayfa_kategorisi


def _extract_rows_from_table(table, kategori: str) -> List[UcretSatiri]:
    satirlar: List[UcretSatiri] = []

    thead = table.find("thead")
    tbody = table.find("tbody")

    header_texts = []
    if thead:
        header_row = thead.find("tr")
        if header_row:
            header_texts = [
                _normalize(c.get_text(strip=True)).lower()
                for c in header_row.find_all(["th", "td"])
            ]

    if tbody:
        data_rows = tbody.find_all("tr")
        if not header_texts and data_rows:
            header_texts = [
                _normalize(c.get_text(strip=True)).lower()
                for c in data_rows[0].find_all(["th", "td"])
            ]
            data_rows = data_rows[1:]
    else:
        all_rows = table.find_all("tr")
        if not all_rows:
            return satirlar
        if not header_texts:
            header_texts = [
                _normalize(c.get_text(strip=True)).lower()
                for c in all_rows[0].find_all(["th", "td"])
            ]
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
        col_masraf    = 0
        col_asg_tutar = 1
        col_asg_oran  = 2
        col_azm_tutar = 3
        col_azm_oran  = 4
        col_aciklama  = 5

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

        raw_aciklama = get(col_aciklama)
        temiz_aciklama, site_tarihi = _parse_aciklama(raw_aciklama)

        satirlar.append(
            UcretSatiri(
                kategori=kategori,
                masraf=masraf,
                asgari_tutar=get(col_asg_tutar),
                asgari_oran=get(col_asg_oran),
                azami_tutar=get(col_azm_tutar),
                azami_oran=get(col_azm_oran),
                aciklama=temiz_aciklama,
                site_guncelleme_tarihi=site_tarihi,
            )
        )

    return satirlar


def _scrape_sayfa(url: str, sayfa_kategorisi: str) -> List[UcretSatiri]:
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise ScraperError("Playwright kurulu değil.") from exc

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        try:
            page = browser.new_page(user_agent=HEADERS["User-Agent"])
            page.goto(url, timeout=60000, wait_until="networkidle")

            for selector in [
                "button[aria-expanded='false']",
                ".accordion-button.collapsed",
                "[data-bs-toggle='collapse']",
                ".card-header button",
                "li[role='tab']",
                ".nav-link:not(.active)",
            ]:
                try:
                    elements = page.query_selector_all(selector)
                    for el in elements:
                        try:
                            el.click(timeout=2000)
                            page.wait_for_timeout(300)
                        except Exception:
                            continue
                except Exception:
                    continue

            page.wait_for_timeout(2000)
            html = page.content()
        finally:
            browser.close()

    soup = BeautifulSoup(html, "lxml")
    tables = soup.find_all("table")

    print(f"  [halkbank] '{sayfa_kategorisi}' — {len(tables)} tablo bulundu", file=sys.stderr)

    tum_satirlar: List[UcretSatiri] = []
    for table in tables:
        kategori = _find_category_title(table, sayfa_kategorisi)
        rows = _extract_rows_from_table(table, kategori)
        tum_satirlar.extend(rows)

    return tum_satirlar


def scrape_halkbank() -> List[UcretSatiri]:
    print(f"[halkbank] 9 alt sayfa çekiliyor...", file=sys.stderr)

    tum_satirlar: List[UcretSatiri] = []

    for sayfa_adi, url in HALKBANK_ALT_SAYFALAR:
        print(f"[halkbank] '{sayfa_adi}' çekiliyor...", file=sys.stderr)
        try:
            satirlar = _scrape_sayfa(url, sayfa_adi)
            print(f"[halkbank] '{sayfa_adi}' — {len(satirlar)} satır bulundu.", file=sys.stderr)
            tum_satirlar.extend(satirlar)
        except Exception as exc:
            print(f"[halkbank] '{sayfa_adi}' çekilemedi: {exc}", file=sys.stderr)
            continue

    if not tum_satirlar:
        raise ScraperError("Halkbank sayfalarından hiçbir ücret satırı çekilemedi.")

    print(f"[halkbank] Toplam {len(tum_satirlar)} satır bulundu.", file=sys.stderr)
    return tum_satirlar


if __name__ == "__main__":
    veriler = scrape_halkbank()
    for v in veriler[:5]:
        print(v)
    print(f"Toplam {len(veriler)} satır bulundu.")
