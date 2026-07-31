"""
Garanti BBVA "Ürün ve Hizmet Ücretleri" sayfasını çeken scraper modülü.
"""

import re
import sys
from dataclasses import dataclass
from typing import List, Optional

import requests
from bs4 import BeautifulSoup

GARANTI_URL = "https://www.garantibbva.com.tr/urun-ve-hizmet-ucretleri"

DATE_PATTERN = re.compile(
    r"G[üu]ncell[ei]nme\s*Tarihi\s*:\s*(\d{2}\.\d{2}\.\d{4})",
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


def _extract_rows_from_table(table, kategori: str) -> List[UcretSatiri]:
    satirlar: List[UcretSatiri] = []
    rows = table.find_all("tr")
    if not rows:
        return satirlar

    for row in rows:
        cells = row.find_all("td")
        if not cells or len(cells) < 2:
            continue

        values = [c.get_text(strip=True) for c in cells]

        masraf = values[0] if len(values) > 0 else ""
        asgari_tutar = values[1] if len(values) > 1 else ""
        asgari_oran = values[2] if len(values) > 2 else ""
        azami_tutar = values[3] if len(values) > 3 else ""
        azami_oran = values[4] if len(values) > 4 else ""
        raw_aciklama = values[5] if len(values) > 5 else ""

        if not masraf:
            continue

        temiz_aciklama, site_tarihi = _parse_aciklama(raw_aciklama)

        satirlar.append(
            UcretSatiri(
                kategori=kategori,
                masraf=masraf,
                asgari_tutar=asgari_tutar,
                asgari_oran=asgari_oran,
                azami_tutar=azami_tutar,
                azami_oran=azami_oran,
                aciklama=temiz_aciklama,
                site_guncelleme_tarihi=site_tarihi,
            )
        )

    return satirlar


def _find_category_title(table) -> str:
    parent = table.find_parent()
    depth = 0
    while parent is not None and depth < 6:
        heading = parent.find_previous(["h1", "h2", "h3", "h4", "button"], recursive=False)
        if heading and heading.get_text(strip=True):
            return heading.get_text(strip=True)
        parent = parent.find_parent()
        depth += 1

    heading = table.find_previous(["h1", "h2", "h3", "h4", "button"])
    if heading and heading.get_text(strip=True):
        return heading.get_text(strip=True)

    return "Genel"


def _scrape_with_requests(url: str = GARANTI_URL) -> Optional[List[UcretSatiri]]:
    try:
        response = requests.get(url, headers=HEADERS, timeout=30)
        response.raise_for_status()
    except requests.RequestException as exc:
        raise ScraperError(f"Sayfaya erişilemedi: {exc}") from exc

    soup = BeautifulSoup(response.text, "lxml")
    tables = soup.find_all("table")

    if not tables:
        return None

    tum_satirlar: List[UcretSatiri] = []
    for table in tables:
        kategori = _find_category_title(table)
        satirlar = _extract_rows_from_table(table, kategori)
        tum_satirlar.extend(satirlar)

    return tum_satirlar if tum_satirlar else None


def _scrape_with_playwright(url: str = GARANTI_URL) -> List[UcretSatiri]:
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise ScraperError(
            "Playwright kurulu değil. 'pip install playwright' ve "
            "'playwright install chromium' komutlarını çalıştırın."
        ) from exc

    tum_satirlar: List[UcretSatiri] = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        try:
            page = browser.new_page(user_agent=HEADERS["User-Agent"])
            page.goto(url, timeout=60000, wait_until="networkidle")

            possible_selectors = [
                "button[aria-expanded='false']",
                ".accordion-header",
                ".accordion-title",
            ]
            for selector in possible_selectors:
                try:
                    elements = page.query_selector_all(selector)
                    for el in elements:
                        try:
                            el.click(timeout=2000)
                            page.wait_for_timeout(200)
                        except Exception:
                            continue
                except Exception:
                    continue

            page.wait_for_timeout(1000)
            html = page.content()
        finally:
            browser.close()

    soup = BeautifulSoup(html, "lxml")
    tables = soup.find_all("table")

    if not tables:
        raise ScraperError("Playwright ile de tablo bulunamadı. Sayfa yapısı değişmiş olabilir.")

    for table in tables:
        kategori = _find_category_title(table)
        satirlar = _extract_rows_from_table(table, kategori)
        tum_satirlar.extend(satirlar)

    return tum_satirlar


def scrape_garanti_bbva(url: str = GARANTI_URL) -> List[UcretSatiri]:
    print(f"[scraper] {url} adresinden veri çekiliyor (requests)...", file=sys.stderr)
    satirlar = _scrape_with_requests(url)

    if satirlar:
        print(f"[scraper] requests ile {len(satirlar)} satır bulundu.", file=sys.stderr)
        return satirlar

    print("[scraper] requests ile tablo bulunamadı, Playwright deneniyor...", file=sys.stderr)
    satirlar = _scrape_with_playwright(url)
    print(f"[scraper] Playwright ile {len(satirlar)} satır bulundu.", file=sys.stderr)

    if not satirlar:
        raise ScraperError(
            "Sayfadan hiçbir ücret satırı çekilemedi. "
            "Sayfa yapısı değişmiş olabilir, scraper'ın güncellenmesi gerekebilir."
        )

    return satirlar


if __name__ == "__main__":
    veriler = scrape_garanti_bbva()
    for v in veriler[:10]:
        print(v)
    print(f"Toplam {len(veriler)} satır bulundu.")
