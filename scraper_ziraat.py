"""
Ziraat Bankası "Ürün ve Hizmet Ücretleri" sayfasını çeken scraper modülü.
"""

import re
import sys
from dataclasses import dataclass
from typing import List, Optional

import requests
from bs4 import BeautifulSoup

ZIRAAT_URL = "https://www.ziraatbank.com.tr/tr/urun-ve-hizmet-ucretleri"

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


def _find_category_title(table) -> str:
    el = table.parent
    depth = 0
    while el is not None and depth < 8:
        for sibling in el.find_all_previous(["h1", "h2", "h3", "h4", "h5", "button"], limit=3):
            text = _normalize(sibling.get_text())
            if len(text) > 5 and text not in ["Müşteri Ol", "Ara", "Kapat", "Menü"]:
                return text
        el = el.parent
        depth += 1
    return "Genel"


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

    print(f"  [ziraat debug] Kategori: {kategori}, Başlık: {header_texts}", file=sys.stderr)

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


def _scrape_with_playwright(url: str = ZIRAAT_URL) -> List[UcretSatiri]:
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

    print(f"\n[ziraat] TOPLAM {len(tables)} TABLO BULUNDU", file=sys.stderr)
    for t_idx, table in enumerate(tables[:5]):
        rows = table.find_all("tr")
        print(f"  Tablo {t_idx+1} ({len(rows)} satır):", file=sys.stderr)
        for r_idx, row in enumerate(rows[:3]):
            cells = row.find_all(["th", "td"])
            vals = [f"[{c.name}]'{_normalize(c.get_text(strip=True))}'" for c in cells]
            print(f"    Satır {r_idx+1}: {vals}", file=sys.stderr)

    if not tables:
        raise ScraperError("Playwright ile tablo bulunamadı.")

    tum_satirlar: List[UcretSatiri] = []
    for table in tables:
        kategori = _find_category_title(table)
        rows = _extract_rows_from_table(table, kategori)
        tum_satirlar.extend(rows)

    return tum_satirlar


def _scrape_with_requests(url: str = ZIRAAT_URL) -> Optional[List[UcretSatiri]]:
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
        rows = _extract_rows_from_table(table, kategori)
        tum_satirlar.extend(rows)

    return tum_satirlar if tum_satirlar else None


def scrape_ziraat(url: str = ZIRAAT_URL) -> List[UcretSatiri]:
    print(f"[ziraat] {url} adresinden veri çekiliyor...", file=sys.stderr)

    print("[ziraat] Playwright deneniyor...", file=sys.stderr)
    try:
        satirlar = _scrape_with_playwright(url)
        if satirlar:
            print(f"[ziraat] Playwright ile {len(satirlar)} satır bulundu.", file=sys.stderr)
            return satirlar
    except Exception as exc:
        print(f"[ziraat] Playwright başarısız: {exc}", file=sys.stderr)

    print("[ziraat] requests ile deneniyor...", file=sys.stderr)
    satirlar = _scrape_with_requests(url)
    if satirlar:
        print(f"[ziraat] requests ile {len(satirlar)} satır bulundu.", file=sys.stderr)
        return satirlar

    raise ScraperError("Ziraat sayfasından hiçbir ücret satırı çekilemedi.")


if __name__ == "__main__":
    veriler = scrape_ziraat()
    for v in veriler[:5]:
        print(v)
    print(f"Toplam {len(veriler)} satır bulundu.")
