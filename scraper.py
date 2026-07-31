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

# Garanti sitesindeki beklenen kolon başlıkları
BEKLENEN_BASLIKLAR = ["masraf", "asgari tutar", "asgari oran", "azami tutar", "azami oran", "açıklama"]


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


def _is_header_row(cells) -> bool:
    """Bir satırın başlık satırı olup olmadığını kontrol et."""
    texts = [_normalize(c.get_text()).lower() for c in cells]
    # Eğer hücreler th ise kesinlikle başlık
    if all(c.name == "th" for c in cells):
        return True
    # td olsa bile içerik olarak başlık kelimelerini içeriyorsa başlık say
    matches = sum(1 for t in texts if any(b in t for b in BEKLENEN_BASLIKLAR))
    return matches >= 2


def _extract_rows_from_table(table, kategori: str) -> List[UcretSatiri]:
    satirlar: List[UcretSatiri] = []
    rows = table.find_all("tr")
    if not rows:
        return satirlar

    # Başlık satırını bul ve kolon indekslerini belirle
    col_masraf = 0
    col_asgari_tutar = 1
    col_asgari_oran = 2
    col_azami_tutar = 3
    col_azami_oran = 4
    col_aciklama = 5
    header_found = False
    data_start_idx = 0

    for i, row in enumerate(rows):
        cells = row.find_all(["th", "td"])
        if not cells:
            continue

        if _is_header_row(cells):
            header_found = True
            data_start_idx = i + 1
            # Kolon eşleştirmesi yap
            for j, cell in enumerate(cells):
                text = _normalize(cell.get_text()).lower()
                if "masraf" in text:
                    col_masraf = j
                elif "asgari" in text and "tutar" in text:
                    col_asgari_tutar = j
                elif "asgari" in text and "oran" in text:
                    col_asgari_oran = j
                elif "azami" in text and "tutar" in text:
                    col_azami_tutar = j
                elif "azami" in text and "oran" in text:
                    col_azami_oran = j
                elif "açıklama" in text or "aciklama" in text:
                    col_aciklama = j
            break

    if not header_found:
        data_start_idx = 0

    # Veri satırlarını oku
    for row in rows[data_start_idx:]:
        cells = row.find_all("td")
        if not cells or len(cells) < 2:
            continue

        values = [_normalize(c.get_text(strip=True)) for c in cells]

        def get(idx):
            return values[idx] if idx < len(values) else ""

        masraf = get(col_masraf)
        if not masraf:
            continue

        # Başlık satırı veri satırı olarak gelmiş olabilir, atla
        if any(b in masraf.lower() for b in BEKLENEN_BASLIKLAR):
            continue

        raw_aciklama = get(col_aciklama)
        temiz_aciklama, site_tarihi = _parse_aciklama(raw_aciklama)

        satirlar.append(
            UcretSatiri(
                kategori=kategori,
                masraf=masraf,
                asgari_tutar=get(col_asgari_tutar),
                asgari_oran=get(col_asgari_oran),
                azami_tutar=get(col_azami_tutar),
                azami_oran=get(col_azami_oran),
                aciklama=temiz_aciklama,
                site_guncelleme_tarihi=site_tarihi,
            )
        )

    return satirlar


def _find_category_title(table) -> str:
    """Tablonun ait olduğu kategori başlığını bul."""
    # Tablonun parent'larında accordion button veya heading ara
    el = table.parent
    depth = 0
    while el is not None and depth < 8:
        # Önce kardeş elementlerde ara (previous sibling)
        for sibling in el.find_all_previous(["h1", "h2", "h3", "h4", "h5", "button"], limit=3):
            text = _normalize(sibling.get_text())
            # Çok kısa veya gereksiz metinleri atla
            if len(text) > 5 and text not in ["Müşteri Ol", "Ara", "Kapat"]:
                return text
        el = el.parent
        depth += 1

    return "Genel"


def _scrape_with_playwright(url: str = GARANTI_URL) -> List[UcretSatiri]:
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise ScraperError("Playwright kurulu değil.") from exc

    tum_satirlar: List[UcretSatiri] = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        try:
            page = browser.new_page(user_agent=HEADERS["User-Agent"])
            page.goto(url, timeout=60000, wait_until="networkidle")

            # Tüm kapalı accordion'ları aç
            for selector in [
                "button[aria-expanded='false']",
                ".accordion-button.collapsed",
                "[data-bs-toggle='collapse']",
                ".card-header button",
            ]:
                try:
                    elements = page.query_selector_all(selector)
                    for el in elements:
                        try:
                            el.click(timeout=2000)
                            page.wait_for_timeout(400)
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

    if not tables:
        raise ScraperError("Playwright ile de tablo bulunamadı.")

    for table in tables:
        kategori = _find_category_title(table)
        rows = _extract_rows_from_table(table, kategori)
        tum_satirlar.extend(rows)

    return tum_satirlar


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
        rows = _extract_rows_from_table(table, kategori)
        tum_satirlar.extend(rows)

    return tum_satirlar if tum_satirlar else None


def scrape_garanti_bbva(url: str = GARANTI_URL) -> List[UcretSatiri]:
    print(f"[scraper] {url} adresinden veri çekiliyor...", file=sys.stderr)

    # Playwright ile dene (accordion'lar açılsın)
    print("[scraper] Playwright deneniyor...", file=sys.stderr)
    try:
        satirlar = _scrape_with_playwright(url)
        if satirlar:
            print(f"[scraper] Playwright ile {len(satirlar)} satır bulundu.", file=sys.stderr)
            return satirlar
    except Exception as exc:
        print(f"[scraper] Playwright başarısız: {exc}", file=sys.stderr)

    # Fallback: requests
    print("[scraper] requests ile deneniyor...", file=sys.stderr)
    satirlar = _scrape_with_requests(url)
    if satirlar:
        print(f"[scraper] requests ile {len(satirlar)} satır bulundu.", file=sys.stderr)
        return satirlar

    raise ScraperError("Sayfadan hiçbir ücret satırı çekilemedi.")


if __name__ == "__main__":
    veriler = scrape_garanti_bbva()
    for v in veriler[:10]:
        print(v)
    print(f"Toplam {len(veriler)} satır bulundu.")
