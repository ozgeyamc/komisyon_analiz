"""
Halkbank "Ürün ve Hizmet Ücretleri" sayfasını çeken scraper modülü.
Ana sayfadaki tüm alt sayfa linklerini otomatik tespit eder.
"""

import re
import sys
from dataclasses import dataclass
from typing import List, Optional

import requests
from bs4 import BeautifulSoup

HALKBANK_ANA_URL = "https://www.halkbank.com.tr/tr/urun-ve-hizmet-ucretleri"

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


def _find_category_title(table, sayfa_kategorisi: str) -> str:
    el = table.parent
    depth = 0
    while el is not None and depth < 8:
        for sibling in el.find_all_previous(["h1", "h2", "h3", "h4", "h5", "button"], limit=3):
            text = _normalize(sibling.get_text())
            if len(text) > 5 and text not in ["Müşteri Ol", "Ara", "Kapat", "Menü", "Ana Sayfa"]:
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
    col_tarih     = find_col(["güncelleme"])
    if col_tarih == -1:
        col_tarih = find_col(["guncelleme"])

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

        site_tarihi = get(col_tarih) if col_tarih >= 0 else ""
        temiz_aciklama, aciklama_tarihi = _parse_aciklama(get(col_aciklama))
        if not site_tarihi:
            site_tarihi = aciklama_tarihi

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


def _alt_sayfa_linklerini_bul(ana_url: str) -> List[tuple]:
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        try:
            page = browser.new_page(user_agent=HEADERS["User-Agent"])
            page.goto(ana_url, timeout=60000, wait_until="networkidle")
            page.wait_for_timeout(2000)
            html = page.content()
        finally:
            browser.close()

    soup = BeautifulSoup(html, "lxml")

    base = "https://www.halkbank.com.tr"
    prefix = "/tr/urun-ve-hizmet-ucretleri/"

    bulunan = {}
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        if href.startswith(base):
            path = href.replace(base, "")
        else:
            path = href

        if path.startswith(prefix) and len(path) > len(prefix):
            clean_path = path.split("#")[0].split("?")[0].rstrip("/")
            full_url = base + clean_path
            alt_kisim = clean_path.replace(prefix, "")
            if "/" not in alt_kisim and alt_kisim:
                link_text = _normalize(a.get_text())
                if not link_text:
                    link_text = alt_kisim.replace("-", " ").title()
                if full_url not in bulunan:
                    bulunan[full_url] = link_text

    linkler = [(ad, url) for url, ad in bulunan.items()]
    print(f"[halkbank] Ana sayfada {len(linkler)} alt sayfa bulundu:", file=sys.stderr)
    for ad, url in linkler:
        print(f"  - {ad}: {url}", file=sys.stderr)

    return linkler


def _scrape_sayfa(url: str, sayfa_kategorisi: str) -> List[UcretSatiri]:
    from playwright.sync_api import sync_playwright

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


def scrape_halkbank(ana_url: str = HALKBANK_ANA_URL) -> List[UcretSatiri]:
    print(f"[halkbank] Alt sayfalar tespit ediliyor: {ana_url}", file=sys.stderr)

    try:
        alt_sayfalar = _alt_sayfa_linklerini_bul(ana_url)
    except Exception as exc:
        print(f"[halkbank] Alt sayfa tespiti başarısız: {exc}", file=sys.stderr)
        alt_sayfalar = []

    if not alt_sayfalar:
        print(f"[halkbank] Fallback liste kullanılıyor...", file=sys.stderr)
        alt_sayfalar = [
            ("Kredi Kartları ve Banka Kartları", f"{ana_url}/kredi-kartlari-ve-banka-kartlari"),
            ("Havale EFT FAST", f"{ana_url}/havale-eft-fast"),
            ("Krediler", f"{ana_url}/krediler"),
            ("Mevduat İşlemleri", f"{ana_url}/mevduat-islemleri"),
            ("Menkul Kıymet İşlemleri", f"{ana_url}/menkul-kiymet-islemleri"),
            ("Dış Ticaret İşlemleri", f"{ana_url}/dis-ticaret-islemleri"),
            ("Çekler ve Senetler", f"{ana_url}/cekler-ve-senetler"),
            ("Sigorta İşlemleri", f"{ana_url}/sigorta-islemleri"),
            ("Diğer İşlemler", f"{ana_url}/diger-islemler"),
        ]

    tum_satirlar: List[UcretSatiri] = []
    for sayfa_adi, url in alt_sayfalar:
        print(f"[halkbank] '{sayfa_adi}' çekiliyor...", file=sys.stderr)
        try:
            satirlar = _scrape_sayfa(url, sayfa_adi)
            print(f"[halkbank] '{sayfa_adi}' — {len(satirlar)} satır.", file=sys.stderr)
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
