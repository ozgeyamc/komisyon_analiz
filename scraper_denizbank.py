"""
DenizBank "Ürün ve Hizmet Ücretleri" sayfasını çeken scraper modülü.

Tablo yapısı:
  Satır 1 : tek hücre — şube/kanal adı  (ör. "Şube", "İnternet-Mobil")
  Satır 2 : başlıklar  (İşlem Türü | Asgari Tutar | ... | Güncelleme Tarihi)
  Satır 3+: veri
"""

import re
import sys
from dataclasses import dataclass
from typing import List

DENIZBANK_URL = "https://www.denizbank.com/urun-ve-hizmet-ucretleri"

DATE_PATTERN = re.compile(
    r"G[üu]ncellenme\s*Tarihi\s*:\s*[\s\xa0]*(\d{2}[./]\d{2}[./]\d{4}(?:[\s\xa0]+\d{2}:\d{2})?)",
    re.IGNORECASE,
)

DATE_PATTERN_ITIBAR = re.compile(
    r"(\d{2}[./]\d{2}[./]\d{4})\s+tarihi\s+itibar",
    re.IGNORECASE,
)

DATE_PATTERN_TR = re.compile(
    r"(?:son\s+)?g[üu]ncellenme\s+tarihi\s*:?\s*[\s\xa0]*"
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


def _normalize(val: str) -> str:
    if val is None:
        return ""
    return str(val).strip().replace("\xa0", " ").replace("\u200b", "").strip()


def _parse_aciklama(raw_aciklama: str):
    raw_aciklama = raw_aciklama.strip()

    match = DATE_PATTERN.search(raw_aciklama)
    if match:
        tarih = match.group(1).replace("/", ".").strip()
        temiz = DATE_PATTERN.sub("", raw_aciklama).strip(" .")
        return _normalize(temiz), tarih

    match2 = DATE_PATTERN_ITIBAR.search(raw_aciklama)
    if match2:
        tarih = match2.group(1).replace("/", ".").strip()
        return _normalize(raw_aciklama), tarih

    match_tr = DATE_PATTERN_TR.search(raw_aciklama)
    if match_tr:
        gun = match_tr.group(1).zfill(2)
        ay = TURKCE_AYLAR.get(match_tr.group(2).lower(), "00")
        yil = match_tr.group(3)
        tarih = f"{gun}.{ay}.{yil}"
        temiz = DATE_PATTERN_TR.sub("", raw_aciklama).strip(" .")
        return _normalize(temiz), tarih

    return _normalize(raw_aciklama), ""


def _find_category_title(el, fallback: str) -> str:
    parent = el.parent
    depth = 0
    while parent is not None and depth < 15:
        baslik = parent.find(["h1", "h2", "h3", "h4", "h5"], recursive=False)
        if baslik:
            text = _normalize(baslik.get_text())
            if len(text) > 3:
                return text
        for sibling in parent.find_all_previous(["h1", "h2", "h3", "h4", "h5"], limit=3):
            text = _normalize(sibling.get_text())
            if len(text) > 5 and text not in ["Müşteri Ol", "Ara", "Kapat", "Menü", "Ana Sayfa"]:
                return text
        parent = parent.parent
        depth += 1
    return fallback


def _extract_from_table(table, ana_kategori: str) -> List[UcretSatiri]:
    """
    DenizBank tablo yapısı:
      Satır 0 : tek hücre = şube/kanal adı  → kategori zenginleştirir
      Satır 1 : başlık satırı
      Satır 2+: veri satırları
    """
    satirlar = []
    rows = table.find_all("tr")
    if len(rows) < 2:
        return satirlar

    # Satır 0 — tek hücreli ise şube adı
    ilk_cells = rows[0].find_all(["th", "td"])
    if len(ilk_cells) == 1:
        sube_adi = _normalize(ilk_cells[0].get_text(strip=True))
        baslik_rows = rows[1:]
    else:
        sube_adi = ""
        baslik_rows = rows

    if not baslik_rows:
        return satirlar

    # Başlık satırı
    header_cells = baslik_rows[0].find_all(["th", "td"])
    header_texts = [_normalize(c.get_text(strip=True)).lower() for c in header_cells]
    data_rows = baslik_rows[1:]

    # Başlık yoksa veya tek hücreliyse atla
    if len(header_texts) < 2:
        return satirlar

    def find_col(keywords):
        for i, h in enumerate(header_texts):
            if all(k in h for k in keywords):
                return i
        return -1

    col_masraf    = find_col(["işlem"])
    if col_masraf == -1:
        col_masraf = find_col(["masraf"])
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
    if col_tarih == -1:
        col_tarih = find_col(["tarih"])

    if col_masraf == -1:
        col_masraf = 0; col_asg_tutar = 1; col_asg_oran = 2
        col_azm_tutar = 3; col_azm_oran = 4; col_aciklama = 5

    # Kategori = ana kategori + şube adı
    kategori = f"{ana_kategori} - {sube_adi}" if sube_adi else ana_kategori

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

        site_tarihi = get(col_tarih).replace("/", ".") if col_tarih >= 0 else ""
        temiz_aciklama, aciklama_tarihi = _parse_aciklama(get(col_aciklama))
        if not site_tarihi:
            site_tarihi = aciklama_tarihi

        satirlar.append(UcretSatiri(
            kategori=kategori,
            masraf=masraf,
            asgari_tutar=get(col_asg_tutar),
            asgari_oran=get(col_asg_oran),
            azami_tutar=get(col_azm_tutar),
            azami_oran=get(col_azm_oran),
            aciklama=temiz_aciklama,
            site_guncelleme_tarihi=site_tarihi,
        ))

    return satirlar


def scrape_denizbank(url: str = DENIZBANK_URL) -> List[UcretSatiri]:
    from playwright.sync_api import sync_playwright
    from bs4 import BeautifulSoup

    print(f"[denizbank] {url} adresinden veri çekiliyor...", file=sys.stderr)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        try:
            page = browser.new_page(user_agent=HEADERS["User-Agent"])
            page.goto(url, timeout=90000, wait_until="domcontentloaded")
            page.wait_for_timeout(5000)

            page.evaluate("""
                () => {
                    document.querySelectorAll('.tab-pane').forEach(el => {
                        el.classList.add('active', 'show');
                        el.style.display = 'block';
                    });
                }
            """)
            page.wait_for_timeout(2000)
            html = page.content()

        finally:
            browser.close()

    soup = BeautifulSoup(html, "lxml")

    tum_satirlar = []
    for i in range(1, 10):
        bolum = soup.find(id=f"tb-{i}")
        if not bolum:
            continue
        baslik_el = bolum.find(["h1", "h2", "h3", "h4", "h5"])
        kategori = _normalize(baslik_el.get_text()) if baslik_el else f"Bölüm {i}"
        print(f"[denizbank] tb-{i} kategorisi: {kategori}", file=sys.stderr)

        tablolar = bolum.find_all("table")
        print(f"[denizbank] tb-{i} tablo sayısı: {len(tablolar)}", file=sys.stderr)
        for table in tablolar:
            rows = _extract_from_table(table, kategori)
            tum_satirlar.extend(rows)

    if not tum_satirlar:
        raise ScraperError("DenizBank sayfasında hiç veri satırı çekilemedi.")

    tum_satirlar = sorted(tum_satirlar, key=lambda s: (s.kategori, s.masraf))
    print(f"[denizbank] Toplam {len(tum_satirlar)} satır bulundu.", file=sys.stderr)
    return tum_satirlar


if __name__ == "__main__":
    veriler = scrape_denizbank()
    for v in veriler[:5]:
        print(v)
    print(f"Toplam {len(veriler)} satır bulundu.")
