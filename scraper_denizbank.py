"""
DenizBank "Ürün ve Hizmet Ücretleri" sayfasını çeken scraper modülü.
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
        temiz_aciklama = DATE_PATTERN.sub("", raw_aciklama).strip(" .")
        return _normalize(temiz_aciklama), tarih

    match_tr = DATE_PATTERN_TR.search(raw_aciklama)
    if match_tr:
        gun = match_tr.group(1).zfill(2)
        ay = TURKCE_AYLAR.get(match_tr.group(2).lower(), "00")
        yil = match_tr.group(3)
        tarih = f"{gun}.{ay}.{yil}"
        temiz_aciklama = DATE_PATTERN_TR.sub("", raw_aciklama).strip(" .")
        return _normalize(temiz_aciklama), tarih

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


def _extract_from_table(table, kategori: str) -> List[UcretSatiri]:
    """
    DenizBank iki farklı tablo yapısı kullanıyor:
    Tip 1: Standart — başlık satırı + veri satırları (Masraf, Asgari Tutar, ...)
    Tip 2: Dikey — ilk satır başlık, ikinci satır+ veri, tek kolonda 'Güncelleme Tarihi' başlığı var
    """
    satirlar = []
    rows = table.find_all("tr")
    if not rows:
        return satirlar

    # Tüm satırları al
    all_rows = rows
    if not all_rows:
        return satirlar

    # İlk satır başlık mı kontrol et
    first_row_cells = all_rows[0].find_all(["th", "td"])
    header_texts = [_normalize(c.get_text(strip=True)).lower() for c in first_row_cells]

    # Tip 2 tespiti: başlık satırında "güncelleme tarihi" veya tek kolon var
    if len(header_texts) <= 2 or "güncelleme tarihi" in " ".join(header_texts) or "guncelleme tarihi" in " ".join(header_texts):
        return _extract_denizbank_tip2(table, kategori)

    # Tip 1: standart tablo
    return _extract_denizbank_tip1(table, kategori)


def _extract_denizbank_tip1(table, kategori: str) -> List[UcretSatiri]:
    """Standart tablo: Masraf | Asgari Tutar | ... | Açıklama"""
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

    col_sube = find_col(["şube"])
    if col_sube == -1:
        col_sube = find_col(["sube"])

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

        sube = get(col_sube) if col_sube >= 0 else ""
        gercek_kategori = f"{kategori} - {sube}" if sube and sube != masraf else kategori

        site_tarihi = get(col_tarih) if col_tarih >= 0 else ""
        site_tarihi = site_tarihi.replace("/", ".")
        temiz_aciklama, aciklama_tarihi = _parse_aciklama(get(col_aciklama))
        if not site_tarihi:
            site_tarihi = aciklama_tarihi

        satirlar.append(UcretSatiri(
            kategori=gercek_kategori, masraf=masraf,
            asgari_tutar=get(col_asg_tutar), asgari_oran=get(col_asg_oran),
            azami_tutar=get(col_azm_tutar), azami_oran=get(col_azm_oran),
            aciklama=temiz_aciklama, site_guncelleme_tarihi=site_tarihi,
        ))
    return satirlar


def _extract_denizbank_tip2(table, kategori: str) -> List[UcretSatiri]:
    """
    DenizBank Tip 2 tablo:
    Satır 1: başlık adı (tek hücre, tablo adı)
    Satır 2+: [Masraf adı] [değer] [değer] ... [Güncelleme Tarihi] [tarih]
    veya dikey key-value çiftleri
    """
    satirlar = []
    rows = table.find_all("tr")
    if len(rows) < 2:
        return satirlar

    # Başlık satırını atla, veri satırlarını işle
    masraf_adi = _normalize(rows[0].get_text(strip=True))
    if not masraf_adi or len(masraf_adi) < 3:
        masraf_adi = kategori

    # Tüm satırlarda key-value çiftlerini ara
    site_tarihi = ""
    for row in rows[1:]:
        cells = row.find_all(["th", "td"])
        if not cells:
            continue
        texts = [_normalize(c.get_text(strip=True)) for c in cells]

        # "Güncelleme Tarihi" - tarih çifti
        for i, t in enumerate(texts):
            if "güncelleme" in t.lower() or "guncelleme" in t.lower():
                if i + 1 < len(texts):
                    tarih_val = texts[i + 1].replace("/", ".")
                    if re.match(r"\d{2}\.\d{2}\.\d{4}", tarih_val):
                        site_tarihi = tarih_val
                elif re.match(r"\d{2}[./]\d{2}[./]\d{4}", t):
                    site_tarihi = t.replace("/", ".")

    # Tarih satırından başka bir şey çekilemiyorsa masraf_adi ile tek satır ekle
    if site_tarihi and masraf_adi != kategori:
        satirlar.append(UcretSatiri(
            kategori=kategori,
            masraf=masraf_adi,
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
