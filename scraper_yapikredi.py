"""
scraper_yapikredi_all.py

Yapı Kredi "Bireysel Ürün ve Hizmet Ücretleri" sayfasından tüm tabloları çeker.
- Tüm gizli sekmeler açılır (130+ tablo bulunur).
- Sütun eşleştirme mantığı EFT ve Havale başlıklarını (İşlem Kanalı, BSMV vb.) içerecek şekilde genişletildi.
"""

import re
import sys
from dataclasses import dataclass
from typing import List, Dict, Tuple, Set

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
    raw = (raw_aciklama or "").strip()
    match = DATE_PATTERN.search(raw)
    if match:
        tarih = match.group(1)
        temiz = DATE_PATTERN.sub("", raw).strip(" .")
        return temiz, tarih
    match_tr = DATE_PATTERN_TR.search(raw)
    if match_tr:
        gun = match_tr.group(1).zfill(2)
        ay = TURKCE_AYLAR.get(match_tr.group(2).lower(), "00")
        yil = match_tr.group(3)
        tarih = f"{gun}.{ay}.{yil}"
        temiz = DATE_PATTERN_TR.sub("", raw).strip(" .")
        return temiz, tarih
    return raw, ""


def _normalize(val: str) -> str:
    return (val or "").strip().replace("\xa0", " ").replace("\u200b", "").strip()


def _find_col_indices_from_headers(header_texts: List[str]) -> Dict[str, int]:
    def norm(h: str) -> str:
        h2 = re.sub(r"\(.*?\)", "", (h or ""))
        h2 = h2.replace("%", " ")
        h2 = re.sub(r"\s+", " ", h2).strip().lower()
        return h2

    headers_norm = [norm(h) for h in header_texts]

    def find_col(keywords: List[str]) -> int:
        for i, h in enumerate(headers_norm):
            # ARTIK KELİMELERİN BİRİ BİLE GEÇİYORSA O SÜTUNU ALIYORUZ
            if any(k in h for k in keywords):
                return i
        return -1

    # EFT tabloları için İşlem ve Kanal kelimelerini başa aldık
    col_masraf = find_col(["işlem", "kanal", "masraf", "hizmet", "ürün", "ücret", "eft", "havale", "gönder", "kategori"])
    if col_masraf == -1:
        col_masraf = 0 # Bulamazsa banko ilk sütun masraf adıdır

    col_asg_tutar = find_col(["asgari", "ücreti", "bsmv dahil", "tutar"])
    col_asg_oran = find_col(["asgari oran", "oran"])
    col_azm_tutar = find_col(["azami tutar", "azami"])
    col_azm_oran = find_col(["azami oran"])
    col_aciklama = find_col(["açıklama", "aciklama", "detay", "not"])
    col_tarih = find_col(["güncelleme", "guncelleme", "tarih"])

    return {
        "masraf": col_masraf,
        "asgari_tutar": col_asg_tutar,
        "asgari_oran": col_asg_oran,
        "azami_tutar": col_azm_tutar,
        "azami_oran": col_azm_oran,
        "aciklama": col_aciklama,
        "tarih": col_tarih,
    }


def scrape_yapikredi_all
