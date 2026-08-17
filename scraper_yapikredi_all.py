"""
scraper_yapikredi_all.py

Yapı Kredi "Bireysel Ürün ve Hizmet Ücretleri" sayfasından tüm tabloları çeker.
- Playwright ile sayfayı açar, cookie/dialog butonlarını denemeye tıklar,
  accordions'u açar, sayfayı kaydırır, sonra JS ile tüm tabloları toplar.
- Python tarafında header eşlemesi yapıp UcretSatiri nesneleri üretir.
- Yinelenenleri temizler ve sonucu Excel'e yazar (pandas yüklüyse).
"""

import re
import sys
from dataclasses import dataclass
from typing import List, Dict, Any, Set, Tuple

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
    """
    header_texts: list of header cell texts (original)
    returns mapping of keys to column indices (or -1 if not found)
    """
    def norm(h: str) -> str:
        h2 = re.sub(r"\(.*?\)", "", (h or ""))
        h2 = h2.replace("%", " ")
        h2 = re.sub(r"\s+", " ", h2).strip().lower()
        return h2
    headers_norm = [norm(h) for h in header_texts]

    def find_col(keywords: List[str]) -> int:
        for i, h in enumerate(headers_norm):
            if all(k in h for k in keywords):
                return i
        return -1

    col_masraf = find_col(["masraf"])
    if col_masraf == -1:
        col_masraf = find_col(["işlem"])
    if col_masraf == -1:
        col_masraf = find_col(["ücret"])
    if col_masraf == -1:
        # try single words that often are used as section titles
        for k in (["eft"], ["gönderim"], ["havale"], ["swift"], ["gelen eft"], ["gelen"]):
            idx = find_col(k)
            if idx != -1:
                col_masraf = idx
                break

    col_asg_tutar = find_col(["asgari", "tutar"]) or find_col(["asgari"]) or find_col(["tutar"])
    col_asg_oran = find_col(["asgari", "oran"]) or find_col(["asgari oran"]) or find_col(["oran"])
    col_azm_tutar = find_col(["azami", "tutar"]) or find_col(["azami"]) or find_col(["tutar"])
    col_azm_oran = find_col(["azami", "oran"]) or find_col(["azami oran"]) or find_col(["oran

