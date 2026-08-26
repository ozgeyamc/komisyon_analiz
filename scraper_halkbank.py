"""
Halkbank "Ürün ve Hizmet Ücretleri" scraper.

Bu sürüm:
- Ana ücret sayfasındaki güncel alt sayfaları otomatik bulur.
- Statik sayfalardaki tüm productservicefee tablo kimliklerini keşfeder.
- Dinamik tabloları Halkbank'ın kullandığı resmî web API'sinden eksiksiz çeker.
- 148 tablo / 367 dinamik satır altına düşerse yarım veri yazmak yerine hata verir.
- 148 API tablosunu bankayı eşzamanlı istekle yormadan sıralı ve kontrollü çeker.
- Requests/TLS yolu tamamlanamazsa Chromium DOM'unu beklemek yerine Playwright'ın
  ayrı HTTP motoruyla aynı resmî API tablolarını sıralı çeken fallback kullanır.
- Ticari ücret PDF'inin dört sayfasını da okur; sonraki sayfalarda başlık satırı
  tekrarlanmasa bile kolon düzenini korur.
- PDF bölüm başlıklarını ücret satırı yapmaz, alt hizmetleri hiyerarşik adlarıyla
  korur ve satır taşmalarını ücret kolonlarına sızmadan birleştirir.
- Güncel Halkbank sayfa başlıklarını / alt başlıklarını MASRAF alanında korur.
- Kart ürün/tablo adlarını MASRAF alanında korur; aynı tutarlı farklı kartları birleştirmez.
- Havale / EFT / FAST başlıklarını kaybetmez.
- "Uluslararası Fon Transferi Ücreti" satırlarını SWIFT filtresinde görünür yapar.
- Duplicate kayıtları temizler.
- Sayfa, tablo, kategori, bütünlük ve para aktarma raporları üretir.
- main.py / update_excel.py ile uyumlu UcretSatiri listesi döndürür.
"""

import io
import re
import sys
import threading
import time
import unicodedata
from dataclasses import dataclass
from typing import Dict, List, Optional, Set, Tuple
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup


SCRAPER_VERSION = "2026-08-26-v8-halkbank-targeted-api-recovery"

HALKBANK_ANA_URL = "https://www.halkbank.com.tr/tr/urun-ve-hizmet-ucretleri"
HALKBANK_API_URL = "https://webapi.halkbank.com.tr/api/productservicefee/{table_id}"

# 25.08.2026 tarihli resmî sayfa envanteri. Bunlar ücret tutarı değil,
# dinamik Halkbank tablolarının eksiksiz yüklenip yüklenmediğini anlamak için
# kullanılan asgari bütünlük eşikleridir. Yeni tablo/satırlar engellenmez;
# yalnız eksik yükleme Excel yazımını durdurur.
MIN_API_TABLE_COUNT = 148
# 25.08.2026'da beş MKK tablosu resmî API'de mevcut tutulup feeItems=[]
# yayımlanmaya başladı. Kalan tablolarda 356 gerçek ücret satırı vardır;
# boş beş hizmet ayrı durum satırı olarak korunur.
MIN_API_ITEM_COUNT = 356
MIN_API_OUTPUT_ROW_COUNT = 361
ALLOWED_EMPTY_API_TABLE_IDS = {
    "147",  # MKK Hesap Açma Ücreti
    "148",  # MKK Şifre Gönderim Ücreti
    "149",  # MKK Hesap Bakım Ücretleri
    "150",  # MKK Menkul Kıymet Transferi Ücretleri
    "152",  # MKK Saklama Ücretleri
}
API_MAX_ATTEMPTS = 4
API_TIMEOUT_SECONDS = 35
API_SUCCESS_DELAY_SECONDS = 0.20

# 24.08.2026 tarihli güncel resmî ticari PDF envanteri. PDF dört sayfada
# 199 numaralı satır ve en az 116 ücret/açıklama içeren gerçek hizmet satırı
# yayımlıyor. Başlık satırları bu sayılara dahil edilmez.
MIN_COMMERCIAL_PDF_PAGE_COUNT = 4
MIN_COMMERCIAL_PDF_NUMBERED_ROWS = 190
MIN_COMMERCIAL_PDF_OUTPUT_ROWS = 116

# 361 API çıktı satırı + 135 doğrulanmış ticari PDF hizmeti.
MIN_FINAL_ROW_COUNT = MIN_API_OUTPUT_ROW_COUNT + 135

REQUIRED_API_TERMS = (
    "mkk hesap bakim ucretleri",
    "mkk hesap acma ucreti",
    "mkk sifre gonderim ucreti",
    "mkk menkul kiymet transferi ucretleri",
    "mkk saklama ucretleri",
)

REQUIRED_COMMERCIAL_TERMS = (
    "nakit yonetimi",
    "uluslararasi fon transferi",
    "fast islemleri",
    "belge ve bilgilendirme",
    "cek defteri",
    "uye isyeri",
)

REQUIRED_FINAL_TERMS = REQUIRED_API_TERMS + REQUIRED_COMMERCIAL_TERMS + (
    "paraf klasik",
    "paraf gold",
    "paraf platinum",
)

_API_THREAD_LOCAL = threading.local()

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "tr-TR,tr;q=0.9,en;q=0.8",
}

# Güncel ana sayfadaki 9 ana bölüm.
FALLBACK_PAGES = [
    (
        "Kredi Kartları ve Banka Kartları",
        f"{HALKBANK_ANA_URL}/kredi-kartlari-ve-banka-kartlari",
    ),
    (
        "Para ve Kıymetli Maden Transferleri",
        f"{HALKBANK_ANA_URL}/para-ve-kiymetli-maden-transferleri",
    ),
    (
        "Bireysel Krediler",
        f"{HALKBANK_ANA_URL}/bireysel-krediler",
    ),
    (
        "Mevduat/Katılım Fonu",
        f"{HALKBANK_ANA_URL}/mevduat-katilim-fonu",
    ),
    (
        "Menkul Kıymet İşlemleri",
        f"{HALKBANK_ANA_URL}/menkul-kiymet-islemleri",
    ),
    (
        "Çekler ve Senetler",
        f"{HALKBANK_ANA_URL}/cekler-ve-senetler",
    ),
    (
        "KOBİ Kredileri",
        f"{HALKBANK_ANA_URL}/kobi-kredileri",
    ),
    (
        "Diğer",
        f"{HALKBANK_ANA_URL}/diger",
    ),
    (
        "Ticari Ücret ve Komisyonları",
        f"{HALKBANK_ANA_URL}/ticari-ucret-ve-komisyonlari",
    ),
]


# =========================================================
# TARİH
# =========================================================

DATE_PATTERN = re.compile(
    r"G[üu]ncellenme\s*Tarihi\s*:\s*[\s\xa0]*"
    r"(\d{1,2}[./]\d{1,2}[./]\d{4}(?:[\s\xa0]+\d{1,2}:\d{2})?)",
    re.IGNORECASE,
)

DATE_PATTERN_ITIBAR = re.compile(
    r"(\d{1,2}[./]\d{1,2}[./]\d{4})\s+tarihi\s+itibar",
    re.IGNORECASE,
)

DATE_PATTERN_TR = re.compile(
    r"(?:son\s+)?g[üu]ncellenme\s+tarihi\s*:?\s*[\s\xa0]*"
    r"(\d{1,2})\s+"
    r"(ocak|şubat|mart|nisan|mayıs|haziran|temmuz|ağustos|eylül|ekim|kasım|aralık)\s+"
    r"(\d{4})",
    re.IGNORECASE,
)

TURKCE_AYLAR = {
    "ocak": "01",
    "şubat": "02",
    "mart": "03",
    "nisan": "04",
    "mayıs": "05",
    "haziran": "06",
    "temmuz": "07",
    "ağustos": "08",
    "eylül": "09",
    "ekim": "10",
    "kasım": "11",
    "aralık": "12",
}


# =========================================================
# VERİ YAPISI
# =========================================================

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


# =========================================================
# NORMALİZASYON
# =========================================================

def _normalize(value: Optional[str]) -> str:
    if value is None:
        return ""

    text = str(value)
    text = text.replace("\xa0", " ")
    text = text.replace("\u200b", "")
    text = text.replace("\r", " ")
    text = text.replace("\n", " ")
    text = re.sub(r"\s+", " ", text)

    return text.strip()


def _normalize_key(value: Optional[str]) -> str:
    text = _normalize(value).lower()

    replacements = {
        "ı": "i",
        "ğ": "g",
        "ü": "u",
        "ş": "s",
        "ö": "o",
        "ç": "c",
    }

    for old, new in replacements.items():
        text = text.replace(old, new)

    # Python'da büyük Türkçe İ harfinin lower() sonucu "i + combining dot"
    # olabilir. Arama/dedup anahtarında görünmez birleşik karakter bırakma.
    text = unicodedata.normalize("NFKD", text)
    text = "".join(
        char
        for char in text
        if not unicodedata.combining(char)
    )

    text = text.replace("%", " ")
    text = re.sub(r"\s+", " ", text)

    return text.strip()


def _same_text(a: str, b: str) -> bool:
    return _normalize_key(a) == _normalize_key(b)


def _requests_get_with_retry(
    requester,
    url: str,
    *,
    headers: Optional[dict] = None,
    timeout: int = 40,
    label: str,
):
    """HTML/PDF isteklerinde geçici GitHub-runner ağ hatalarını tolere eder."""
    last_error: Optional[Exception] = None

    for attempt in range(1, API_MAX_ATTEMPTS + 1):
        try:
            response = requester.get(
                url,
                headers=headers,
                timeout=timeout,
            )
            response.raise_for_status()
            return response
        except Exception as exc:
            last_error = exc
            if attempt < API_MAX_ATTEMPTS:
                wait_seconds = 2 ** (attempt - 1)
                print(
                    f"[halkbank][retry] {label} | "
                    f"deneme={attempt}/{API_MAX_ATTEMPTS} | "
                    f"bekleme={wait_seconds}s | hata={exc}",
                    file=sys.stderr,
                )
                time.sleep(wait_seconds)

    raise ScraperError(
        f"Halkbank {label} alınamadı: {url} | hata={last_error}"
    )


def _parse_aciklama(raw_aciklama: str) -> Tuple[str, str]:
    raw = _normalize(raw_aciklama)

    if not raw:
        return "", ""

    match = DATE_PATTERN.search(raw)
    if match:
        tarih = match.group(1).replace("/", ".").strip()
        temiz = _normalize(
            DATE_PATTERN.sub("", raw)
        ).strip(" .:-")
        return temiz, tarih

    match_itibar = DATE_PATTERN_ITIBAR.search(raw)
    if match_itibar:
        tarih = match_itibar.group(1).replace("/", ".").strip()
        return raw, tarih

    match_tr = DATE_PATTERN_TR.search(raw)
    if match_tr:
        gun = match_tr.group(1).zfill(2)
        ay = TURKCE_AYLAR.get(
            match_tr.group(2).lower(),
            "",
        )
        yil = match_tr.group(3)

        if ay:
            tarih = f"{gun}.{ay}.{yil}"
            temiz = _normalize(
                DATE_PATTERN_TR.sub("", raw)
            ).strip(" .:-")
            return temiz, tarih

    return raw, ""


# =========================================================
# TRANSFER TERİMLERİ
# =========================================================

TRANSFER_TERMS = (
    "fast",
    "eft",
    "havale",
    "swift",
)


def _has_transfer_term(text: str, term: str) -> bool:
    normalized = _normalize_key(text)

    patterns = {
        "fast": r"(?<![a-z0-9])fast(?![a-z0-9])",
        "eft": r"(?<![a-z0-9])eft(?![a-z0-9])",
        "havale": r"(?<![a-z0-9])havale(?![a-z0-9])",
        "swift": r"(?<![a-z0-9])swift(?![a-z0-9])",
    }

    pattern = patterns.get(term)

    if pattern:
        return re.search(pattern, normalized) is not None

    return term in normalized


# =========================================================
# ALT SAYFA KEŞFİ
# =========================================================

INVALID_LINK_TEXTS = {
    "detaylı bilgi",
    "detayli bilgi",
    "ürün ve hizmet ücretleri",
    "urun ve hizmet ucretleri",
}


def _page_name_from_slug(url: str) -> str:
    slug = urlparse(url).path.rstrip("/").split("/")[-1]

    special = {
        "kredi-kartlari-ve-banka-kartlari": "Kredi Kartları ve Banka Kartları",
        "para-ve-kiymetli-maden-transferleri": "Para ve Kıymetli Maden Transferleri",
        "bireysel-krediler": "Bireysel Krediler",
        "mevduat-katilim-fonu": "Mevduat/Katılım Fonu",
        "menkul-kiymet-islemleri": "Menkul Kıymet İşlemleri",
        "cekler-ve-senetler": "Çekler ve Senetler",
        "kobi-kredileri": "KOBİ Kredileri",
        "diger": "Diğer",
        "ticari-ucret-ve-komisyonlari": "Ticari Ücret ve Komisyonları",
    }

    if slug in special:
        return special[slug]

    return slug.replace("-", " ").title()


def _discover_pages_requests(
    ana_url: str,
    session: requests.Session,
) -> List[Tuple[str, str]]:
    response = _requests_get_with_retry(
        session,
        ana_url,
        headers=HEADERS,
        timeout=40,
        label="ana ücret sayfası",
    )

    soup = BeautifulSoup(
        response.text,
        "lxml",
    )

    prefix_path = (
        urlparse(ana_url).path.rstrip("/")
        + "/"
    )

    found: Dict[str, str] = {}

    for a in soup.find_all("a", href=True):
        href = _normalize(
            a.get("href")
        )

        if not href:
            continue

        full_url = urljoin(
            ana_url + "/",
            href,
        )

        parsed = urlparse(full_url)

        if parsed.netloc not in {
            "www.halkbank.com.tr",
            "halkbank.com.tr",
        }:
            continue

        clean_path = parsed.path.rstrip("/")

        if not clean_path.startswith(
            prefix_path
        ):
            continue

        rest = clean_path[
            len(prefix_path):
        ]

        # Sadece ana ücret sayfasının doğrudan alt sayfaları.
        if not rest or "/" in rest:
            continue

        canonical = (
            "https://www.halkbank.com.tr"
            + clean_path
        )

        text = _normalize(
            a.get_text(
                " ",
                strip=True,
            )
        )

        if (
            not text
            or _normalize_key(text)
            in {
                _normalize_key(x)
                for x in INVALID_LINK_TEXTS
            }
        ):
            text = _page_name_from_slug(
                canonical
            )

        found[canonical] = text

    pages = [
        (name, url)
        for url, name in found.items()
    ]

    # Ana sayfadaki kartların doğal sırasını korumak zor olabilir.
    # Bilinen güncel sırayı öne al, yeni bir sayfa çıkarsa sona ekle.
    order = {
        url: i
        for i, (_, url)
        in enumerate(FALLBACK_PAGES)
    }

    pages.sort(
        key=lambda item: (
            order.get(item[1], 999),
            item[0],
        )
    )

    return pages


def _discover_pages(
    ana_url: str,
    session: requests.Session,
) -> List[Tuple[str, str]]:
    try:
        pages = _discover_pages_requests(
            ana_url,
            session,
        )

        if pages:
            print(
                f"[halkbank] Ana sayfada "
                f"{len(pages)} alt sayfa bulundu.",
                file=sys.stderr,
            )

            for name, url in pages:
                print(
                    f"    - {name}: {url}",
                    file=sys.stderr,
                )

            return pages

    except Exception as exc:
        print(
            f"[halkbank][UYARI] "
            f"Alt sayfa keşfi başarısız: {exc}",
            file=sys.stderr,
        )

    print(
        "[halkbank][UYARI] "
        "Güncel fallback alt sayfa listesi kullanılıyor.",
        file=sys.stderr,
    )

    return list(FALLBACK_PAGES)


# =========================================================
# TİCARİ ÜCRET PDF
# =========================================================

def _extract_commercial_pdf_url(
    html: str,
    base_url: str,
) -> str:
    soup = BeautifulSoup(
        html,
        "lxml",
    )

    # Önce açık PDF linkleri.
    for a in soup.find_all(
        "a",
        href=True,
    ):
        href = _normalize(
            a.get("href")
        )

        if ".pdf" not in href.lower():
            continue

        text = _normalize_key(
            a.get_text(
                " ",
                strip=True,
            )
        )

        href_key = _normalize_key(
            href
        )

        if (
            "ticari" in text
            or "komisyon" in text
            or "ticari" in href_key
            or "komisyon" in href_key
        ):
            return urljoin(
                base_url,
                href,
            )

    return ""


def _pdf_clean_cell(
    value: Optional[str],
) -> str:
    return _normalize(
        value or ""
    )


def _numeric_prefix(
    text: str,
) -> Tuple[str, str]:
    """
    '3.3.1. Elektronik Fon Transferi...' ->
    ('3.3.1', 'Elektronik Fon Transferi...')
    """
    text = _normalize(text)

    match = re.match(
        r"^\s*(\d+(?:\.\d+)*)\.?\s+(.*)$",
        text,
    )

    if not match:
        return "", text

    return (
        match.group(1),
        _normalize(
            match.group(2)
        ),
    )


def _commercial_pdf_rows(
    pdf_bytes: bytes,
) -> List[UcretSatiri]:
    """
    Halkbank'ın resmi 'Ticari Ücret ve Komisyonlar Tarifesi' PDF'ini
    tablo olarak okur.

    Not:
    requirements.txt içinde pdfplumber bulunmalıdır.
    """
    try:
        import pdfplumber
    except ImportError as exc:
        raise ScraperError(
            "Ticari Halkbank PDF'i için pdfplumber gerekli. "
            "requirements.txt dosyasına 'pdfplumber>=0.11.0' ekleyin."
        ) from exc

    raw_records: List[Dict[str, str]] = []
    page_count = 0
    canonical_columns: Optional[Dict[str, int]] = None

    def _find_columns(header_row: List[str]) -> Optional[Dict[str, int]]:
        header = [_normalize_key(cell) for cell in header_row]

        def find_col(*needles: str) -> int:
            for idx, cell in enumerate(header):
                if all(needle in cell for needle in needles):
                    return idx
            return -1

        name = find_col("kalem", "adi")
        min_amount = find_col("asgari", "tutar")
        max_amount = find_col("azami", "tutar")

        if min(name, min_amount, max_amount) < 0:
            return None

        date = find_col("guncelle")
        if date == -1:
            date = find_col("tarih")
        if date == -1 and len(header) >= 8:
            date = len(header) - 1

        # Halkbank PDF'inde kod kolonu başlıksız ilk kolondur.
        code = 0 if name > 0 else -1
        return {
            "code": code,
            "name": name,
            "currency": find_col("para", "birimi"),
            "min_amount": min_amount,
            "min_rate": find_col("asgari", "oran"),
            "max_amount": max_amount,
            "max_rate": find_col("azami", "oran"),
            "desc": find_col("aciklama"),
            "date": date,
        }

    def _clean_code(value: str) -> str:
        value = _normalize(value).strip(" .")
        return value if re.fullmatch(r"\d+(?:\.\d+)*", value) else ""

    def _clean_numeric_cell(value: str) -> str:
        value = _normalize(value)
        if not value:
            return ""

        # PDF metin katmanı bazen binlik ayıracını rakamdan ayırıyor:
        # "1 .960" -> "1.960", "3 6.000" -> "36.000".
        if re.fullmatch(r"[\d\s.,%+-]+", value):
            value = re.sub(r"\s*([.,])\s*", r"\1", value)
            value = re.sub(r"(?<=\d)\s+(?=\d)", "", value)
        return value

    def _is_overflow_fragment(value: str) -> bool:
        if not value:
            return False
        key = _normalize_key(value)
        return bool(key) and not re.search(r"\d", key) and len(key) <= 16

    def _has_fee_payload(record: Dict[str, str]) -> bool:
        return any(
            record[field]
            for field in (
                "min_amount",
                "min_rate",
                "max_amount",
                "max_rate",
                "desc",
            )
        )

    def _with_currency(currency: str, amount: str) -> str:
        if not currency or not amount:
            return amount
        currency = _normalize(currency).upper()
        amount = _normalize(amount)
        suffixes = [currency]
        if currency == "TRY":
            suffixes.append("TL")
        for suffix in suffixes:
            amount = re.sub(
                rf"\s+{re.escape(suffix)}$",
                "",
                amount,
                flags=re.IGNORECASE,
            )
        if amount.upper().startswith(currency):
            return amount
        return f"{currency} {amount}"

    with pdfplumber.open(
        io.BytesIO(pdf_bytes)
    ) as pdf:
        page_count = len(pdf.pages)
        for page_no, page in enumerate(
            pdf.pages,
            start=1,
        ):
            tables = page.extract_tables(
                table_settings={
                    "vertical_strategy": "lines",
                    "horizontal_strategy": "lines",
                    "intersection_tolerance": 5,
                    "snap_tolerance": 3,
                    "join_tolerance": 3,
                }
            )

            # Bazı PDF sürümlerinde çizgi algısı zayıf olabilir.
            if not tables:
                tables = page.extract_tables()

            for table in tables:
                if not table:
                    continue

                header_index = -1
                columns: Optional[Dict[str, int]] = None

                for i, row in enumerate(
                    table[:8]
                ):
                    cleaned = [_pdf_clean_cell(x) for x in row]
                    detected = _find_columns(cleaned)
                    if detected is not None:
                        header_index = i
                        columns = detected
                        canonical_columns = detected
                        break

                # Sonraki sayfalarda kolon başlığı tekrarlanmıyor.
                # İlk sayfada doğrulanan düzeni aynen kullan.
                if columns is None:
                    columns = canonical_columns
                if columns is None:
                    continue

                data_rows = table[header_index + 1:] if header_index >= 0 else table

                for raw_row in data_rows:
                    row = [
                        _pdf_clean_cell(x)
                        for x in raw_row
                    ]

                    if not any(row):
                        continue

                    def get(
                        index: int,
                    ) -> str:
                        if (
                            index < 0
                            or index >= len(row)
                        ):
                            return ""
                        return row[index]

                    raw_code = get(columns["code"])
                    raw_name = get(columns["name"])

                    if not raw_name:
                        continue

                    number = _clean_code(raw_code)
                    prefixed_number, name = _numeric_prefix(raw_name)
                    if not number:
                        number = prefixed_number

                    if not name:
                        name = raw_name

                    record = {
                        "number": number,
                        "name": _normalize(name),
                        "currency": get(columns["currency"]),
                        "min_amount": _clean_numeric_cell(get(columns["min_amount"])),
                        "min_rate": _clean_numeric_cell(get(columns["min_rate"])),
                        "max_amount": _clean_numeric_cell(get(columns["max_amount"])),
                        "max_rate": _clean_numeric_cell(get(columns["max_rate"])),
                        "desc": get(columns["desc"]),
                        "date": get(columns["date"]).replace("/", "."),
                        "page": str(page_no),
                    }

                    # Dört uzun kredi adında kapanış parantezi asgari tutar
                    # kolonuna taşıyor. Bu metin ücret değildir.
                    if (
                        _is_overflow_fragment(record["min_amount"])
                        and not any(
                            record[field]
                            for field in (
                                "currency",
                                "min_rate",
                                "max_amount",
                                "max_rate",
                                "desc",
                                "date",
                            )
                        )
                    ):
                        record["name"] = _normalize(
                            f"{record['name']}{record['min_amount']}"
                        )
                        record["min_amount"] = ""

                    # Sayfa başlığı / kolon başlığı gibi numarasız
                    # metinler ticari tarife satırı değildir.
                    if record["number"]:
                        raw_records.append(record)

    all_codes = {record["number"] for record in raw_records}
    if page_count < MIN_COMMERCIAL_PDF_PAGE_COUNT:
        raise ScraperError(
            "Halkbank ticari PDF sayfa envanteri eksik: "
            f"bulunan={page_count} | beklenen_en_az={MIN_COMMERCIAL_PDF_PAGE_COUNT}"
        )
    if len(raw_records) < MIN_COMMERCIAL_PDF_NUMBERED_ROWS:
        raise ScraperError(
            "Halkbank ticari PDF numaralı satır envanteri eksik: "
            f"bulunan={len(raw_records)} | "
            f"beklenen_en_az={MIN_COMMERCIAL_PDF_NUMBERED_ROWS}"
        )

    result: List[UcretSatiri] = []
    headings: Dict[str, str] = {}
    skipped_headings = 0
    unspecified_rows = 0

    for record in raw_records:
        number = record["number"]
        name = record["name"]
        has_descendant = any(code.startswith(number + ".") for code in all_codes)
        has_payload = _has_fee_payload(record)

        # Alt satırı bulunan ve kendi ücret/açıklaması olmayan satır başlıktır.
        # Yalnız tarih veya para birimi bulunması onu ücret satırı yapmaz.
        if has_descendant and not has_payload:
            headings[number] = name
            skipped_headings += 1
            continue

        parts = number.split(".")
        parent_codes = [".".join(parts[:i]) for i in range(1, len(parts))]
        parents = [headings[code] for code in parent_codes if code in headings]

        min_amount = record["min_amount"]
        max_amount = record["max_amount"]
        currency = record["currency"]

        min_amount = _with_currency(currency, min_amount)
        max_amount = _with_currency(currency, max_amount)

        masraf = _normalize(" - ".join([part for part in parents + [name] if part]))

        # Ticari PDF'teki uluslararası transfer hiyerarşisi SWIFT filtresinde
        # de görünür kalsın.
        combined = _normalize_key(" ".join(parents + [name]))
        if "uluslararasi fon transfer" in combined and not _has_transfer_term(masraf, "swift"):
            masraf = f"SWIFT - {masraf}"

        desc = record["desc"]
        if not has_payload:
            desc = "Ücret tutarı belirtilmemiş."
            unspecified_rows += 1

        result.append(
            UcretSatiri(
                kategori="Ticari Ücret ve Komisyonları",
                masraf=masraf,
                asgari_tutar=min_amount,
                asgari_oran=record["min_rate"],
                azami_tutar=max_amount,
                azami_oran=record["max_rate"],
                aciklama=desc,
                site_guncelleme_tarihi=record["date"],
            )
        )

    if len(result) < MIN_COMMERCIAL_PDF_OUTPUT_ROWS:
        raise ScraperError(
            "Halkbank ticari PDF hizmet satırları eksik: "
            f"bulunan={len(result)} | beklenen_en_az={MIN_COMMERCIAL_PDF_OUTPUT_ROWS}"
        )

    searchable = "\n".join(
        _normalize_key(f"{row.kategori} {row.masraf}") for row in result
    )
    missing_terms = [term for term in REQUIRED_COMMERCIAL_TERMS if term not in searchable]
    if missing_terms:
        raise ScraperError(
            "Halkbank ticari PDF kritik bölümleri eksik: "
            + ", ".join(missing_terms)
        )

    print(
        "[halkbank][ticari-pdf] "
        f"sayfa={page_count} | numaralı={len(raw_records)} | "
        f"başlık={skipped_headings} | hizmet={len(result)} | "
        f"tutarı_belirtilmemiş={unspecified_rows}",
        file=sys.stderr,
    )

    return result


def _scrape_commercial_pdf(
    page_html: str,
    page_url: str,
    session: Optional[
        requests.Session
    ] = None,
) -> List[UcretSatiri]:
    pdf_url = _extract_commercial_pdf_url(
        page_html,
        page_url,
    )

    if not pdf_url:
        raise ScraperError(
            "Ticari ücret PDF linki bulunamadı."
        )

    requester = (
        session
        if session is not None
        else requests
    )

    response = _requests_get_with_retry(
        requester,
        pdf_url,
        headers=HEADERS,
        timeout=60,
        label="ticari ücret PDF'i",
    )

    print(
        f"[halkbank] Ticari PDF bulundu: "
        f"{pdf_url}",
        file=sys.stderr,
    )

    rows = _commercial_pdf_rows(
        response.content
    )

    print(
        f"[halkbank] Ticari PDF: "
        f"{len(rows)} satır parse edildi.",
        file=sys.stderr,
    )

    return rows


# =========================================================
# TABLO NORMALİZASYONU
# =========================================================

def _root_tables(soup: BeautifulSoup):
    return [
        table
        for table in soup.find_all("table")
        if table.find_parent("table") is None
    ]


def _table_to_rows(
    table,
) -> List[List[str]]:
    """
    rowspan / colspan içeren ücret tablolarını düz bir grid'e çevirir.
    """
    rows: List[List[str]] = []
    active_spans: Dict[
        int,
        Tuple[str, int],
    ] = {}

    trs = [
        tr
        for tr in table.find_all("tr")
        if tr.find_parent("table") is table
    ]

    for tr in trs:
        cells = tr.find_all(
            ["th", "td"],
            recursive=False,
        )

        if not cells and not active_spans:
            continue

        row: List[str] = []
        col = 0
        cell_index = 0

        guard = max(
            len(cells) * 8
            + len(active_spans)
            + 20,
            30,
        )

        while (
            cell_index < len(cells)
            or active_spans
        ) and col < guard:

            if col in active_spans:
                text, remaining = (
                    active_spans[col]
                )

                row.append(text)

                if remaining <= 1:
                    del active_spans[col]
                else:
                    active_spans[col] = (
                        text,
                        remaining - 1,
                    )

                col += 1
                continue

            if cell_index >= len(cells):
                future_cols = [
                    c
                    for c in active_spans
                    if c > col
                ]

                if not future_cols:
                    break

                next_col = min(future_cols)

                while col < next_col:
                    row.append("")
                    col += 1

                continue

            cell = cells[cell_index]
            cell_index += 1

            text = _normalize(
                cell.get_text(
                    " ",
                    strip=True,
                )
            )

            try:
                rowspan = max(
                    int(
                        cell.get(
                            "rowspan",
                            1,
                        )
                    ),
                    1,
                )
            except Exception:
                rowspan = 1

            try:
                colspan = max(
                    int(
                        cell.get(
                            "colspan",
                            1,
                        )
                    ),
                    1,
                )
            except Exception:
                colspan = 1

            for _ in range(colspan):
                row.append(text)

                if rowspan > 1:
                    active_spans[col] = (
                        text,
                        rowspan - 1,
                    )

                col += 1

        if any(_normalize(x) for x in row):
            rows.append(row)

    return rows


# =========================================================
# HEADER
# =========================================================

def _header_score(row: List[str]) -> int:
    headers = [
        _normalize_key(x)
        for x in row
    ]

    tests = [
        ["masraf"],
        ["islem", "turu"],
        ["ucret"],
        ["urun"],
        ["asgari", "tutar"],
        ["asgari", "oran"],
        ["azami", "tutar"],
        ["azami", "oran"],
        ["aciklama"],
        ["guncelleme"],
        ["guncellenme"],
    ]

    score = 0

    for keywords in tests:
        if any(
            all(
                keyword in header
                for keyword in keywords
            )
            for header in headers
        ):
            score += 1

    return score


def _find_header_index(
    rows: List[List[str]],
) -> int:
    best_index = -1
    best_score = 0

    for index, row in enumerate(
        rows[:10]
    ):
        score = _header_score(row)

        if score > best_score:
            best_score = score
            best_index = index

    if best_score < 3:
        return -1

    return best_index


def _find_col(
    headers: List[str],
    keywords: List[str],
) -> int:
    for i, header in enumerate(headers):
        if all(
            keyword in header
            for keyword in keywords
        ):
            return i

    return -1


def _find_first_col(
    headers: List[str],
) -> int:
    candidates = [
        ["masraf"],
        ["islem", "turu"],
        ["islem"],
        ["urun", "hizmet"],
        ["urun"],
        ["ucret"],
    ]

    for keywords in candidates:
        index = _find_col(
            headers,
            keywords,
        )

        if index != -1:
            return index

    return 0


def _find_columns(
    header_row: List[str],
) -> Dict[str, int]:
    headers = [
        _normalize_key(x)
        for x in header_row
    ]

    result = {
        "masraf": _find_first_col(
            headers
        ),
        "asgari_tutar": _find_col(
            headers,
            ["asgari", "tutar"],
        ),
        "asgari_oran": _find_col(
            headers,
            ["asgari", "oran"],
        ),
        "azami_tutar": _find_col(
            headers,
            ["azami", "tutar"],
        ),
        "azami_oran": _find_col(
            headers,
            ["azami", "oran"],
        ),
        "aciklama": _find_col(
            headers,
            ["aciklama"],
        ),
        "tarih": _find_col(
            headers,
            ["guncelleme"],
        ),
    }

    if result["tarih"] == -1:
        result["tarih"] = _find_col(
            headers,
            ["guncellenme"],
        )

    if result["tarih"] == -1:
        result["tarih"] = _find_col(
            headers,
            ["tarih"],
        )

    # Yaygın 7 kolonlu format.
    if len(header_row) >= 7:
        fallback = {
            "masraf": 0,
            "asgari_tutar": 1,
            "asgari_oran": 2,
            "azami_tutar": 3,
            "azami_oran": 4,
            "aciklama": 5,
            "tarih": 6,
        }

        for key, index in fallback.items():
            if result[key] == -1:
                result[key] = index

    # 6 kolonlu formatta tarih açıklamadan gelebilir.
    elif len(header_row) >= 6:
        fallback = {
            "masraf": 0,
            "asgari_tutar": 1,
            "asgari_oran": 2,
            "azami_tutar": 3,
            "azami_oran": 4,
            "aciklama": 5,
        }

        for key, index in fallback.items():
            if result[key] == -1:
                result[key] = index

    return result


def _row_is_same_header(
    row: List[str],
    header: List[str],
) -> bool:
    left = [
        _normalize_key(x)
        for x in row
    ]
    right = [
        _normalize_key(x)
        for x in header
    ]

    while left and not left[-1]:
        left.pop()

    while right and not right[-1]:
        right.pop()

    return left == right


# =========================================================
# TABLO BAŞLIĞI / MASRAF
# =========================================================

INVALID_TITLES = {
    "müşteri ol",
    "ara",
    "kapat",
    "menü",
    "ana sayfa",
    "ürün ve hizmet ücretleri",
    "urun ve hizmet ucretleri",
    "masraf",
    "işlem türü",
    "islem turu",
    "açıklama",
    "aciklama",
    "güncelleme tarihi",
    "guncelleme tarihi",
}


def _is_valid_title(text: str) -> bool:
    text = _normalize(text)

    if not text:
        return False

    if len(text) < 2 or len(text) > 220:
        return False

    return _normalize_key(text) not in {
        _normalize_key(x)
        for x in INVALID_TITLES
    }


def _find_table_title(
    table,
    page_category: str,
) -> str:
    """
    Halkbank ücret alt sayfalarında H3'ler ürün/işlem gruplarını temsil ediyor:
      Havale
      EFT
      Uluslararası Fon Transferi Ücreti
      FAST
      HGS
      ...
    """
    section_title = ""

    for heading in table.find_all_previous(
        ["h3", "h4", "h5", "h6"],
        limit=20,
    ):
        text = _normalize(
            heading.get_text(
                " ",
                strip=True,
            )
        )

        if not _is_valid_title(text):
            continue

        if _same_text(
            text,
            page_category,
        ):
            continue

        section_title = text
        break

    # Dinamik bileşen ürün adını caption içinde yayımlar. Sadece H3
    # kullanılırsa aynı tutarlı farklı Paraf ürünleri dedup'ta birleşir.
    caption = table.find("caption")
    table_name = ""

    if caption is not None:
        table_name = _normalize(caption.get_text(" ", strip=True))
        if not _is_valid_title(table_name):
            table_name = ""

    return _combine_api_table_title(section_title, table_name)


def _build_masraf(
    raw_masraf: str,
    table_title: str,
) -> str:
    raw_masraf = _normalize(
        raw_masraf
    )
    table_title = _normalize(
        table_title
    )

    masraf = raw_masraf

    if table_title:
        title_key = _normalize_key(
            table_title
        )
        raw_key = _normalize_key(
            raw_masraf
        )

        # H3 ürün/işlem başlığını koru.
        if (
            title_key != raw_key
            and title_key not in raw_key
        ):
            masraf = _normalize(
                f"{table_title} - {raw_masraf}"
            )

    combined_key = _normalize_key(
        " ".join(
            [
                table_title,
                raw_masraf,
            ]
        )
    )

    # Halkbank ücret sayfasında SWIFT kalemi
    # "Uluslararası Fon Transferi Ücreti" başlığıyla yayınlanıyor.
    if (
        (
            "uluslararasi fon transfer" in combined_key
            or "yurtdisi doviz transfer" in combined_key
            or "yurt disi doviz transfer" in combined_key
        )
        and not _has_transfer_term(
            masraf,
            "swift",
        )
    ):
        masraf = _normalize(
            f"SWIFT - {masraf}"
        )

    return masraf


# =========================================================
# TEK SAYFA PARSER
# =========================================================

def _parse_page(
    html: str,
    page_category: str,
    page_url: str,
    source_name: str,
) -> Tuple[
    List[UcretSatiri],
    Dict[str, int],
]:
    soup = BeautifulSoup(
        html,
        "lxml",
    )

    tables = _root_tables(soup)

    stats: Dict[str, int] = {
        "tables_total": len(tables),
        "fee_tables": 0,
        "ignored_tables": 0,
        "zero_record_tables": 0,
        "candidate_rows": 0,
        "parsed_before_dedup": 0,
        "repeated_headers": 0,
        "notes": 0,
        "invalid_rows": 0,
    }

    rows_out: List[
        UcretSatiri
    ] = []

    for table_index, table in enumerate(
        tables
    ):
        grid = _table_to_rows(
            table
        )

        if not grid:
            stats[
                "ignored_tables"
            ] += 1
            continue

        header_index = _find_header_index(
            grid
        )

        if header_index == -1:
            stats[
                "ignored_tables"
            ] += 1
            continue

        stats[
            "fee_tables"
        ] += 1

        header = [
            _normalize(x)
            for x in grid[
                header_index
            ]
        ]

        columns = _find_columns(
            header
        )

        table_title = _find_table_title(
            table,
            page_category,
        )

        table_record_count = 0

        for raw_row in grid[
            header_index + 1:
        ]:
            row = [
                _normalize(x)
                for x in raw_row
            ]

            if not row or not any(row):
                continue

            stats[
                "candidate_rows"
            ] += 1

            if _row_is_same_header(
                row,
                header,
            ):
                stats[
                    "repeated_headers"
                ] += 1
                continue

            meaningful = sum(
                1
                for value in row
                if value
            )

            if meaningful < 2:
                stats[
                    "notes"
                ] += 1
                continue

            def get(index: int) -> str:
                if (
                    index < 0
                    or index >= len(row)
                ):
                    return ""

                return _normalize(
                    row[index]
                )

            raw_masraf = get(
                columns["masraf"]
            )

            if not raw_masraf:
                stats[
                    "invalid_rows"
                ] += 1
                continue

            asgari_tutar = get(
                columns[
                    "asgari_tutar"
                ]
            )
            asgari_oran = get(
                columns[
                    "asgari_oran"
                ]
            )
            azami_tutar = get(
                columns[
                    "azami_tutar"
                ]
            )
            azami_oran = get(
                columns[
                    "azami_oran"
                ]
            )

            aciklama_raw = get(
                columns[
                    "aciklama"
                ]
            )

            (
                aciklama,
                aciklama_tarihi,
            ) = _parse_aciklama(
                aciklama_raw
            )

            site_tarihi = get(
                columns[
                    "tarih"
                ]
            ).replace("/", ".")

            if not site_tarihi:
                site_tarihi = (
                    aciklama_tarihi
                )

            if not any(
                [
                    asgari_tutar,
                    asgari_oran,
                    azami_tutar,
                    azami_oran,
                    aciklama,
                    site_tarihi,
                ]
            ):
                stats[
                    "invalid_rows"
                ] += 1
                continue

            masraf = _build_masraf(
                raw_masraf,
                table_title,
            )

            stats[
                "parsed_before_dedup"
            ] += 1

            rows_out.append(
                UcretSatiri(
                    kategori=page_category,
                    masraf=masraf,
                    asgari_tutar=asgari_tutar,
                    asgari_oran=asgari_oran,
                    azami_tutar=azami_tutar,
                    azami_oran=azami_oran,
                    aciklama=aciklama,
                    site_guncelleme_tarihi=site_tarihi,
                )
            )

            table_record_count += 1

        if table_record_count == 0:
            stats[
                "zero_record_tables"
            ] += 1

            print(
                f"[halkbank][DEBUG][{source_name}] "
                f"{page_category} | tablo={table_index} | "
                f"başlık={table_title or '-'} | "
                f"header={header}",
                file=sys.stderr,
            )

    print(
        f"[halkbank][{source_name}] "
        f"{page_category}: "
        f"root tablo={stats['tables_total']}, "
        f"ücret tablosu={stats['fee_tables']}, "
        f"satır={len(rows_out)}",
        file=sys.stderr,
    )

    return rows_out, stats


# =========================================================
# TÜM SAYFALAR - REQUESTS
# =========================================================

def _merge_stats(
    target: Dict[str, int],
    source: Dict[str, int],
) -> None:
    for key, value in source.items():
        target[key] = (
            target.get(key, 0)
            + int(value)
        )


def _empty_total_stats() -> Dict[str, int]:
    return {
        "pages_total": 0,
        "pages_with_rows": 0,
        "pages_no_rows": 0,
        "pages_failed": 0,
        "tables_total": 0,
        "fee_tables": 0,
        "ignored_tables": 0,
        "zero_record_tables": 0,
        "candidate_rows": 0,
        "parsed_before_dedup": 0,
        "duplicates": 0,
        "repeated_headers": 0,
        "notes": 0,
        "invalid_rows": 0,
    }


# =========================================================
# RESMÎ HALKBANK API - DİNAMİK TABLOLAR
# =========================================================

def _api_worker_session() -> requests.Session:
    """Her worker için ayrı Session kullan; bağlantı havuzunu güvenle koru."""
    session = getattr(_API_THREAD_LOCAL, "session", None)

    if session is None:
        session = requests.Session()
        session.headers.update(HEADERS)
        _API_THREAD_LOCAL.session = session

    return session


def _discover_api_table_specs(
    html: str,
    page_category: str,
    page_url: str,
) -> List[Tuple[str, str, str, str]]:
    """
    Statik sayfadaki her productservicefee bileşeninin resmî tablo kimliğini
    ve bağlı olduğu accordion başlığını çıkarır.

    Dönüş: (kategori, bölüm başlığı, tablo kimliği, sayfa URL'i)
    """
    soup = BeautifulSoup(html, "lxml")
    specs: List[Tuple[str, str, str, str]] = []
    seen_ids: Set[str] = set()

    for wrapper in soup.select("[data-productservicefeetableid]"):
        table_id = _normalize(
            wrapper.get("data-productservicefeetableid")
        )

        if not table_id or not table_id.isdigit() or table_id in seen_ids:
            continue

        section_title = ""
        accordion = wrapper.find_parent(
            "div",
            class_="cmp-accordion__item",
        )

        if accordion is not None:
            title_node = accordion.select_one(
                ".cmp-accordion__title"
            )
            if title_node is not None:
                section_title = _normalize(
                    title_node.get_text(" ", strip=True)
                )

        if not section_title:
            heading = wrapper.find_previous(
                ["h3", "h4", "h5", "h6"]
            )
            if heading is not None:
                section_title = _normalize(
                    heading.get_text(" ", strip=True)
                )

        if not _is_valid_title(section_title):
            section_title = ""

        seen_ids.add(table_id)
        specs.append(
            (
                page_category,
                section_title,
                table_id,
                page_url,
            )
        )

    return specs


def _combine_api_table_title(
    section_title: str,
    table_name: str,
) -> str:
    """Accordion başlığı ile ürün/tablo adını kayıpsız birleştirir."""
    section_title = _normalize(section_title)
    table_name = _normalize(table_name)

    if not section_title:
        return table_name
    if not table_name:
        return section_title

    section_key = _normalize_key(section_title)
    table_key = _normalize_key(table_name)

    if (
        section_key == table_key
        or section_key in table_key
    ):
        return table_name

    if table_key in section_key:
        return section_title

    return _normalize(
        f"{section_title} - {table_name}"
    )


def _api_item_to_row(
    item: dict,
    page_category: str,
    section_title: str,
    table_name: str,
) -> UcretSatiri:
    raw_masraf = _normalize(item.get("costName"))

    if not raw_masraf:
        raise ScraperError(
            f"Halkbank API boş costName döndürdü: "
            f"kategori={page_category} | tablo={table_name or section_title}"
        )

    aciklama, aciklama_tarihi = _parse_aciklama(
        _normalize(item.get("description"))
    )
    site_tarihi = _normalize(
        item.get("updatedDate")
    ).replace("/", ".")

    if not site_tarihi:
        site_tarihi = aciklama_tarihi

    full_title = _combine_api_table_title(
        section_title,
        table_name,
    )

    return UcretSatiri(
        kategori=page_category,
        masraf=_build_masraf(
            raw_masraf,
            full_title,
        ),
        asgari_tutar=_normalize(item.get("minAmount")),
        asgari_oran=_normalize(item.get("minRate")),
        azami_tutar=_normalize(item.get("maxAmount")),
        azami_oran=_normalize(item.get("maxRate")),
        aciklama=aciklama,
        site_guncelleme_tarihi=site_tarihi,
    )


def _api_payload_to_table(
    payload,
) -> Tuple[str, List[dict]]:
    if not isinstance(payload, dict):
        raise ScraperError("API JSON nesnesi döndürmedi.")

    data = payload.get("data")
    if not isinstance(data, dict):
        raise ScraperError("API data alanı boş/geçersiz.")

    fee_items = data.get("feeItems")
    if not isinstance(fee_items, list):
        raise ScraperError("API feeItems alanı geçersiz.")

    return _normalize(data.get("tableName")), fee_items


def _fetch_official_api_table(
    spec: Tuple[str, str, str, str],
) -> Tuple[Tuple[str, str, str, str], str, List[dict]]:
    page_category, section_title, table_id, page_url = spec
    api_url = HALKBANK_API_URL.format(table_id=table_id)
    last_error: Optional[Exception] = None

    for attempt in range(1, API_MAX_ATTEMPTS + 1):
        try:
            response = _api_worker_session().get(
                api_url,
                headers={
                    "Referer": page_url,
                    "Origin": "https://www.halkbank.com.tr",
                    "Accept": "application/json, text/plain, */*",
                },
                timeout=API_TIMEOUT_SECONDS,
            )
            response.raise_for_status()
            table_name, fee_items = _api_payload_to_table(response.json())

            if not fee_items and table_id not in ALLOWED_EMPTY_API_TABLE_IDS:
                raise ScraperError(
                    "Beklenmeyen API tablosu feeItems=[] döndürdü."
                )

            # GitHub runner'ından 148 isteği bir anda yollamak Halkbank API'sinde
            # geçici 429/5xx ve boş tablo oluşturabiliyor. Başarılı istekler
            # arasında da kısa bir nefes bırak.
            if API_SUCCESS_DELAY_SECONDS:
                time.sleep(API_SUCCESS_DELAY_SECONDS)

            return spec, table_name, fee_items

        except Exception as exc:
            last_error = exc
            if attempt < API_MAX_ATTEMPTS:
                time.sleep(2 ** attempt)

    raise ScraperError(
        "Halkbank resmî API tablosu alınamadı: "
        f"kategori={page_category} | bölüm={section_title or '-'} | "
        f"table_id={table_id} | hata={last_error}"
    )


def _validate_official_api_inventory(
    specs: List[Tuple[str, str, str, str]],
    api_rows: List[UcretSatiri],
    fee_item_count: int,
    empty_table_ids: Set[str],
) -> None:
    if len(specs) < MIN_API_TABLE_COUNT:
        raise ScraperError(
            "Halkbank resmî API tablo envanteri eksik: "
            f"bulunan={len(specs)} | beklenen_en_az={MIN_API_TABLE_COUNT}"
        )

    unexpected_empty = empty_table_ids - ALLOWED_EMPTY_API_TABLE_IDS
    if unexpected_empty:
        raise ScraperError(
            "Halkbank beklenmeyen resmî API tabloları boş: "
            + ", ".join(sorted(unexpected_empty))
        )

    if fee_item_count < MIN_API_ITEM_COUNT:
        raise ScraperError(
            "Halkbank resmî API ücret satırları eksik: "
            f"bulunan={fee_item_count} | beklenen_en_az={MIN_API_ITEM_COUNT}"
        )

    if len(api_rows) < MIN_API_OUTPUT_ROW_COUNT:
        raise ScraperError(
            "Halkbank resmî API çıktı satırları eksik: "
            f"bulunan={len(api_rows)} | "
            f"beklenen_en_az={MIN_API_OUTPUT_ROW_COUNT}"
        )

    searchable = "\n".join(
        _normalize_key(f"{row.kategori} {row.masraf}") for row in api_rows
    )
    missing_terms = [term for term in REQUIRED_API_TERMS if term not in searchable]
    if missing_terms:
        raise ScraperError(
            "Halkbank kritik resmî API kalemleri eksik: "
            + ", ".join(missing_terms)
        )


def _validate_complete_halkbank_rows(
    rows: List[UcretSatiri],
) -> None:
    """Kaynak yolu ne olursa olsun final Halkbank bütünlüğünü doğrular."""
    if len(rows) < MIN_FINAL_ROW_COUNT:
        raise ScraperError(
            "Halkbank final satır sayısı eksik: "
            f"bulunan={len(rows)} | beklenen_en_az={MIN_FINAL_ROW_COUNT}"
        )

    searchable = "\n".join(
        _normalize_key(f"{row.kategori} {row.masraf}") for row in rows
    )
    missing_terms = [term for term in REQUIRED_FINAL_TERMS if term not in searchable]
    if missing_terms:
        raise ScraperError(
            "Halkbank final kritik kalemleri/ürünleri eksik: "
            + ", ".join(missing_terms)
        )


def _discover_official_inventory(
    pages: List[Tuple[str, str]],
    session: requests.Session,
) -> Tuple[
    List[Tuple[str, str, str, str]],
    List[Tuple[str, str, str]],
]:
    api_specs: List[Tuple[str, str, str, str]] = []
    commercial_pages: List[Tuple[str, str, str]] = []

    for page_category, page_url in pages:
        response = _requests_get_with_retry(
            session,
            page_url,
            headers=HEADERS,
            timeout=40,
            label=f"alt ücret sayfası ({page_category})",
        )

        if page_category == "Ticari Ücret ve Komisyonları":
            commercial_pages.append((page_category, page_url, response.text))
            continue

        page_specs = _discover_api_table_specs(
            response.text,
            page_category,
            page_url,
        )
        if not page_specs:
            raise ScraperError(
                "Halkbank alt sayfasında resmî API tablo kimliği bulunamadı: "
                f"{page_category} | {page_url}"
            )
        api_specs.extend(page_specs)

    table_ids = [spec[2] for spec in api_specs]
    if len(table_ids) != len(set(table_ids)):
        raise ScraperError(
            "Halkbank resmî API envanterinde duplicate table_id bulundu."
        )
    if len(api_specs) < MIN_API_TABLE_COUNT:
        raise ScraperError(
            "Halkbank resmî API tablo envanteri eksik: "
            f"bulunan={len(api_specs)} | beklenen_en_az={MIN_API_TABLE_COUNT}"
        )
    if len(commercial_pages) != 1:
        raise ScraperError(
            "Halkbank ticari tarife sayfası envanteri geçersiz: "
            f"bulunan={len(commercial_pages)} | beklenen=1"
        )

    return api_specs, commercial_pages


def _assemble_official_result(
    pages: List[Tuple[str, str]],
    api_specs: List[Tuple[str, str, str, str]],
    api_results: Dict[str, Tuple[str, List[dict]]],
    commercial_pages: List[Tuple[str, str, str]],
    session: requests.Session,
    source_label: str,
) -> Tuple[List[UcretSatiri], Dict[str, int]]:
    stats = _empty_total_stats()
    stats["pages_total"] = len(pages)
    stats["api_tables_discovered"] = len(api_specs)
    stats["api_tables_loaded"] = len(api_results)
    stats["api_items"] = 0
    stats["api_output_rows"] = 0
    stats["api_empty_tables"] = 0

    api_rows: List[UcretSatiri] = []
    page_row_counts: Dict[str, int] = {}
    fee_item_count = 0
    empty_table_ids: Set[str] = set()

    # Sonuçlar hangi HTTP motorundan gelirse gelsin sayfa/tablo sırasını koru.
    for spec in api_specs:
        page_category, section_title, table_id, _ = spec
        if table_id not in api_results:
            raise ScraperError(
                f"Halkbank resmî API table_id={table_id} sonucu eksik."
            )
        table_name, fee_items = api_results[table_id]

        if not fee_items:
            empty_table_ids.add(table_id)
            if table_id not in ALLOWED_EMPTY_API_TABLE_IDS:
                raise ScraperError(
                    "Halkbank beklenmeyen resmî API tablosu boş: "
                    f"table_id={table_id} | tablo={table_name or section_title}"
                )
            if not table_name:
                raise ScraperError(
                    "Halkbank boş resmî API tablosunun adı da boş: "
                    f"table_id={table_id}"
                )
            api_rows.append(
                UcretSatiri(
                    kategori=page_category,
                    masraf=_build_masraf(table_name, section_title),
                    aciklama=(
                        "Ücret tutarı belirtilmemiş. Resmî Halkbank API "
                        "tablosu yayımlanıyor ancak ücret satırı boş."
                    ),
                )
            )
            page_row_counts[page_category] = page_row_counts.get(page_category, 0) + 1
            continue

        fee_item_count += len(fee_items)

        for item in fee_items:
            if not isinstance(item, dict):
                raise ScraperError(
                    f"Halkbank API table_id={table_id} geçersiz satır döndürdü."
                )
            api_rows.append(
                _api_item_to_row(
                    item,
                    page_category,
                    section_title,
                    table_name,
                )
            )
            page_row_counts[page_category] = page_row_counts.get(page_category, 0) + 1

    _validate_official_api_inventory(
        api_specs,
        api_rows,
        fee_item_count,
        empty_table_ids,
    )

    all_rows = list(api_rows)
    stats["api_items"] = fee_item_count
    stats["api_output_rows"] = len(api_rows)
    stats["api_empty_tables"] = len(empty_table_ids)
    stats["tables_total"] += len(api_specs)
    stats["fee_tables"] += len(api_specs)
    stats["candidate_rows"] += len(api_rows)
    stats["parsed_before_dedup"] += len(api_rows)
    stats["pages_with_rows"] += len(page_row_counts)

    for _, page_url, page_html in commercial_pages:
        rows = _scrape_commercial_pdf(page_html, page_url, session=session)
        if not rows:
            raise ScraperError(
                "Halkbank ticari resmî PDF hiç ücret satırı üretmedi."
            )
        all_rows.extend(rows)
        stats["pages_with_rows"] += 1
        stats["tables_total"] += 1
        stats["fee_tables"] += 1
        stats["candidate_rows"] += len(rows)
        stats["parsed_before_dedup"] += len(rows)

    stats["pages_no_rows"] = (
        stats["pages_total"] - stats["pages_with_rows"] - stats["pages_failed"]
    )
    print(
        f"[halkbank][{source_label}] "
        f"tablo={stats['api_tables_loaded']}/{stats['api_tables_discovered']} | "
        f"ücret_satırı={stats['api_items']} | "
        f"boş_tarife={stats['api_empty_tables']} | "
        f"api_çıktı={stats['api_output_rows']} | "
        f"toplam_ham={len(all_rows)}",
        file=sys.stderr,
    )
    return all_rows, stats


def _fetch_api_specs_playwright(
    api_specs: List[Tuple[str, str, str, str]],
) -> Dict[str, Tuple[str, List[dict]]]:
    """Yalnız verilen API tablolarını Playwright HTTP motoruyla kurtarır."""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise ScraperError(
            "Halkbank Playwright HTTP fallback için playwright gerekli."
        ) from exc

    api_results: Dict[str, Tuple[str, List[dict]]] = {}

    with sync_playwright() as playwright:
        request_context = playwright.request.new_context(
            extra_http_headers={
                "User-Agent": HEADERS["User-Agent"],
                "Accept-Language": HEADERS["Accept-Language"],
                "Accept": "application/json, text/plain, */*",
            },
            timeout=API_TIMEOUT_SECONDS * 1000,
        )
        try:
            total = len(api_specs)
            for index, spec in enumerate(api_specs, start=1):
                page_category, section_title, table_id, page_url = spec
                api_url = HALKBANK_API_URL.format(table_id=table_id)
                last_error: Optional[Exception] = None

                for attempt in range(1, API_MAX_ATTEMPTS + 1):
                    response = None
                    try:
                        response = request_context.get(
                            api_url,
                            headers={
                                "Referer": page_url,
                                "Origin": "https://www.halkbank.com.tr",
                            },
                            timeout=API_TIMEOUT_SECONDS * 1000,
                            fail_on_status_code=False,
                        )
                        if not response.ok:
                            raise ScraperError(
                                f"HTTP {response.status}: "
                                f"{_normalize(response.status_text)}"
                            )
                        table_name, fee_items = _api_payload_to_table(response.json())
                        if (
                            not fee_items
                            and table_id not in ALLOWED_EMPTY_API_TABLE_IDS
                        ):
                            raise ScraperError(
                                "Beklenmeyen API tablosu feeItems=[] döndürdü."
                            )
                        api_results[table_id] = (table_name, fee_items)
                        if API_SUCCESS_DELAY_SECONDS:
                            time.sleep(API_SUCCESS_DELAY_SECONDS)
                        break
                    except Exception as exc:
                        last_error = exc
                        if attempt < API_MAX_ATTEMPTS:
                            time.sleep(2 ** (attempt - 1))
                    finally:
                        if response is not None:
                            try:
                                response.dispose()
                            except Exception:
                                pass
                else:
                    raise ScraperError(
                        "Halkbank Playwright HTTP API tablosu alınamadı: "
                        f"kategori={page_category} | "
                        f"bölüm={section_title or '-'} | "
                        f"table_id={table_id} | hata={last_error}"
                    )

                if index % 25 == 0 or index == total:
                    print(
                        "[halkbank][playwright-http-recovery] "
                        f"tablo={index}/{total}",
                        file=sys.stderr,
                    )
        finally:
            request_context.dispose()

    return api_results


def _scrape_all_official_api(
    pages: List[Tuple[str, str]],
    session: requests.Session,
) -> Tuple[List[UcretSatiri], Dict[str, int]]:
    """
    Resmî API tablolarını requests ile sıralı çeker. Tekil/geçici bir tablo
    hatası olursa başarılı 147 tabloyu atmak yerine yalnız başarısız tablo(lar)
    Playwright HTTP motoruyla yeniden alınır.
    """
    api_specs, commercial_pages = _discover_official_inventory(pages, session)
    api_results: Dict[str, Tuple[str, List[dict]]] = {}
    failed_specs: List[Tuple[str, str, str, str]] = []
    failed_errors: Dict[str, str] = {}

    total = len(api_specs)
    for index, spec in enumerate(api_specs, start=1):
        try:
            loaded_spec, table_name, fee_items = _fetch_official_api_table(spec)
            api_results[loaded_spec[2]] = (table_name, fee_items)
        except Exception as exc:
            failed_specs.append(spec)
            failed_errors[spec[2]] = str(exc)
            print(
                "[halkbank][UYARI] requests API tablosu kurtarmaya ayrıldı: "
                f"table_id={spec[2]} | hata={exc}",
                file=sys.stderr,
            )

        if index % 25 == 0 or index == total:
            print(
                "[halkbank][requests-api] "
                f"tablo={index}/{total} | "
                f"başarılı={len(api_results)} | "
                f"kurtarılacak={len(failed_specs)}",
                file=sys.stderr,
            )

    source_label = "official-api-sequential"
    if failed_specs:
        # İlk dört deneme aynı bağlantı havuzunu kullanır. GitHub runner'da
        # geçici DNS/TLS bozulması oturuma yapışabildiği için yalnız başarısız
        # tabloları kısa bekleme sonrası yepyeni Session ile bir tur daha dene.
        old_session = getattr(_API_THREAD_LOCAL, "session", None)
        if old_session is not None:
            try:
                old_session.close()
            except Exception:
                pass
        _API_THREAD_LOCAL.session = None
        time.sleep(3)

        still_failed: List[Tuple[str, str, str, str]] = []
        for spec in failed_specs:
            try:
                loaded_spec, table_name, fee_items = _fetch_official_api_table(spec)
                api_results[loaded_spec[2]] = (table_name, fee_items)
                print(
                    "[halkbank][requests-fresh-session] "
                    f"table_id={loaded_spec[2]} kurtarıldı.",
                    file=sys.stderr,
                )
            except Exception as exc:
                still_failed.append(spec)
                failed_errors[spec[2]] = (
                    f"ilk={failed_errors.get(spec[2], '-')} | "
                    f"yeni_oturum={exc}"
                )
        failed_specs = still_failed
        if not failed_specs:
            source_label = "requests + fresh-session-recovery"

    if failed_specs:
        print(
            "[halkbank][UYARI] Yalnız başarısız API tabloları "
            "Playwright HTTP motoruyla yeniden çekilecek: "
            + ", ".join(spec[2] for spec in failed_specs),
            file=sys.stderr,
        )
        try:
            recovered = _fetch_api_specs_playwright(failed_specs)
        except Exception as exc:
            details = " | ".join(
                f"{table_id}: {error}"
                for table_id, error in failed_errors.items()
            )
            raise ScraperError(
                "Halkbank başarısız API tabloları kurtarılamadı. "
                f"Requests={details} | Playwright={exc}"
            ) from exc
        api_results.update(recovered)
        source_label = "requests + targeted-playwright-recovery"

    return _assemble_official_result(
        pages,
        api_specs,
        api_results,
        commercial_pages,
        session,
        source_label,
    )


def _scrape_all_playwright_request_api(
    pages: List[Tuple[str, str]],
    session: requests.Session,
) -> Tuple[List[UcretSatiri], Dict[str, int]]:
    """Son çare olarak 148 tablonun tamamını Playwright HTTP ile çeker."""
    api_specs, commercial_pages = _discover_official_inventory(pages, session)
    api_results = _fetch_api_specs_playwright(api_specs)

    return _assemble_official_result(
        pages,
        api_specs,
        api_results,
        commercial_pages,
        session,
        "playwright-http-sequential",
    )


def _scrape_all_requests(
    pages: List[Tuple[str, str]],
    session: requests.Session,
) -> Tuple[
    List[UcretSatiri],
    Dict[str, int],
]:
    all_rows: List[
        UcretSatiri
    ] = []

    total = _empty_total_stats()
    total["pages_total"] = len(pages)

    for page_category, page_url in pages:
        try:
            response = session.get(
                page_url,
                headers=HEADERS,
                timeout=40,
            )
            response.raise_for_status()

            if (
                page_category
                == "Ticari Ücret ve Komisyonları"
            ):
                rows = _scrape_commercial_pdf(
                    response.text,
                    page_url,
                    session=session,
                )

                stats = {
                    "tables_total": 1,
                    "fee_tables": 1,
                    "ignored_tables": 0,
                    "zero_record_tables": (
                        0 if rows else 1
                    ),
                    "candidate_rows": len(rows),
                    "parsed_before_dedup": len(rows),
                    "repeated_headers": 0,
                    "notes": 0,
                    "invalid_rows": 0,
                }
            else:
                rows, stats = _parse_page(
                    response.text,
                    page_category,
                    page_url,
                    "requests",
                )

            _merge_stats(
                total,
                stats,
            )

            if rows:
                total[
                    "pages_with_rows"
                ] += 1
            else:
                total[
                    "pages_no_rows"
                ] += 1

            all_rows.extend(rows)

        except Exception as exc:
            total[
                "pages_failed"
            ] += 1

            print(
                f"[halkbank][HATA][requests] "
                f"{page_category} çekilemedi: {exc}",
                file=sys.stderr,
            )

    return all_rows, total


# =========================================================
# PLAYWRIGHT FALLBACK - TEK BROWSER
# =========================================================

def _scrape_all_playwright(
    pages: List[Tuple[str, str]],
) -> Tuple[
    List[UcretSatiri],
    Dict[str, int],
]:
    try:
        from playwright.sync_api import (
            sync_playwright,
        )
    except ImportError as exc:
        raise ScraperError(
            "Playwright kurulu değil."
        ) from exc

    all_rows: List[
        UcretSatiri
    ] = []

    total = _empty_total_stats()
    total["pages_total"] = len(pages)

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True
        )

        try:
            context = browser.new_context(
                user_agent=HEADERS[
                    "User-Agent"
                ],
                locale="tr-TR",
                timezone_id="Europe/Istanbul",
                viewport={
                    "width": 1440,
                    "height": 1000,
                },
                extra_http_headers={
                    "Accept-Language": (
                        HEADERS[
                            "Accept-Language"
                        ]
                    ),
                },
            )

            page = context.new_page()

            for (
                page_category,
                page_url,
            ) in pages:
                try:
                    page.goto(
                        page_url,
                        timeout=90000,
                        wait_until=(
                            "domcontentloaded"
                        ),
                    )

                    page.wait_for_timeout(
                        1200
                    )

                    # Halkbank'ta özellikle kredi kartı sayfasındaki
                    # ücret bölümleri klasik aria-expanded dışında H3 /
                    # accordion wrapper ile tetiklenebiliyor.
                    for click_round in range(3):
                        try:
                            clicked = page.evaluate("""
                            () => {
                                const targets = new Set();

                                document.querySelectorAll(
                                    "[aria-expanded='false'], "
                                    "[data-bs-toggle='collapse'], "
                                    "[data-toggle='collapse'], "
                                    ".accordion-button, "
                                    ".accordion-header, "
                                    "[role='button']"
                                ).forEach(el => targets.add(el));

                                // H3'e bağlı parent/trigger'ları da ekle.
                                document.querySelectorAll("h3").forEach(h => {
                                    targets.add(h);

                                    const parent = h.parentElement;
                                    if (parent) targets.add(parent);

                                    const button = h.closest("button");
                                    if (button) targets.add(button);

                                    const roleButton = h.closest(
                                        "[role='button']"
                                    );
                                    if (roleButton) targets.add(roleButton);
                                });

                                let count = 0;

                                for (const el of targets) {
                                    try {
                                        el.click();
                                        count++;
                                    } catch (_) {}
                                }

                                return count;
                            }
                            """)

                            if clicked:
                                page.wait_for_timeout(
                                    900
                                )

                            # Kredi kartı içeriği geldiyse daha fazla tıklama yapma.
                            current_tables = page.locator(
                                "table"
                            ).count()

                            if current_tables > 0:
                                break

                        except Exception:
                            break

                    # Dinamik bileşenlerin hepsi ayrı API çağrısıyla doluyor.
                    # Sabit süre beklemek yerine her wrapper içinde gerçek veri satırı
                    # oluşana kadar bekle; eksik DOM'u başarılı kabul etme.
                    if page_category != "Ticari Ücret ve Komisyonları":
                        expected_components = page.locator(
                            "[data-productservicefeetableid]"
                        ).count()

                        if expected_components <= 0:
                            raise ScraperError(
                                "Playwright sayfasında Halkbank API tablo "
                                f"bileşeni bulunamadı: {page_category}"
                            )

                        page.wait_for_function(
                            """
                            () => {
                                const wrappers = Array.from(
                                    document.querySelectorAll(
                                        "[data-productservicefeetableid]"
                                    )
                                );
                                return wrappers.length > 0 && wrappers.every(
                                    wrapper =>
                                        wrapper.querySelector("table") !== null &&
                                        wrapper.querySelector("tbody tr") !== null
                                );
                            }
                            """,
                            timeout=180000,
                        )

                        loaded_components = page.evaluate(
                            """
                            () => Array.from(
                                document.querySelectorAll(
                                    "[data-productservicefeetableid]"
                                )
                            ).filter(
                                wrapper =>
                                    wrapper.querySelector("table") !== null &&
                                    wrapper.querySelector("tbody tr") !== null
                            ).length
                            """
                        )

                        if loaded_components != expected_components:
                            raise ScraperError(
                                "Playwright Halkbank tablo yüklemesi eksik: "
                                f"kategori={page_category} | "
                                f"yüklenen={loaded_components}/{expected_components}"
                            )

                        print(
                            "[halkbank][playwright-complete] "
                            f"{page_category}: "
                            f"{loaded_components}/{expected_components} tablo doldu.",
                            file=sys.stderr,
                        )

                    # XHR/JS render tamamlanması için kısa bekleme.
                    try:
                        page.wait_for_load_state(
                            "networkidle",
                            timeout=5000,
                        )
                    except Exception:
                        pass

                    # Lazy içerik için kısa scroll.
                    for _ in range(45):
                        try:
                            page.evaluate(
                                "() => window.scrollBy("
                                "0, Math.max("
                                "window.innerHeight * 0.9, 700"
                                "))"
                            )
                            page.wait_for_timeout(
                                50
                            )

                            bottom = page.evaluate(
                                "() => "
                                "(window.innerHeight + window.scrollY) "
                                ">= "
                                "((document.documentElement "
                                "? document.documentElement.scrollHeight "
                                ": 0) - 120)"
                            )

                            if bottom:
                                break

                        except Exception:
                            break

                    html = page.content()

                    if (
                        page_category
                        == "Ticari Ücret ve Komisyonları"
                    ):
                        rows = _scrape_commercial_pdf(
                            html,
                            page_url,
                        )

                        stats = {
                            "tables_total": 1,
                            "fee_tables": 1,
                            "ignored_tables": 0,
                            "zero_record_tables": (
                                0 if rows else 1
                            ),
                            "candidate_rows": len(rows),
                            "parsed_before_dedup": len(rows),
                            "repeated_headers": 0,
                            "notes": 0,
                            "invalid_rows": 0,
                        }
                    else:
                        rows, stats = _parse_page(
                            html,
                            page_category,
                            page_url,
                            "playwright",
                        )

                    if (
                        page_category
                        == "Kredi Kartları ve Banka Kartları"
                        and not rows
                    ):
                        try:
                            h3_texts = page.locator(
                                "h3"
                            ).all_text_contents()

                            print(
                                "[halkbank][DEBUG] "
                                "Kredi kartı sayfası hâlâ 0 tablo. "
                                f"H3 sayısı={len(h3_texts)} | "
                                f"H3={h3_texts[:20]}",
                                file=sys.stderr,
                            )

                            print(
                                "[halkbank][DEBUG] "
                                "Kredi kartı selector sayıları: "
                                f"button={page.locator('button').count()}, "
                                f"aria-expanded={page.locator('[aria-expanded]').count()}, "
                                f"collapse={page.locator('[class*=collapse]').count()}",
                                file=sys.stderr,
                            )
                        except Exception:
                            pass

                    _merge_stats(
                        total,
                        stats,
                    )

                    if rows:
                        total[
                            "pages_with_rows"
                        ] += 1
                    else:
                        total[
                            "pages_no_rows"
                        ] += 1

                    all_rows.extend(
                        rows
                    )

                except Exception as exc:
                    total[
                        "pages_failed"
                    ] += 1

                    print(
                        f"[halkbank][HATA][playwright] "
                        f"{page_category} çekilemedi: {exc}",
                        file=sys.stderr,
                    )

        finally:
            browser.close()

    expected_fee_tables = MIN_API_TABLE_COUNT + 1  # dinamik tablolar + ticari PDF
    if (
        total.get("pages_failed", 0)
        or total.get("pages_with_rows", 0) != len(pages)
        or total.get("fee_tables", 0) < expected_fee_tables
    ):
        raise ScraperError(
            "Halkbank eksiksiz Playwright envanteri doğrulanamadı: "
            f"sayfa={total.get('pages_with_rows', 0)}/{len(pages)} | "
            f"hatalı_sayfa={total.get('pages_failed', 0)} | "
            f"ücret_tablosu={total.get('fee_tables', 0)}/"
            f"{expected_fee_tables}"
        )

    return all_rows, total


# =========================================================
# DEDUP / RAPORLAR
# =========================================================

def _deduplicate(
    rows: List[UcretSatiri],
    stats: Dict[str, int],
) -> List[UcretSatiri]:
    result: List[
        UcretSatiri
    ] = []

    seen: Set[
        Tuple[
            str,
            str,
            str,
            str,
            str,
            str,
            str,
            str,
        ]
    ] = set()

    duplicates = 0

    for row in rows:
        key = (
            row.kategori,
            row.masraf,
            row.asgari_tutar,
            row.asgari_oran,
            row.azami_tutar,
            row.azami_oran,
            row.aciklama,
            row.site_guncelleme_tarihi,
        )

        if key in seen:
            duplicates += 1
            continue

        seen.add(key)
        result.append(row)

    stats["duplicates"] = duplicates

    return result


def _print_category_report(
    rows: List[UcretSatiri],
) -> None:
    counts: Dict[
        str,
        int,
    ] = {}

    for row in rows:
        counts[row.kategori] = (
            counts.get(
                row.kategori,
                0,
            )
            + 1
        )

    print(
        "",
        file=sys.stderr,
    )

    print(
        "[halkbank] ===== KATEGORİ RAPORU =====",
        file=sys.stderr,
    )

    for category, count in sorted(
        counts.items(),
        key=lambda item: (
            -item[1],
            item[0],
        ),
    ):
        print(
            f"[halkbank] "
            f"{category} -> "
            f"{count} kayıt",
            file=sys.stderr,
        )

    print(
        "[halkbank] ===========================",
        file=sys.stderr,
    )


def _print_integrity_report(
    stats: Dict[str, int],
    result_count: int,
) -> None:
    candidate_rows = stats.get(
        "candidate_rows",
        0,
    )

    explained = (
        stats.get(
            "parsed_before_dedup",
            0,
        )
        + stats.get(
            "repeated_headers",
            0,
        )
        + stats.get(
            "notes",
            0,
        )
        + stats.get(
            "invalid_rows",
            0,
        )
    )

    print(
        "",
        file=sys.stderr,
    )

    print(
        "[halkbank] ===== BÜTÜNLÜK KONTROLÜ =====",
        file=sys.stderr,
    )

    print(
        f"[halkbank] Alt sayfa: "
        f"{stats.get('pages_total', 0)}",
        file=sys.stderr,
    )

    print(
        f"[halkbank] Veri üreten sayfa: "
        f"{stats.get('pages_with_rows', 0)}",
        file=sys.stderr,
    )

    print(
        f"[halkbank] 0 kayıt üreten sayfa: "
        f"{stats.get('pages_no_rows', 0)}",
        file=sys.stderr,
    )

    print(
        f"[halkbank] Hatalı sayfa: "
        f"{stats.get('pages_failed', 0)}",
        file=sys.stderr,
    )

    print(
        f"[halkbank] Toplam root tablo: "
        f"{stats.get('tables_total', 0)}",
        file=sys.stderr,
    )

    print(
        f"[halkbank] Ücret tablosu: "
        f"{stats.get('fee_tables', 0)}",
        file=sys.stderr,
    )

    if stats.get("api_tables_discovered", 0):
        print(
            f"[halkbank] Resmî API tablo: "
            f"{stats.get('api_tables_loaded', 0)}/"
            f"{stats.get('api_tables_discovered', 0)}",
            file=sys.stderr,
        )
        print(
            f"[halkbank] Resmî API ücret satırı: "
            f"{stats.get('api_items', 0)}",
            file=sys.stderr,
        )
        print(
            f"[halkbank] Resmî API boş tarife/durum satırı: "
            f"{stats.get('api_empty_tables', 0)}",
            file=sys.stderr,
        )
        print(
            f"[halkbank] Resmî API toplam çıktı satırı: "
            f"{stats.get('api_output_rows', stats.get('api_items', 0))}",
            file=sys.stderr,
        )

    print(
        f"[halkbank] İlgisiz/atlanan tablo: "
        f"{stats.get('ignored_tables', 0)}",
        file=sys.stderr,
    )

    print(
        f"[halkbank] Ham ücret satırı adayı: "
        f"{candidate_rows}",
        file=sys.stderr,
    )

    print(
        f"[halkbank] Parse edilen (dedup öncesi): "
        f"{stats.get('parsed_before_dedup', 0)}",
        file=sys.stderr,
    )

    print(
        f"[halkbank] Duplicate: "
        f"{stats.get('duplicates', 0)}",
        file=sys.stderr,
    )

    print(
        f"[halkbank] Tekrarlanan header: "
        f"{stats.get('repeated_headers', 0)}",
        file=sys.stderr,
    )

    print(
        f"[halkbank] Not / tek hücreli satır: "
        f"{stats.get('notes', 0)}",
        file=sys.stderr,
    )

    print(
        f"[halkbank] Geçersiz / boş veri satırı: "
        f"{stats.get('invalid_rows', 0)}",
        file=sys.stderr,
    )

    print(
        f"[halkbank] 0 kayıt üreten ücret tablosu: "
        f"{stats.get('zero_record_tables', 0)}",
        file=sys.stderr,
    )

    print(
        f"[halkbank] Excel'e gidecek benzersiz satır: "
        f"{result_count}",
        file=sys.stderr,
    )

    if (
        candidate_rows == explained
        and stats.get(
            "pages_failed",
            0,
        ) == 0
        and stats.get(
            "pages_no_rows",
            0,
        ) == 0
        and stats.get(
            "zero_record_tables",
            0,
        ) == 0
    ):
        print(
            "[halkbank] BÜTÜNLÜK: OK - "
            "tüm alt sayfalar veri üretti ve "
            "aday satırların tamamı açıklandı.",
            file=sys.stderr,
        )
    else:
        print(
            "[halkbank] BÜTÜNLÜK: UYARI - "
            "0 kayıt üreten / hatalı alt sayfa "
            "veya açıklanamayan satır var.",
            file=sys.stderr,
        )

    print(
        "[halkbank] ===============================",
        file=sys.stderr,
    )


def _print_transfer_report(
    rows: List[UcretSatiri],
) -> None:
    print(
        "",
        file=sys.stderr,
    )

    print(
        "[halkbank] ===== PARA AKTARMA KONTROLÜ =====",
        file=sys.stderr,
    )

    for label, term in [
        ("FAST", "fast"),
        ("EFT", "eft"),
        ("Havale", "havale"),
        ("SWIFT", "swift"),
    ]:
        found = [
            row
            for row in rows
            if _has_transfer_term(
                row.masraf,
                term,
            )
        ]

        print(
            f"[halkbank] {label}: "
            f"{len(found)} kayıt",
            file=sys.stderr,
        )

        for row in found[:12]:
            print(
                f"    - {row.masraf}",
                file=sys.stderr,
            )

        if not found:
            print(
                f"[halkbank][UYARI] "
                f"{label} MASRAF alanında hiç bulunamadı.",
                file=sys.stderr,
            )

    print(
        "[halkbank] =================================",
        file=sys.stderr,
    )


def _print_extra_product_report(
    rows: List[UcretSatiri],
) -> None:
    print(
        "",
        file=sys.stderr,
    )

    print(
        "[halkbank] ===== EK ÜRÜN KONTROLÜ =====",
        file=sys.stderr,
    )

    for label, term in [
        ("Fatura", "fatura"),
        ("HGS", "hgs"),
        ("Kiralık Kasa", "kiralik kasa"),
    ]:
        found = [
            row
            for row in rows
            if (
                term in _normalize_key(
                    row.masraf
                )
                or term in _normalize_key(
                    row.kategori
                )
            )
        ]

        print(
            f"[halkbank] {label}: "
            f"{len(found)} kayıt",
            file=sys.stderr,
        )

        for row in found[:8]:
            print(
                f"    - "
                f"[{row.kategori}] "
                f"{row.masraf}",
                file=sys.stderr,
            )

    print(
        "[halkbank] =============================",
        file=sys.stderr,
    )


# =========================================================
# ANA SCRAPER
# =========================================================

def scrape_halkbank(
    ana_url: str = HALKBANK_ANA_URL,
) -> List[UcretSatiri]:
    print(
        f"[halkbank] SÜRÜM: "
        f"{SCRAPER_VERSION}",
        file=sys.stderr,
    )

    print(
        f"[halkbank] "
        f"Alt sayfalar tespit ediliyor: "
        f"{ana_url}",
        file=sys.stderr,
    )

    session = requests.Session()
    session.headers.update(
        HEADERS
    )

    pages = _discover_pages(
        ana_url,
        session,
    )

    if not pages:
        raise ScraperError(
            "Halkbank alt sayfa listesi boş."
        )

    # Önce sıralı requests yolu. GitHub Actions ağında TLS/429/5xx olursa
    # Chromium DOM'unu beklemek yerine Playwright'ın ayrı HTTP motoru aynı
    # 148 resmî tabloyu yine sıralı ve bütünlük kontrollü çeker.
    try:
        rows, stats = _scrape_all_official_api(pages, session)
        source = "official-api-sequential + commercial-pdf"
    except Exception as api_exc:
        print(
            "[halkbank][UYARI] Doğrudan resmî API yolu tamamlanamadı; "
            "sıralı Playwright HTTP fallback çalışacak. "
            f"Hata={api_exc}",
            file=sys.stderr,
        )
        try:
            rows, stats = _scrape_all_playwright_request_api(pages, session)
            source = "playwright-http-sequential + commercial-pdf"
        except Exception as playwright_exc:
            raise ScraperError(
                "Halkbank hem requests hem Playwright HTTP "
                "yolunda başarısız oldu. "
                f"API={api_exc} | Playwright={playwright_exc}"
            ) from playwright_exc

    if not rows or stats is None:
        raise ScraperError(
            "Halkbank sayfalarından hiçbir ücret satırı çekilemedi."
        )

    rows = _deduplicate(
        rows,
        stats,
    )

    rows = sorted(
        rows,
        key=lambda row: (
            _normalize_key(
                row.kategori
            ),
            _normalize_key(
                row.masraf
            ),
        ),
    )

    _validate_complete_halkbank_rows(rows)

    print(
        f"[halkbank] Kullanılan kaynak: "
        f"{source}",
        file=sys.stderr,
    )

    _print_category_report(
        rows
    )

    _print_integrity_report(
        stats,
        len(rows),
    )

    _print_transfer_report(
        rows
    )

    _print_extra_product_report(
        rows
    )

    print(
        f"[halkbank] Toplam "
        f"{len(rows)} benzersiz satır bulundu.",
        file=sys.stderr,
    )

    return rows


# =========================================================
# TEST
# =========================================================

if __name__ == "__main__":
    try:
        veriler = scrape_halkbank()

        print()
        print("=" * 70)
        print("HALKBANK SCRAPER")
        print("=" * 70)
        print(
            f"Toplam çekilen ücret: "
            f"{len(veriler)}"
        )
        print()

        for i, row in enumerate(
            veriler[:40],
            start=1,
        ):
            print(
                f"{i}. "
                f"[{row.kategori}] "
                f"{row.masraf}"
            )
            print(
                f"   Asgari Tutar : "
                f"{row.asgari_tutar}"
            )
            print(
                f"   Asgari Oran  : "
                f"{row.asgari_oran}"
            )
            print(
                f"   Azami Tutar  : "
                f"{row.azami_tutar}"
            )
            print(
                f"   Azami Oran   : "
                f"{row.azami_oran}"
            )
            print(
                f"   Açıklama     : "
                f"{row.aciklama}"
            )
            print(
                f"   Tarih        : "
                f"{row.site_guncelleme_tarihi}"
            )
            print("-" * 70)

    except Exception as exc:
        print(
            f"[halkbank][HATA] {exc}",
            file=sys.stderr,
        )
        sys.exit(1)
