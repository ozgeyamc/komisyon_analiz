"""
Halkbank "Ürün ve Hizmet Ücretleri" scraper.

Bu sürüm:
- Ana ücret sayfasındaki güncel alt sayfaları otomatik bulur.
- requests'i birincil kullanır; tüm sayfalar tek Session ile hızlı çekilir.
- requests sonucu eksik görünürse tek Chromium oturumu ile Playwright fallback kullanır.
- Her alt sayfa için ayrı browser açmaz.
- Güncel Halkbank sayfa başlıklarını / alt başlıklarını MASRAF alanında korur.
- Havale / EFT / FAST başlıklarını kaybetmez.
- "Uluslararası Fon Transferi Ücreti" satırlarını SWIFT filtresinde görünür yapar.
- Duplicate kayıtları temizler.
- Sayfa, tablo, kategori, bütünlük ve para aktarma raporları üretir.
- main.py / update_excel.py ile uyumlu UcretSatiri listesi döndürür.
"""

import io
import re
import sys
from dataclasses import dataclass
from typing import Dict, List, Optional, Set, Tuple
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup


SCRAPER_VERSION = "2026-08-19-v3-halkbank-card-commercial-fix"

HALKBANK_ANA_URL = "https://www.halkbank.com.tr/tr/urun-ve-hizmet-ucretleri"

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

    text = text.replace("%", " ")
    text = re.sub(r"\s+", " ", text)

    return text.strip()


def _same_text(a: str, b: str) -> bool:
    return _normalize_key(a) == _normalize_key(b)


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
    response = session.get(
        ana_url,
        headers=HEADERS,
        timeout=40,
    )
    response.raise_for_status()

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

    result: List[
        UcretSatiri
    ] = []

    # Numara seviyesine göre başlık hiyerarşisi.
    hierarchy: Dict[
        int,
        str,
    ] = {}

    with pdfplumber.open(
        io.BytesIO(pdf_bytes)
    ) as pdf:
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

                for i, row in enumerate(
                    table[:8]
                ):
                    cells = [
                        _normalize_key(
                            _pdf_clean_cell(x)
                        )
                        for x in row
                    ]

                    joined = " | ".join(
                        cells
                    )

                    if (
                        "kalem adi" in joined
                        and "asgari" in joined
                        and "azami" in joined
                    ):
                        header_index = i
                        break

                if header_index == -1:
                    continue

                header = [
                    _normalize_key(
                        _pdf_clean_cell(x)
                    )
                    for x in table[
                        header_index
                    ]
                ]

                def find_col(
                    *needles: str,
                ) -> int:
                    for idx, cell in enumerate(
                        header
                    ):
                        if all(
                            needle in cell
                            for needle in needles
                        ):
                            return idx
                    return -1

                col_name = find_col(
                    "kalem",
                    "adi",
                )
                col_currency = find_col(
                    "para",
                    "birimi",
                )
                col_min_amount = find_col(
                    "asgari",
                    "tutar",
                )
                col_min_rate = find_col(
                    "asgari",
                    "oran",
                )
                col_max_amount = find_col(
                    "azami",
                    "tutar",
                )
                col_max_rate = find_col(
                    "azami",
                    "oran",
                )
                col_desc = find_col(
                    "aciklama",
                )
                col_date = find_col(
                    "guncelle",
                )

                if col_date == -1:
                    col_date = find_col(
                        "tarih",
                    )

                # PDF extractor bazen tarih başlığını boş verir;
                # en sağ kolon tarih görünümündeyse fallback yap.
                if (
                    col_date == -1
                    and len(header) >= 8
                ):
                    col_date = len(
                        header
                    ) - 1

                data_rows = table[
                    header_index + 1:
                ]

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

                    raw_name = get(
                        col_name
                    )

                    if not raw_name:
                        continue

                    number, name = (
                        _numeric_prefix(
                            raw_name
                        )
                    )

                    if not name:
                        name = raw_name

                    currency = get(
                        col_currency
                    )
                    min_amount = get(
                        col_min_amount
                    )
                    min_rate = get(
                        col_min_rate
                    )
                    max_amount = get(
                        col_max_amount
                    )
                    max_rate = get(
                        col_max_rate
                    )
                    desc = get(
                        col_desc
                    )
                    date = get(
                        col_date
                    ).replace("/", ".")

                    has_value = any(
                        [
                            currency,
                            min_amount,
                            min_rate,
                            max_amount,
                            max_rate,
                            desc,
                            date,
                        ]
                    )

                    # Sadece başlık satırıysa hierarchy güncelle.
                    if (
                        number
                        and not has_value
                    ):
                        level = len(
                            number.split(".")
                        )

                        hierarchy[
                            level
                        ] = name

                        for old_level in list(
                            hierarchy
                        ):
                            if old_level > level:
                                del hierarchy[
                                    old_level
                                ]

                        continue

                    # Para birimini tutarlarda kaybetme.
                    if (
                        currency
                        and min_amount
                        and not min_amount.upper().startswith(
                            currency.upper()
                        )
                    ):
                        min_amount = (
                            f"{currency} "
                            f"{min_amount}"
                        )

                    if (
                        currency
                        and max_amount
                        and not max_amount.upper().startswith(
                            currency.upper()
                        )
                    ):
                        max_amount = (
                            f"{currency} "
                            f"{max_amount}"
                        )

                    if number:
                        level = len(
                            number.split(".")
                        )

                        parents = [
                            hierarchy[l]
                            for l in sorted(
                                hierarchy
                            )
                            if l < level
                        ]
                    else:
                        parents = []

                    masraf = " - ".join(
                        [
                            p
                            for p in (
                                parents
                                + [name]
                            )
                            if p
                        ]
                    )

                    masraf = _normalize(
                        masraf
                    )

                    # Ticari PDF'teki EFT/Havale/FAST/SWIFT hiyerarşisi
                    # satır adına yansısın.
                    combined = _normalize_key(
                        " ".join(
                            parents
                            + [name]
                        )
                    )

                    if (
                        "uluslararasi fon transfer" in combined
                        and not _has_transfer_term(
                            masraf,
                            "swift",
                        )
                    ):
                        masraf = (
                            f"SWIFT - {masraf}"
                        )

                    result.append(
                        UcretSatiri(
                            kategori=(
                                "Ticari Ücret ve Komisyonları"
                            ),
                            masraf=masraf,
                            asgari_tutar=min_amount,
                            asgari_oran=min_rate,
                            azami_tutar=max_amount,
                            azami_oran=max_rate,
                            aciklama=desc,
                            site_guncelleme_tarihi=date,
                        )
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

    response = requester.get(
        pdf_url,
        headers=HEADERS,
        timeout=60,
    )
    response.raise_for_status()

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

        return text

    return ""


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

    rows: List[
        UcretSatiri
    ] = []

    stats: Optional[
        Dict[str, int]
    ] = None

    source = ""

    # -----------------------------------------------------
    # 1) REQUESTS - hızlı yol
    # -----------------------------------------------------

    try:
        request_rows, request_stats = (
            _scrape_all_requests(
                pages,
                session,
            )
        )

        print(
            f"[halkbank] requests sonucu: "
            f"{len(request_rows)} ham parse satırı.",
            file=sys.stderr,
        )

        # Güncel ücret sayfalarının önemli kısmı server-rendered.
        # Çok az tablo/satır görülürse dinamik içeriği kaçırmış say.
        if (
            request_rows
            and request_stats.get(
                "fee_tables",
                0,
            ) >= 15
            and request_stats.get(
                "candidate_rows",
                0,
            ) >= 80
            and request_stats.get(
                "pages_failed",
                0,
            ) == 0
        ):
            rows = request_rows
            stats = request_stats
            source = "requests"

        else:
            print(
                "[halkbank][UYARI] "
                "requests sonucu eksik görünüyor; "
                "tek-browser Playwright fallback çalışacak.",
                file=sys.stderr,
            )

    except Exception as exc:
        print(
            f"[halkbank] requests başarısız: {exc}",
            file=sys.stderr,
        )

    # -----------------------------------------------------
    # 2) PLAYWRIGHT - yalnızca gerektiğinde
    # -----------------------------------------------------

    if not rows:
        playwright_rows, playwright_stats = (
            _scrape_all_playwright(
                pages
            )
        )

        rows = playwright_rows
        stats = playwright_stats
        source = "playwright"

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
