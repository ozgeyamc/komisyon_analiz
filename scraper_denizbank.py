"""
DenizBank "Ürün ve Hizmet Ücretleri" scraper.

Bu sürüm:
- Server-rendered sayfa için requests'i birincil kullanır.
- requests beklenenden az veri döndürürse Playwright fallback çalıştırır.
- tb-1, tb-2 ... gibi bölüm sayısını sabitlemez; tüm root tabloları tarar.
- DenizBank'ın "kanal satırı + header + veri" tablo yapısını korur.
- Ana kategori + alt kategori + kanal + işlem adını hiyerarşik olarak işler.
- EFT / FAST / Havale / SWIFT ifadelerinin Excel MASRAF filtresinde kaybolmasını önler.
- Duplicate kayıtları temizler.
- Bütünlük, kategori ve para aktarma raporu üretir.
- main.py / update_excel.py ile uyumlu UcretSatiri listesi döndürür.
"""

import re
import sys
from dataclasses import dataclass
from typing import Dict, List, Optional, Set, Tuple

import requests
from bs4 import BeautifulSoup


SCRAPER_VERSION = "2026-08-19-v2-denizbank-integrity"

DENIZBANK_URL = "https://www.denizbank.com/urun-ve-hizmet-ucretleri"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "tr-TR,tr;q=0.9,en;q=0.8",
}


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
# METİN YARDIMCILARI
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


def _join_unique(parts: List[str], sep: str = " - ") -> str:
    clean: List[str] = []

    for part in parts:
        part = _normalize(part)

        if not part:
            continue

        if any(_same_text(part, existing) for existing in clean):
            continue

        clean.append(part)

    return sep.join(clean)


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
        # Açıklama bilgi içeriyor; silmiyoruz.
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
    """
    Kelime sınırıyla arar.
    Örneğin 'defteri' yanlışlıkla EFT sayılmaz.
    """
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


def _contains_transfer_term(text: str) -> bool:
    return any(
        _has_transfer_term(text, term)
        for term in TRANSFER_TERMS
    )


# =========================================================
# DOM / KATEGORİ
# =========================================================

TB_ID_PATTERN = re.compile(r"^tb-\d+$", re.IGNORECASE)

INVALID_TITLES = {
    "müşteri ol",
    "ara",
    "kapat",
    "menü",
    "ana sayfa",
    "bireysel",
    "ürün ve hizmet ücretleri",
    "ürün hizmet ücretleri",
    "işlem türü",
    "asgari tutar",
    "asgari oran",
    "azami tutar",
    "azami oran",
    "açıklama",
    "güncelleme tarihi",
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


def _section_for_table(table):
    return table.find_parent(id=TB_ID_PATTERN)


def _main_category(section) -> str:
    if not section:
        return "Genel"

    # Bölümün kendi ana başlığını bul.
    for heading in section.find_all(
        ["h1", "h2", "h3", "h4", "h5"],
        limit=5,
    ):
        text = _normalize(
            heading.get_text(" ", strip=True)
        )

        if _is_valid_title(text):
            return text

    return "Genel"


def _sub_category(table, section, main_category: str) -> str:
    """
    Tabloya en yakın önceki heading'i bulur.
    Section varsa aynı tb-* bölümünün dışına çıkmaz.
    """
    for heading in table.find_all_previous(
        ["h1", "h2", "h3", "h4", "h5"],
        limit=30,
    ):
        if section is not None:
            heading_section = heading.find_parent(id=TB_ID_PATTERN)

            if heading_section is not section:
                continue

        text = _normalize(
            heading.get_text(" ", strip=True)
        )

        if not _is_valid_title(text):
            continue

        if _same_text(text, main_category):
            continue

        return text

    return ""


def _channel_name(table) -> str:
    """
    DenizBank'ta çoğu ücret tablosunun ilk satırı tek hücreli kanal:
      Şube
      İnternet-Mobil-Kiosk
      Tüm Kanallar
      ...
    """
    first_tr = table.find("tr")

    if not first_tr:
        return ""

    cells = first_tr.find_all(
        ["th", "td"],
        recursive=False,
    )

    if len(cells) != 1:
        return ""

    text = _normalize(
        cells[0].get_text(" ", strip=True)
    )

    if not _is_valid_title(text):
        return ""

    return text


def _fallback_heading_context(table) -> Tuple[str, str]:
    """
    tb-* parent bulunamazsa sayfa genelinde yakın iki heading'den
    güvenli bir kategori üretir.
    """
    headings = []

    for heading in table.find_all_previous(
        ["h1", "h2", "h3", "h4", "h5"],
        limit=8,
    ):
        text = _normalize(
            heading.get_text(" ", strip=True)
        )

        if not _is_valid_title(text):
            continue

        if any(
            _same_text(text, x)
            for x in headings
        ):
            continue

        headings.append(text)

        if len(headings) >= 2:
            break

    if not headings:
        return "Genel", ""

    if len(headings) == 1:
        return headings[0], ""

    # En yakın heading alt kategori, bir önceki daha genel kabul edilir.
    return headings[1], headings[0]


# =========================================================
# TABLO SATIRLARI
# =========================================================

def _root_tables(soup):
    return [
        table
        for table in soup.find_all("table")
        if table.find_parent("table") is None
    ]


def _table_rows(table) -> List[List[str]]:
    """
    Root tr/cell'leri kullanır.
    DenizBank ücret tabloları sabit kolonlu olduğu için burada
    colspan'ı kanal satırında koruyup veri satırlarını olduğu gibi alıyoruz.
    """
    rows: List[List[str]] = []

    for tr in table.find_all("tr"):
        if tr.find_parent("table") is not table:
            continue

        cells = tr.find_all(
            ["th", "td"],
            recursive=False,
        )

        if not cells:
            continue

        values = [
            _normalize(
                cell.get_text(" ", strip=True)
            )
            for cell in cells
        ]

        if any(values):
            rows.append(values)

    return rows


# =========================================================
# HEADER
# =========================================================

def _header_score(row: List[str]) -> int:
    headers = [
        _normalize_key(cell)
        for cell in row
    ]

    tests = [
        ["islem", "turu"],
        ["masraf"],
        ["asgari", "tutar"],
        ["asgari", "oran"],
        ["azami", "tutar"],
        ["azami", "oran"],
        ["aciklama"],
        ["guncelleme", "tarihi"],
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
        rows[:6]
    ):
        score = _header_score(row)

        if score > best_score:
            best_score = score
            best_index = index

    # İşlem Türü + ücret kolonlarından birkaçını mutlaka görmeliyiz.
    if best_score < 4:
        return -1

    return best_index


def _find_col(
    headers: List[str],
    keywords: List[str],
) -> int:
    for index, header in enumerate(headers):
        if all(
            keyword in header
            for keyword in keywords
        ):
            return index

    return -1


def _find_columns(
    header_row: List[str],
) -> Dict[str, int]:
    headers = [
        _normalize_key(x)
        for x in header_row
    ]

    masraf = _find_col(
        headers,
        ["islem", "turu"],
    )

    if masraf == -1:
        masraf = _find_col(
            headers,
            ["masraf"],
        )

    if masraf == -1:
        masraf = 0

    result = {
        "masraf": masraf,
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
            ["guncelleme", "tarihi"],
        ),
    }

    # Güncel DenizBank yapısı 7 kolon.
    if len(header_row) >= 7:
        fallbacks = {
            "masraf": 0,
            "asgari_tutar": 1,
            "asgari_oran": 2,
            "azami_tutar": 3,
            "azami_oran": 4,
            "aciklama": 5,
            "tarih": 6,
        }

        for key, index in fallbacks.items():
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

    return left == right


# =========================================================
# MASRAF / KATEGORİ
# =========================================================

def _build_category(
    main_category: str,
    sub_category: str,
) -> str:
    return (
        _join_unique(
            [main_category, sub_category]
        )
        or "Genel"
    )


def _build_masraf(
    raw_masraf: str,
    sub_category: str,
    channel: str,
) -> str:
    """
    Hiyerarşiyi MASRAF alanında da korur.

    Örnek:
      Swift - Şube + Çağrı Merkezi -
      Uluslararası Fon Transferi Hesaptan Giden

    Böylece SWIFT kelimesi satır isminde geçmese bile Excel filtresinde
    ilgili ücret bulunur.
    """
    raw_masraf = _normalize(
        raw_masraf
    )

    parts: List[str] = []

    # Transfer alt kategorisini MASRAF'a özellikle taşı.
    # Diğer alt kategoriler de ürün bağlamını korumak için eklenir.
    if sub_category:
        sub_key = _normalize_key(
            sub_category
        )
        raw_key = _normalize_key(
            raw_masraf
        )

        if (
            sub_key != raw_key
            and sub_key not in raw_key
        ):
            parts.append(
                sub_category
            )

    if channel:
        channel_key = _normalize_key(
            channel
        )
        raw_key = _normalize_key(
            raw_masraf
        )

        if (
            channel_key != raw_key
            and channel_key not in raw_key
        ):
            parts.append(channel)

    parts.append(raw_masraf)

    masraf = _join_unique(parts)

    # "Swift" başlığı altında raw satır Swift demese bile prefix yukarıda
    # zaten gelmiş olur. Yine de açık bir güvenlik alias'ı bırakıyoruz.
    combined = _normalize_key(
        " ".join(
            [
                sub_category,
                channel,
                raw_masraf,
            ]
        )
    )

    if (
        "uluslararasi fon transfer" in combined
        and "swift" in _normalize_key(sub_category)
        and not _has_transfer_term(
            masraf,
            "swift",
        )
    ):
        masraf = f"Swift - {masraf}"

    return _normalize(masraf)


# =========================================================
# PARSER
# =========================================================

def _parse_soup(
    soup: BeautifulSoup,
    source_name: str,
) -> Tuple[
    List[UcretSatiri],
    Dict[str, int],
]:
    tables = _root_tables(soup)

    tb_sections = soup.find_all(
        id=TB_ID_PATTERN
    )

    stats: Dict[str, int] = {
        "tb_sections": len(tb_sections),
        "tables_total": len(tables),
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

    results: List[
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

    kategori_sayilari: Dict[
        str,
        int,
    ] = {}

    for table_index, table in enumerate(
        tables
    ):
        rows = _table_rows(table)

        if not rows:
            stats["ignored_tables"] += 1
            continue

        header_index = _find_header_index(
            rows
        )

        if header_index == -1:
            stats["ignored_tables"] += 1
            continue

        stats["fee_tables"] += 1

        header = [
            _normalize(x)
            for x in rows[header_index]
        ]

        columns = _find_columns(
            header
        )

        section = _section_for_table(
            table
        )

        if section is not None:
            main_category = _main_category(
                section
            )
            sub_category = _sub_category(
                table,
                section,
                main_category,
            )
        else:
            (
                main_category,
                sub_category,
            ) = _fallback_heading_context(
                table
            )

        channel = _channel_name(
            table
        )

        kategori = _build_category(
            main_category,
            sub_category,
        )

        table_record_count = 0

        for raw_row in rows[
            header_index + 1:
        ]:
            row = [
                _normalize(x)
                for x in raw_row
            ]

            if not row or not any(row):
                continue

            stats["candidate_rows"] += 1

            if _row_is_same_header(
                row,
                header,
            ):
                stats["repeated_headers"] += 1
                continue

            meaningful_cells = sum(
                1
                for value in row
                if value
            )

            if meaningful_cells < 2:
                stats["notes"] += 1
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
                stats["invalid_rows"] += 1
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

            # Gerçek ücret kayıtlarında en az bir değer/açıklama/tarih olmalı.
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
                stats["invalid_rows"] += 1
                continue

            masraf = _build_masraf(
                raw_masraf=raw_masraf,
                sub_category=sub_category,
                channel=channel,
            )

            stats[
                "parsed_before_dedup"
            ] += 1

            key = (
                kategori,
                masraf,
                asgari_tutar,
                asgari_oran,
                azami_tutar,
                azami_oran,
                aciklama,
                site_tarihi,
            )

            if key in seen:
                stats["duplicates"] += 1
                continue

            seen.add(key)

            results.append(
                UcretSatiri(
                    kategori=kategori,
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

            kategori_sayilari[
                kategori
            ] = (
                kategori_sayilari.get(
                    kategori,
                    0,
                )
                + 1
            )

        if table_record_count == 0:
            stats[
                "zero_record_tables"
            ] += 1

            print(
                f"[denizbank][DEBUG][{source_name}] "
                f"Ücret tablosu {table_index} 0 kayıt üretti | "
                f"Kategori: {kategori} | "
                f"Kanal: {channel or '-'} | "
                f"Satır: {len(rows)} | "
                f"Header: {header}",
                file=sys.stderr,
            )

    print(
        f"[denizbank][{source_name}] "
        f"tb-* bölüm sayısı: "
        f"{stats['tb_sections']}",
        file=sys.stderr,
    )

    print(
        f"[denizbank][{source_name}] "
        f"Toplam root tablo: "
        f"{stats['tables_total']}",
        file=sys.stderr,
    )

    print(
        f"[denizbank][{source_name}] "
        f"Ücret tablosu: "
        f"{stats['fee_tables']}",
        file=sys.stderr,
    )

    print(
        f"[denizbank][{source_name}] "
        f"İlgisiz/atlanan tablo: "
        f"{stats['ignored_tables']}",
        file=sys.stderr,
    )

    print(
        f"[denizbank][{source_name}] "
        f"0 kayıt üreten ücret tablosu: "
        f"{stats['zero_record_tables']}",
        file=sys.stderr,
    )

    print(
        f"[denizbank][{source_name}] "
        f"Benzersiz sonuç: "
        f"{len(results)}",
        file=sys.stderr,
    )

    print(
        "",
        file=sys.stderr,
    )

    print(
        f"[denizbank][{source_name}] "
        "===== KATEGORİ RAPORU =====",
        file=sys.stderr,
    )

    for kategori, count in sorted(
        kategori_sayilari.items(),
        key=lambda item: (
            -item[1],
            item[0],
        ),
    ):
        print(
            f"[denizbank][{source_name}] "
            f"{kategori} -> "
            f"{count} kayıt",
            file=sys.stderr,
        )

    print(
        f"[denizbank][{source_name}] "
        "===========================",
        file=sys.stderr,
    )

    return results, stats


# =========================================================
# REQUESTS
# =========================================================

def _scrape_with_requests(
    url: str = DENIZBANK_URL,
) -> Tuple[
    List[UcretSatiri],
    Dict[str, int],
]:
    response = requests.get(
        url,
        headers=HEADERS,
        timeout=45,
    )

    response.raise_for_status()

    soup = BeautifulSoup(
        response.text,
        "lxml",
    )

    return _parse_soup(
        soup,
        source_name="requests",
    )


# =========================================================
# PLAYWRIGHT FALLBACK
# =========================================================

def _scrape_with_playwright(
    url: str = DENIZBANK_URL,
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

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True
        )

        try:
            context = browser.new_context(
                user_agent=HEADERS[
                    "User-Agent"
                ],
                viewport={
                    "width": 1440,
                    "height": 1000,
                },
                locale="tr-TR",
                timezone_id=(
                    "Europe/Istanbul"
                ),
                extra_http_headers={
                    "Accept-Language": (
                        HEADERS[
                            "Accept-Language"
                        ]
                    ),
                },
            )

            page = context.new_page()

            page.goto(
                url,
                timeout=120000,
                wait_until=(
                    "domcontentloaded"
                ),
            )

            page.wait_for_timeout(
                2500
            )

            # Hidden tab pane içeriklerini DOM'da görünür hale getir.
            try:
                page.evaluate("""
                () => {
                    document.querySelectorAll(
                        '.tab-pane, [role="tabpanel"]'
                    ).forEach(el => {
                        el.classList.add('active', 'show');
                        el.style.display = 'block';
                        el.hidden = false;
                    });
                }
                """)
            except Exception:
                pass

            # Lazy içerik için kısa, güvenli scroll.
            stable = 0
            previous_height = 0

            for _ in range(60):
                try:
                    height = page.evaluate(
                        "() => document.documentElement "
                        "? document.documentElement.scrollHeight : 0"
                    )

                    page.evaluate(
                        "() => window.scrollBy("
                        "0, Math.max(window.innerHeight * 0.8, 600)"
                        ")"
                    )

                    page.wait_for_timeout(
                        80
                    )

                    bottom = page.evaluate(
                        "() => (window.innerHeight + window.scrollY) "
                        ">= ((document.documentElement "
                        "? document.documentElement.scrollHeight : 0) - 100)"
                    )

                    if (
                        bottom
                        and height == previous_height
                    ):
                        stable += 1
                    else:
                        stable = 0

                    previous_height = height

                    if stable >= 4:
                        break

                except Exception:
                    break

            page.wait_for_timeout(
                800
            )

            html = page.content()

        finally:
            browser.close()

    soup = BeautifulSoup(
        html,
        "lxml",
    )

    return _parse_soup(
        soup,
        source_name="playwright",
    )


# =========================================================
# RAPORLAR
# =========================================================

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
        "[denizbank] ===== BÜTÜNLÜK KONTROLÜ =====",
        file=sys.stderr,
    )

    print(
        f"[denizbank] Toplam root tablo: "
        f"{stats.get('tables_total', 0)}",
        file=sys.stderr,
    )

    print(
        f"[denizbank] Ücret tablosu: "
        f"{stats.get('fee_tables', 0)}",
        file=sys.stderr,
    )

    print(
        f"[denizbank] Ham ücret satırı adayı: "
        f"{candidate_rows}",
        file=sys.stderr,
    )

    print(
        f"[denizbank] Parse edilen (dedup öncesi): "
        f"{stats.get('parsed_before_dedup', 0)}",
        file=sys.stderr,
    )

    print(
        f"[denizbank] Duplicate: "
        f"{stats.get('duplicates', 0)}",
        file=sys.stderr,
    )

    print(
        f"[denizbank] Tekrarlanan header: "
        f"{stats.get('repeated_headers', 0)}",
        file=sys.stderr,
    )

    print(
        f"[denizbank] Not / tek hücreli satır: "
        f"{stats.get('notes', 0)}",
        file=sys.stderr,
    )

    print(
        f"[denizbank] Geçersiz / boş veri satırı: "
        f"{stats.get('invalid_rows', 0)}",
        file=sys.stderr,
    )

    print(
        f"[denizbank] 0 kayıt üreten ücret tablosu: "
        f"{stats.get('zero_record_tables', 0)}",
        file=sys.stderr,
    )

    print(
        f"[denizbank] Excel'e gidecek benzersiz satır: "
        f"{result_count}",
        file=sys.stderr,
    )

    if candidate_rows == explained:
        print(
            "[denizbank] BÜTÜNLÜK: OK - "
            "aday satırların tamamı açıklandı.",
            file=sys.stderr,
        )
    else:
        print(
            "[denizbank] BÜTÜNLÜK: UYARI - "
            f"{candidate_rows - explained} "
            "aday satır açıklanamadı.",
            file=sys.stderr,
        )

    print(
        "[denizbank] ===============================",
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
        "[denizbank] ===== PARA AKTARMA KONTROLÜ =====",
        file=sys.stderr,
    )

    for label, needle in [
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
                needle,
            )
        ]

        print(
            f"[denizbank] {label}: "
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
                f"[denizbank][UYARI] "
                f"{label} MASRAF alanında hiç bulunamadı.",
                file=sys.stderr,
            )

    print(
        "[denizbank] =================================",
        file=sys.stderr,
    )


def _print_product_checks(
    rows: List[UcretSatiri],
) -> None:
    checks = [
        ("Fatura", "fatura"),
        ("HGS", "hgs"),
    ]

    print(
        "",
        file=sys.stderr,
    )

    print(
        "[denizbank] ===== EK ÜRÜN KONTROLÜ =====",
        file=sys.stderr,
    )

    for label, needle in checks:
        found = [
            row
            for row in rows
            if (
                needle in _normalize_key(
                    row.masraf
                )
                or needle in _normalize_key(
                    row.kategori
                )
            )
        ]

        print(
            f"[denizbank] {label}: "
            f"{len(found)} kayıt",
            file=sys.stderr,
        )

        for row in found[:8]:
            print(
                f"    - [{row.kategori}] "
                f"{row.masraf}",
                file=sys.stderr,
            )

    print(
        "[denizbank] =============================",
        file=sys.stderr,
    )


# =========================================================
# ANA SCRAPER
# =========================================================

def scrape_denizbank(
    url: str = DENIZBANK_URL,
) -> List[UcretSatiri]:
    print(
        f"[denizbank] SÜRÜM: "
        f"{SCRAPER_VERSION}",
        file=sys.stderr,
    )

    print(
        f"[denizbank] {url} "
        "adresinden veri çekiliyor...",
        file=sys.stderr,
    )

    rows: List[
        UcretSatiri
    ] = []

    stats: Optional[
        Dict[str, int]
    ] = None

    source = ""

    # -----------------------------------------------------
    # 1) REQUESTS
    # -----------------------------------------------------

    try:
        (
            request_rows,
            request_stats,
        ) = _scrape_with_requests(
            url
        )

        print(
            f"[denizbank] requests sonucu: "
            f"{len(request_rows)} "
            "benzersiz satır.",
            file=sys.stderr,
        )

        # DenizBank sayfası server-rendered. Yine de eksik HTML gelirse
        # otomatik olarak Playwright'a düş.
        if (
            request_rows
            and request_stats.get(
                "fee_tables",
                0,
            ) >= 20
            and request_stats.get(
                "candidate_rows",
                0,
            ) >= 100
        ):
            rows = request_rows
            stats = request_stats
            source = "requests"

        else:
            print(
                "[denizbank][UYARI] "
                "requests sonucu beklenenden küçük; "
                "Playwright deneniyor.",
                file=sys.stderr,
            )

    except Exception as exc:
        print(
            f"[denizbank] requests başarısız: "
            f"{exc}",
            file=sys.stderr,
        )

    # -----------------------------------------------------
    # 2) PLAYWRIGHT FALLBACK
    # -----------------------------------------------------

    if not rows:
        try:
            (
                playwright_rows,
                playwright_stats,
            ) = _scrape_with_playwright(
                url
            )

            print(
                f"[denizbank] Playwright sonucu: "
                f"{len(playwright_rows)} "
                "benzersiz satır.",
                file=sys.stderr,
            )

            rows = playwright_rows
            stats = playwright_stats
            source = "playwright"

        except Exception as exc:
            print(
                f"[denizbank] Playwright başarısız: "
                f"{exc}",
                file=sys.stderr,
            )

    if not rows or stats is None:
        raise ScraperError(
            "DenizBank sayfasından hiçbir "
            "ücret satırı çekilemedi."
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
        f"[denizbank] Kullanılan kaynak: "
        f"{source}",
        file=sys.stderr,
    )

    _print_integrity_report(
        stats,
        len(rows),
    )

    _print_transfer_report(
        rows
    )

    _print_product_checks(
        rows
    )

    print(
        f"[denizbank] Toplam "
        f"{len(rows)} benzersiz satır bulundu.",
        file=sys.stderr,
    )

    return rows


# =========================================================
# TEST
# =========================================================

if __name__ == "__main__":
    try:
        veriler = scrape_denizbank()

        print()
        print("=" * 70)
        print("DENİZBANK SCRAPER")
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
            f"[denizbank][HATA] "
            f"{exc}",
            file=sys.stderr,
        )
        sys.exit(1)
