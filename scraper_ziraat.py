"""
Ziraat Bankası "Ürün ve Hizmet Ücretleri" scraper.

Bu sürüm:
- Güncel Ziraat Ürün ve Hizmet Ücretleri sayfasındaki tüm ücret tablolarını tarar.
- requests'i birincil kullanır; sonuç eksik görünürse Playwright fallback yapar.
- Root table + ücret header kontrolü ile ilgisiz tabloları ayırır.
- Ana kategoriyi güncel Ziraat bölüm adlarından belirler.
- Kategori altındaki başlık hiyerarşisini MASRAF alanında korur.

Örnek:
    Para Aktarma | EFT - Eft Mesaj Ücreti - 0-14.000 TL arası
    Para Aktarma | FAST - Mobil/ İnternet/Düzenli Ödeme - 0-8.300 TL
    Para Aktarma | Swift - Bankamızdan Hesaptan Ödenen Havaleler(YP) - 0 - 250 USD arası (USD)

- Duplicate kayıtları temizler.
- Bütünlük, kategori, FAST/EFT/Havale/SWIFT ve ek ürün raporları üretir.
- main.py / update_excel.py ile uyumlu UcretSatiri listesi döndürür.
"""

import re
import sys
from dataclasses import dataclass
from typing import Dict, List, Optional, Set, Tuple
from xml.etree import ElementTree as ET

import requests
from bs4 import BeautifulSoup


SCRAPER_VERSION = "2026-08-24-v5-ziraat-full-descriptions"

ZIRAAT_URL = "https://www.ziraatbank.com.tr/tr/urun-ve-hizmet-ucretleri"
ZIRAAT_XML_URL = (
    "https://www.ziraatbank.com.tr/"
    "TuketiciVerileri/TuketiciVerileri.xml"
)

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
    r"(\d{1,2}[./]\d{1,2}[./]\d{4}"
    r"(?:[\s\xa0]+\d{1,2}:\d{2}(?::\d{2})?)?)",
    re.IGNORECASE,
)

DATE_PATTERN_ITIBAR = re.compile(
    r"(\d{1,2}[./]\d{1,2}[./]\d{4})\s+tarihi\s+itibar",
    re.IGNORECASE,
)

DATE_PATTERN_TR = re.compile(
    r"(?:son\s+)?g[üu]ncellenme\s+tarihi\s*:?\s*[\s\xa0]*"
    r"(\d{1,2})\s+"
    r"(ocak|şubat|mart|nisan|mayıs|haziran|temmuz|ağustos|eylül|ekim|kasım|aralık)"
    r"\s+(\d{4})",
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

def _normalize(value) -> str:
    if value is None:
        return ""

    text = str(value)
    text = text.replace("\xa0", " ")
    text = text.replace("\u200b", "")
    text = text.replace("\r", " ")
    text = text.replace("\n", " ")
    text = re.sub(r"\s+", " ", text)

    return text.strip()


def _normalize_key(value) -> str:
    text = _normalize(value).lower()

    # Türkçe büyük İ lower() sonrası "i\u0307" üretebilir.
    # Combining dot'u temizle ki "İşlemleri" -> "islemleri" eşleşsin.
    text = text.replace("\u0307", "")

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

    text = re.sub(r"\s+", " ", text)

    return text.strip()


def _normalize_tutar(value) -> str:
    if value is None:
        return ""

    if isinstance(value, bool):
        return ""

    if isinstance(value, (int, float)):
        if value == 0:
            return ""

        if isinstance(value, float) and value.is_integer():
            return str(int(value))

        return str(value)

    text = _normalize(value)

    if not text:
        return ""

    if text in {
        "0",
        "0.0",
        "0.00",
        "0,0",
        "0,00",
        "-",
    }:
        return ""

    try:
        # "1.234,56" -> 1234.56
        numeric_text = text

        if "," in numeric_text:
            numeric_text = (
                numeric_text
                .replace(".", "")
                .replace(",", ".")
            )

        if float(numeric_text) == 0:
            return ""

    except Exception:
        pass

    return text


def _normalize_tarih(value) -> str:
    text = _normalize(value)

    if not text:
        return ""

    # ISO: 2026-04-28T10:10:41
    match = re.match(
        r"^(\d{4})-(\d{2})-(\d{2})"
        r"[T\s](\d{2}):(\d{2})",
        text,
    )

    if match:
        return (
            f"{match.group(3)}."
            f"{match.group(2)}."
            f"{match.group(1)} "
            f"{match.group(4)}:"
            f"{match.group(5)}"
        )

    # 08.05.202613:58:49
    match = re.match(
        r"^(\d{1,2}[./]\d{1,2}[./]\d{4})"
        r"(\d{1,2}:\d{2})(?::\d{2})?",
        text,
    )

    if match:
        return (
            match.group(1).replace("/", ".")
            + " "
            + match.group(2)
        )

    # 08.05.2026 13:58:49
    match = re.match(
        r"^(\d{1,2}[./]\d{1,2}[./]\d{4})"
        r"\s+(\d{1,2}:\d{2})(?::\d{2})?",
        text,
    )

    if match:
        return (
            match.group(1).replace("/", ".")
            + " "
            + match.group(2)
        )

    # 08.05.2026
    match = re.match(
        r"^(\d{1,2}[./]\d{1,2}[./]\d{4})$",
        text,
    )

    if match:
        return (
            match.group(1)
            .replace("/", ".")
        )

    return text.replace("/", ".")


def _parse_aciklama(
    raw_aciklama: str,
) -> Tuple[str, str]:
    text = _normalize(
        raw_aciklama
    )

    if not text:
        return "", ""

    match = DATE_PATTERN.search(
        text
    )

    if match:
        date = _normalize_tarih(
            match.group(1)
        )

        clean = _normalize(
            DATE_PATTERN.sub(
                "",
                text,
            )
        ).strip(" .:-")

        return clean, date

    match = DATE_PATTERN_ITIBAR.search(
        text
    )

    if match:
        return (
            text,
            _normalize_tarih(
                match.group(1)
            ),
        )

    match = DATE_PATTERN_TR.search(
        text
    )

    if match:
        day = match.group(1).zfill(2)
        month = TURKCE_AYLAR.get(
            match.group(2).lower(),
            "",
        )
        year = match.group(3)

        date = (
            f"{day}.{month}.{year}"
            if month
            else ""
        )

        clean = _normalize(
            DATE_PATTERN_TR.sub(
                "",
                text,
            )
        ).strip(" .:-")

        return clean, date

    return text, ""


# =========================================================
# TRANSFER TERİMLERİ
# =========================================================

def _has_transfer_term(
    text: str,
    term: str,
) -> bool:
    normalized = _normalize_key(
        text
    )

    patterns = {
        "fast": r"(?<![a-z0-9])fast(?![a-z0-9])",
        "eft": r"(?<![a-z0-9])eft(?![a-z0-9])",
        "havale": r"(?<![a-z0-9])havale(?![a-z0-9])",
        "swift": r"(?<![a-z0-9])swift(?![a-z0-9])",
    }

    pattern = patterns.get(
        term
    )

    if pattern:
        return re.search(
            pattern,
            normalized,
        ) is not None

    return term in normalized


# =========================================================
# GÜNCEL ANA KATEGORİLER
# =========================================================

CATEGORY_ALIASES = {
    "atm kullanim": "ATM Kullanım",
    "bireysel krediler": "Bireysel Krediler",
    "cekler ve senetler": "Çekler ve Senetler",
    "dis ticaret": "Dış Ticaret",
    "kiralik kasa ucretleri": "Kiralık Kasa Ücretleri",
    "kobi kredileri": "Ticari Krediler",
    "ticari krediler": "Ticari Krediler",
    "kredi kartlari ve banka kartlari": "Kredi Kartları ve Banka Kartları",
    "menkul kiymet islemleri": "Menkul Kıymet İşlemleri",
    "mevduat hesaplari": "Mevduat Hesapları",
    "para aktarma": "Para Aktarma",
    "uye isyeri ve pos urun ve hizmetleri azami ucretleri":
        "Üye İşyeri ve POS Ürün ve Hizmetleri Azami Ücretleri",
    "diger": "Diğer",
}

IGNORE_HEADINGS = {
    "ürün ve hizmet ücretleri",
    "urun ve hizmet ucretleri",
    "fiyatlar ve oranlar",
    "diğer faiz oranları",
    "diger faiz oranlari",
    "geçmiş dönem mevduat / kredi faiz oranları",
    "gecmis donem mevduat / kredi faiz oranlari",
    "müşteri ol",
    "musteri ol",
    "ara",
    "menü",
    "menu",
    "kapat",
    "ana sayfa",
    "anasayfa",
}


def _heading_text(
    heading,
) -> str:
    return _normalize(
        heading.get_text(
            " ",
            strip=True,
        )
    )


TRANSFER_GROUP_ALIASES = {
    "eft": "EFT",
    "fast": "FAST",
    "havale": "Havale",
    "swift": "Swift",
    "altin eft": "Altın EFT",
}


def _previous_context_texts(
    table,
    limit: int = 12000,
) -> List[str]:
    """
    Ziraat ana kategori etiketleri her zaman heading değildir.
    Tablo öncesindeki görünür metinlerden en yakın exact kategori bulunur.
    Başka ücret tablolarının hücreleri bağlamdan çıkarılır.
    """
    result: List[str] = []

    for node in table.find_all_previous(
        string=True,
        limit=limit,
    ):
        parent = getattr(
            node,
            "parent",
            None,
        )

        if parent is not None:
            name = (
                parent.name.lower()
                if getattr(parent, "name", None)
                else ""
            )

            if name in {
                "script",
                "style",
                "noscript",
                "svg",
                "path",
                "option",
            }:
                continue

            if parent.find_parent(
                "table"
            ) is not None:
                continue

        text = _normalize(
            str(node)
        )

        if (
            text
            and text not in result
        ):
            result.append(
                text
            )

    return result


def _find_category_from_text(
    table,
) -> str:
    for text in _previous_context_texts(
        table
    ):
        key = _normalize_key(
            text
        )

        if key in CATEGORY_ALIASES:
            return CATEGORY_ALIASES[
                key
            ]

    return "Genel"


def _heading_level(
    element,
) -> int:
    name = (
        element.name.lower()
        if getattr(
            element,
            "name",
            None,
        )
        else ""
    )

    match = re.fullmatch(
        r"h([1-6])",
        name,
    )

    if match:
        return int(
            match.group(1)
        )

    return 9


def _nearby_heading_elements(
    table,
    limit: int = 50,
):
    """
    En yakın başlık elemanlarını yakın -> uzak döndürür.
    """
    result = []

    for element in table.find_all_previous(
        [
            "h1",
            "h2",
            "h3",
            "h4",
            "h5",
            "h6",
            "button",
        ],
        limit=limit,
    ):
        text = _heading_text(
            element
        )

        if not text:
            continue

        key = _normalize_key(
            text
        )

        if key in {
            _normalize_key(x)
            for x in IGNORE_HEADINGS
        }:
            continue

        result.append(
            (
                element,
                text,
                _heading_level(
                    element
                ),
            )
        )

    return result


def _find_category_and_hierarchy(
    table,
) -> Tuple[str, List[str]]:
    """
    Ana kategori exact görünür metinden bulunur.

    Hiyerarşi yalnız gerçek heading parent-child ilişkisini taşır.
    Önceki sibling bölümün FAST / Swift / Altın EFT başlığı sonraki
    BKM / POS / başka tabloya sızmaz.
    """
    category = _find_category_from_text(
        table
    )

    headings = _nearby_heading_elements(
        table,
        limit=50,
    )

    local_text = ""
    local_level = 99
    local_index = -1

    for index, (
        element,
        text,
        level,
    ) in enumerate(
        headings
    ):
        key = _normalize_key(
            text
        )

        if key in CATEGORY_ALIASES:
            continue

        if len(text) > 180:
            continue

        local_text = text
        local_level = level
        local_index = index
        break

    hierarchy: List[str] = []

    if local_text:
        local_key = _normalize_key(
            local_text
        )

        if (
            category == "Para Aktarma"
            and local_key
            in TRANSFER_GROUP_ALIASES
        ):
            hierarchy.append(
                TRANSFER_GROUP_ALIASES[
                    local_key
                ]
            )

            return category, hierarchy

    if (
        category == "Para Aktarma"
        and local_index >= 0
    ):
        for (
            element,
            text,
            level,
        ) in headings[
            local_index + 1:
        ]:
            key = _normalize_key(
                text
            )

            if key in CATEGORY_ALIASES:
                break

            # Parent başlık daha üst h seviyesinde olmalı.
            # h2 parent + h3 local -> geçerli.
            # h2 sibling + h2 local -> geçersiz.
            if (
                key in TRANSFER_GROUP_ALIASES
                and level < local_level
            ):
                hierarchy.append(
                    TRANSFER_GROUP_ALIASES[
                        key
                    ]
                )
                break

    if local_text:
        local_key = _normalize_key(
            local_text
        )

        if not any(
            _normalize_key(x)
            == local_key
            for x in hierarchy
        ):
            hierarchy.append(
                local_text
            )

    return category, hierarchy


def _build_masraf(
    raw_masraf: str,
    hierarchy: List[str],
) -> str:
    raw = _normalize(
        raw_masraf
    )

    parts: List[str] = []

    for value in hierarchy:
        value = _normalize(
            value
        )

        if not value:
            continue

        value_key = _normalize_key(
            value
        )

        if any(
            _normalize_key(existing)
            == value_key
            for existing in parts
        ):
            continue

        parts.append(
            value
        )

    raw_key = _normalize_key(
        raw
    )

    # Ham satır zaten üst başlığı içeriyorsa tekrar yazma.
    if not any(
        raw_key == _normalize_key(part)
        or _normalize_key(part) in raw_key
        for part in parts
    ):
        parts.append(
            raw
        )

    masraf = (
        " - ".join(parts)
        if parts
        else raw
    )

    # Swift grubu hierarchy içinde zaten korunur.
    # Kelime sınırı kontrolü yalnız raporlama tarafında yapılır.

    return _normalize(
        masraf
    )


# =========================================================
# TABLO GRID
# =========================================================

def _root_tables(
    soup: BeautifulSoup,
):
    return [
        table
        for table in soup.find_all(
            "table"
        )
        if table.find_parent(
            "table"
        ) is None
    ]


def _table_to_rows(
    table,
) -> List[List[str]]:
    """
    rowspan / colspan içeren tabloları düz grid'e çevirir.
    """
    rows: List[List[str]] = []

    active_spans: Dict[
        int,
        Tuple[str, int],
    ] = {}

    trs = [
        tr
        for tr in table.find_all(
            "tr"
        )
        if tr.find_parent(
            "table"
        ) is table
    ]

    for tr in trs:
        cells = tr.find_all(
            ["th", "td"],
            recursive=False,
        )

        if (
            not cells
            and not active_spans
        ):
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

                row.append(
                    text
                )

                if remaining <= 1:
                    del active_spans[
                        col
                    ]
                else:
                    active_spans[
                        col
                    ] = (
                        text,
                        remaining - 1,
                    )

                col += 1
                continue

            if cell_index >= len(
                cells
            ):
                future_cols = [
                    c
                    for c in active_spans
                    if c > col
                ]

                if not future_cols:
                    break

                next_col = min(
                    future_cols
                )

                while col < next_col:
                    row.append("")
                    col += 1

                continue

            cell = cells[
                cell_index
            ]
            cell_index += 1

            # Ziraat, 200 karakterden uzun açıklamaları görünür hücrede
            # "... Devamı" şeklinde kısaltıyor. Tam resmî metin aynı td'nin
            # data-full-aciklama niteliğinde tutuluyor. Görünür metni almak
            # açıklamanın Excel'e eksik gitmesine neden olduğu için nitelik
            # varsa onu birincil kaynak kabul ediyoruz.
            full_description = cell.get(
                "data-full-aciklama"
            )

            if full_description is not None:
                text = _normalize(
                    full_description
                )
            else:
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

            for _ in range(
                colspan
            ):
                row.append(
                    text
                )

                if rowspan > 1:
                    active_spans[
                        col
                    ] = (
                        text,
                        rowspan - 1,
                    )

                col += 1

        if any(
            _normalize(value)
            for value in row
        ):
            rows.append(
                row
            )

    return rows


# =========================================================
# HEADER
# =========================================================

def _header_score(
    row: List[str],
) -> int:
    headers = [
        _normalize_key(
            value
        )
        for value in row
    ]

    tests = [
        ["masraf"],
        ["asgari", "tutar"],
        ["asgari", "oran"],
        ["azami", "tutar"],
        ["azami", "oran"],
        ["aciklama"],
        ["guncelleme"],
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
        rows[:8]
    ):
        score = _header_score(
            row
        )

        if score > best_score:
            best_score = score
            best_index = index

    if best_score < 4:
        return -1

    return best_index


def _find_col(
    headers: List[str],
    keywords: List[str],
) -> int:
    for index, header in enumerate(
        headers
    ):
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
        _normalize_key(
            value
        )
        for value in header_row
    ]

    cols = {
        "masraf": _find_col(
            headers,
            ["masraf"],
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

    # Güncel Ziraat ücret tablosu 7 kolon.
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
            if cols[
                key
            ] == -1:
                cols[
                    key
                ] = index

    return cols


def _row_is_same_header(
    row: List[str],
    header: List[str],
) -> bool:
    left = [
        _normalize_key(
            value
        )
        for value in row
    ]
    right = [
        _normalize_key(
            value
        )
        for value in header
    ]

    while (
        left
        and not left[-1]
    ):
        left.pop()

    while (
        right
        and not right[-1]
    ):
        right.pop()

    return left == right


# =========================================================
# HTML PARSER
# =========================================================

def _parse_html(
    html: str,
    source_name: str,
) -> Tuple[
    List[UcretSatiri],
    Dict[str, int],
]:
    soup = BeautifulSoup(
        html,
        "lxml",
    )

    tables = _root_tables(
        soup
    )

    stats: Dict[str, int] = {
        "root_tables": len(
            tables
        ),
        "fee_tables": 0,
        "ignored_tables": 0,
        "zero_record_tables": 0,
        "candidate_rows": 0,
        "parsed_before_dedup": 0,
        "repeated_headers": 0,
        "notes": 0,
        "invalid_rows": 0,
        "missing_category_tables": 0,
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

        header_index = (
            _find_header_index(
                grid
            )
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
            _normalize(
                value
            )
            for value in grid[
                header_index
            ]
        ]

        columns = _find_columns(
            header
        )

        (
            category,
            hierarchy,
        ) = _find_category_and_hierarchy(
            table
        )

        if category == "Genel":
            stats[
                "missing_category_tables"
            ] += 1

            print(
                f"[ziraat][DEBUG][{source_name}] "
                f"GENEL kalan tablo={table_index} | "
                f"yakın bağlam={_context_headings(table, limit=15)}",
                file=sys.stderr,
            )

        table_record_count = 0

        for raw_row in grid[
            header_index + 1:
        ]:
            row = [
                _normalize(
                    value
                )
                for value in raw_row
            ]

            if (
                not row
                or not any(row)
            ):
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

            def get(
                index: int,
            ) -> str:
                if (
                    index < 0
                    or index >= len(row)
                ):
                    return ""

                return _normalize(
                    row[index]
                )

            raw_masraf = get(
                columns[
                    "masraf"
                ]
            )

            if not raw_masraf:
                stats[
                    "invalid_rows"
                ] += 1
                continue

            asgari_tutar = (
                _normalize_tutar(
                    get(
                        columns[
                            "asgari_tutar"
                        ]
                    )
                )
            )

            asgari_oran = (
                _normalize_tutar(
                    get(
                        columns[
                            "asgari_oran"
                        ]
                    )
                )
            )

            azami_tutar = (
                _normalize_tutar(
                    get(
                        columns[
                            "azami_tutar"
                        ]
                    )
                )
            )

            azami_oran = (
                _normalize_tutar(
                    get(
                        columns[
                            "azami_oran"
                        ]
                    )
                )
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

            site_tarihi = (
                _normalize_tarih(
                    get(
                        columns[
                            "tarih"
                        ]
                    )
                )
            )

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
                hierarchy,
            )

            rows_out.append(
                UcretSatiri(
                    kategori=category,
                    masraf=masraf,
                    asgari_tutar=asgari_tutar,
                    asgari_oran=asgari_oran,
                    azami_tutar=azami_tutar,
                    azami_oran=azami_oran,
                    aciklama=aciklama,
                    site_guncelleme_tarihi=site_tarihi,
                )
            )

            stats[
                "parsed_before_dedup"
            ] += 1

            table_record_count += 1

        if table_record_count == 0:
            stats[
                "zero_record_tables"
            ] += 1

            print(
                f"[ziraat][DEBUG][{source_name}] "
                f"tablo={table_index} | "
                f"kategori={category} | "
                f"hiyerarşi={hierarchy} | "
                f"header={header}",
                file=sys.stderr,
            )

    return rows_out, stats


# =========================================================
# RESMÎ XML PARSER
# =========================================================

def _parse_xml(
    xml_content: bytes,
) -> Tuple[
    List[UcretSatiri],
    Dict[str, int],
]:
    """
    Ziraat ücret sayfasının bizzat kullandığı resmî XML'i parse eder.

    Görsel sayfa, 200 karakterden uzun açıklamaları "... Devamı" diye
    kısaltır. XML'deki Aciklama alanı ise metnin tamamını içerir. Bu yüzden
    XML birincil, oluşturulmuş HTML ise yalnız yedek kaynaktır.
    """
    root = ET.fromstring(
        xml_content
    )

    groups = root.findall(
        "IslemGrubu"
    )

    stats: Dict[str, int] = {
        "root_tables": len(groups),
        "fee_tables": 0,
        "ignored_tables": 0,
        "zero_record_tables": 0,
        "candidate_rows": 0,
        "parsed_before_dedup": 0,
        "repeated_headers": 0,
        "notes": 0,
        "invalid_rows": 0,
        "missing_category_tables": 0,
    }

    rows_out: List[
        UcretSatiri
    ] = []

    for group in groups:
        raw_category = _normalize(
            group.get(
                "IslemGrubuAdi"
            )
        )
        category = CATEGORY_ALIASES.get(
            _normalize_key(
                raw_category
            ),
            raw_category or "Genel",
        )

        if category == "Genel":
            stats[
                "missing_category_tables"
            ] += 1

        for operation in group.findall(
            "Islem"
        ):
            operation_name = _normalize(
                operation.get(
                    "IslemAdi"
                )
            )

            for item in operation.findall(
                "Kalem"
            ):
                stats[
                    "fee_tables"
                ] += 1

                item_name = _normalize(
                    item.get(
                        "KalemAdi"
                    )
                )

                hierarchy: List[str] = []

                if category == "Para Aktarma":
                    operation_key = _normalize_key(
                        operation_name
                    )
                    operation_group = (
                        TRANSFER_GROUP_ALIASES.get(
                            operation_key
                        )
                    )

                    if operation_group:
                        hierarchy.append(
                            operation_group
                        )

                if (
                    item_name
                    and not any(
                        _normalize_key(existing)
                        == _normalize_key(item_name)
                        for existing in hierarchy
                    )
                ):
                    hierarchy.append(
                        item_name
                    )

                item_record_count = 0

                for fee in item.findall(
                    "Masraf"
                ):
                    stats[
                        "candidate_rows"
                    ] += 1

                    raw_masraf = _normalize(
                        fee.get(
                            "MasrafAdi"
                        )
                    )

                    if not raw_masraf:
                        stats[
                            "invalid_rows"
                        ] += 1
                        continue

                    asgari_tutar = _normalize_tutar(
                        fee.findtext(
                            "AsgariTutar"
                        )
                    )
                    asgari_oran = _normalize_tutar(
                        fee.findtext(
                            "AsgariOran"
                        )
                    )
                    azami_tutar = _normalize_tutar(
                        fee.findtext(
                            "AzamiTutar"
                        )
                    )
                    azami_oran = _normalize_tutar(
                        fee.findtext(
                            "AzamiOran"
                        )
                    )

                    (
                        aciklama,
                        aciklama_tarihi,
                    ) = _parse_aciklama(
                        fee.findtext(
                            "Aciklama"
                        )
                        or ""
                    )

                    site_tarihi = _normalize_tarih(
                        fee.findtext(
                            "GuncellemeTarihi"
                        )
                    )

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

                    rows_out.append(
                        UcretSatiri(
                            kategori=category,
                            masraf=_build_masraf(
                                raw_masraf,
                                hierarchy,
                            ),
                            asgari_tutar=asgari_tutar,
                            asgari_oran=asgari_oran,
                            azami_tutar=azami_tutar,
                            azami_oran=azami_oran,
                            aciklama=aciklama,
                            site_guncelleme_tarihi=site_tarihi,
                        )
                    )

                    stats[
                        "parsed_before_dedup"
                    ] += 1
                    item_record_count += 1

                if item_record_count == 0:
                    stats[
                        "zero_record_tables"
                    ] += 1

    return rows_out, stats


def _scrape_xml(
    url: str = ZIRAAT_XML_URL,
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

    return _parse_xml(
        response.content
    )


# =========================================================
# ÇEKME YÖNTEMLERİ
# =========================================================

def _scrape_requests(
    url: str,
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

    return _parse_html(
        response.text,
        "requests",
    )


def _scrape_playwright(
    url: str,
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

            page.goto(
                url,
                timeout=90000,
                wait_until=(
                    "domcontentloaded"
                ),
            )

            page.wait_for_timeout(
                1000
            )

            # Gizli accordion/tab içeriği varsa topluca aç.
            try:
                page.evaluate("""
                () => {
                    const targets = new Set();

                    document.querySelectorAll(
                        "[aria-expanded='false'], "
                        "[data-bs-toggle='collapse'], "
                        "[data-toggle='collapse'], "
                        ".accordion-button, "
                        "[role='button']"
                    ).forEach(el => targets.add(el));

                    for (const el of targets) {
                        try {
                            el.click();
                        } catch (_) {}
                    }
                }
                """)
            except Exception:
                pass

            page.wait_for_timeout(
                700
            )

            # Lazy load ihtimaline karşı kısa scroll.
            for _ in range(50):
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
                        "(window.innerHeight + window.scrollY) >= "
                        "((document.documentElement "
                        "? document.documentElement.scrollHeight "
                        ": 0) - 120)"
                    )

                    if bottom:
                        break

                except Exception:
                    break

            html = page.content()

        finally:
            browser.close()

    return _parse_html(
        html,
        "playwright",
    )


# =========================================================
# DEDUP
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

        seen.add(
            key
        )

        result.append(
            row
        )

    stats[
        "duplicates"
    ] = duplicates

    return result


# =========================================================
# RAPORLAR
# =========================================================

def _print_category_report(
    rows: List[UcretSatiri],
) -> None:
    counts: Dict[
        str,
        int,
    ] = {}

    for row in rows:
        counts[
            row.kategori
        ] = (
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
        "[ziraat] ===== KATEGORİ RAPORU =====",
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
            f"[ziraat] "
            f"{category} -> "
            f"{count} kayıt",
            file=sys.stderr,
        )

    print(
        "[ziraat] ===========================",
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
        "[ziraat] ===== BÜTÜNLÜK KONTROLÜ =====",
        file=sys.stderr,
    )

    print(
        f"[ziraat] Toplam root tablo: "
        f"{stats.get('root_tables', 0)}",
        file=sys.stderr,
    )

    print(
        f"[ziraat] Ücret tablosu: "
        f"{stats.get('fee_tables', 0)}",
        file=sys.stderr,
    )

    print(
        f"[ziraat] İlgisiz/atlanan tablo: "
        f"{stats.get('ignored_tables', 0)}",
        file=sys.stderr,
    )

    print(
        f"[ziraat] Ham ücret satırı adayı: "
        f"{candidate_rows}",
        file=sys.stderr,
    )

    print(
        f"[ziraat] Parse edilen "
        f"(dedup öncesi): "
        f"{stats.get('parsed_before_dedup', 0)}",
        file=sys.stderr,
    )

    print(
        f"[ziraat] Duplicate: "
        f"{stats.get('duplicates', 0)}",
        file=sys.stderr,
    )

    print(
        f"[ziraat] Tekrarlanan header: "
        f"{stats.get('repeated_headers', 0)}",
        file=sys.stderr,
    )

    print(
        f"[ziraat] Not / tek hücreli satır: "
        f"{stats.get('notes', 0)}",
        file=sys.stderr,
    )

    print(
        f"[ziraat] Geçersiz / boş veri satırı: "
        f"{stats.get('invalid_rows', 0)}",
        file=sys.stderr,
    )

    print(
        f"[ziraat] 0 kayıt üreten ücret tablosu: "
        f"{stats.get('zero_record_tables', 0)}",
        file=sys.stderr,
    )

    print(
        f"[ziraat] Kategorisi Genel kalan ücret tablosu: "
        f"{stats.get('missing_category_tables', 0)}",
        file=sys.stderr,
    )

    print(
        f"[ziraat] Excel'e gidecek "
        f"benzersiz satır: "
        f"{result_count}",
        file=sys.stderr,
    )

    if (
        candidate_rows == explained
        and stats.get(
            "zero_record_tables",
            0,
        ) == 0
        and stats.get(
            "missing_category_tables",
            0,
        ) == 0
        and stats.get(
            "fee_tables",
            0,
        ) > 0
    ):
        print(
            "[ziraat] BÜTÜNLÜK: OK - "
            "tüm tanınan ücret tabloları "
            "kategoriye bağlandı ve aday "
            "satırların tamamı açıklandı.",
            file=sys.stderr,
        )
    else:
        print(
            "[ziraat] BÜTÜNLÜK: UYARI",
            file=sys.stderr,
        )

    print(
        "[ziraat] ===============================",
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
        "[ziraat] ===== PARA AKTARMA KONTROLÜ =====",
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
            f"[ziraat] {label}: "
            f"{len(found)} kayıt",
            file=sys.stderr,
        )

        for row in found[:15]:
            print(
                f"    - "
                f"[{row.kategori}] "
                f"{row.masraf}",
                file=sys.stderr,
            )

        if not found:
            print(
                f"[ziraat][UYARI] "
                f"{label} MASRAF alanında "
                "hiç bulunamadı.",
                file=sys.stderr,
            )

    print(
        "[ziraat] =================================",
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
        "[ziraat] ===== EK ÜRÜN KONTROLÜ =====",
        file=sys.stderr,
    )

    for label, term in [
        ("Fatura", "fatura"),
        ("HGS", "hgs"),
        ("Kiralık Kasa", "kiralik kasa"),
        ("POS", "pos"),
    ]:
        found = [
            row
            for row in rows
            if (
                term
                in _normalize_key(
                    row.masraf
                )
                or term
                in _normalize_key(
                    row.kategori
                )
            )
        ]

        print(
            f"[ziraat] {label}: "
            f"{len(found)} kayıt",
            file=sys.stderr,
        )

        for row in found[:10]:
            print(
                f"    - "
                f"[{row.kategori}] "
                f"{row.masraf}",
                file=sys.stderr,
            )

    print(
        "[ziraat] =============================",
        file=sys.stderr,
    )


def _print_context_examples(
    rows: List[UcretSatiri],
) -> None:
    print(
        "",
        file=sys.stderr,
    )

    print(
        "[ziraat] ===== MASRAF ÖRNEKLERİ =====",
        file=sys.stderr,
    )

    # Transfer ve birkaç genel örnek.
    selected: List[
        UcretSatiri
    ] = []

    for row in rows:
        if (
            row.kategori
            == "Para Aktarma"
            and len(selected) < 20
        ):
            selected.append(
                row
            )

    if len(selected) < 20:
        for row in rows:
            if row in selected:
                continue

            selected.append(
                row
            )

            if len(selected) >= 20:
                break

    for row in selected:
        print(
            f"    - "
            f"[{row.kategori}] "
            f"{row.masraf}",
            file=sys.stderr,
        )

    print(
        "[ziraat] ===========================",
        file=sys.stderr,
    )


# =========================================================
# ANA SCRAPER
# =========================================================

def scrape_ziraat(
    url: str = ZIRAAT_URL,
) -> List[UcretSatiri]:
    print(
        f"[ziraat] SÜRÜM: "
        f"{SCRAPER_VERSION}",
        file=sys.stderr,
    )

    print(
        f"[ziraat] "
        f"{url} adresinden "
        "veri çekiliyor...",
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
    # RESMÎ XML - birincil ve tam açıklamalı yol
    # -----------------------------------------------------

    try:
        xml_rows, xml_stats = (
            _scrape_xml()
        )

        print(
            f"[ziraat][xml] "
            f"grup={xml_stats['root_tables']}, "
            f"ücret_tablosu={xml_stats['fee_tables']}, "
            f"aday={xml_stats['candidate_rows']}, "
            f"parse={len(xml_rows)}",
            file=sys.stderr,
        )

        if (
            xml_stats.get(
                "fee_tables",
                0,
            ) >= 30
            and xml_stats.get(
                "candidate_rows",
                0,
            ) >= 120
            and len(
                xml_rows
            ) >= 120
        ):
            rows = xml_rows
            stats = xml_stats
            source = "official-xml"
        else:
            print(
                "[ziraat][UYARI] "
                "Resmî XML sonucu eksik görünüyor; "
                "HTML yolları denenecek.",
                file=sys.stderr,
            )

    except Exception as exc:
        print(
            f"[ziraat][UYARI] "
            f"Resmî XML başarısız: {exc}",
            file=sys.stderr,
        )

    # -----------------------------------------------------
    # REQUESTS - birincil ve hızlı yol
    # -----------------------------------------------------

    if not rows:
        try:
            request_rows, request_stats = (
                _scrape_requests(
                    url
                )
            )

            print(
                f"[ziraat][requests] "
                f"root={request_stats['root_tables']}, "
                f"ücret={request_stats['fee_tables']}, "
                f"aday={request_stats['candidate_rows']}, "
                f"parse={len(request_rows)}",
                file=sys.stderr,
            )

            # Güncel Ziraat sayfası çok sayıda server-rendered ücret tablosu
            # içeriyor. Küçük bir sonuç gelirse eksik HTML varsay.
            if (
                request_stats.get(
                    "fee_tables",
                    0,
                ) >= 30
                and request_stats.get(
                    "candidate_rows",
                    0,
                ) >= 120
                and len(
                    request_rows
                ) >= 120
            ):
                rows = request_rows
                stats = request_stats
                source = "requests"

            else:
                print(
                    "[ziraat][UYARI] "
                    "requests sonucu eksik görünüyor; "
                    "Playwright fallback çalışacak.",
                    file=sys.stderr,
                )

        except Exception as exc:
            print(
                f"[ziraat][UYARI] "
                f"requests başarısız: {exc}",
                file=sys.stderr,
            )

    # -----------------------------------------------------
    # PLAYWRIGHT - yalnız gerektiğinde
    # -----------------------------------------------------

    if not rows:
        playwright_rows, playwright_stats = (
            _scrape_playwright(
                url
            )
        )

        print(
            f"[ziraat][playwright] "
            f"root={playwright_stats['root_tables']}, "
            f"ücret={playwright_stats['fee_tables']}, "
            f"aday={playwright_stats['candidate_rows']}, "
            f"parse={len(playwright_rows)}",
            file=sys.stderr,
        )

        rows = playwright_rows
        stats = playwright_stats
        source = "playwright"

    if (
        not rows
        or stats is None
    ):
        raise ScraperError(
            "Ziraat sayfasından "
            "hiçbir ücret satırı "
            "çekilemedi."
        )

    rows = _deduplicate(
        rows,
        stats,
    )

    truncated_descriptions = [
        row
        for row in rows
        if re.search(
            r"(?:\.\.\.|…)[\s\xa0]*Devam[ıi]\s*$",
            row.aciklama or "",
            flags=re.IGNORECASE,
        )
    ]

    if truncated_descriptions:
        examples = "; ".join(
            row.masraf
            for row in truncated_descriptions[:5]
        )
        raise ScraperError(
            "Ziraat tam açıklama kontrolü başarısız: "
            f"{len(truncated_descriptions)} satır hâlâ '... Devamı' ile "
            f"bitiyor. Örnekler: {examples}"
        )

    print(
        "[ziraat] Tam açıklama kontrolü: OK - "
        "'... Devamı' ile kesilmiş açıklama yok.",
        file=sys.stderr,
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
        f"[ziraat] Kullanılan kaynak: "
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

    _print_context_examples(
        rows
    )

    print(
        f"[ziraat] Toplam "
        f"{len(rows)} benzersiz "
        "satır bulundu.",
        file=sys.stderr,
    )

    return rows


# =========================================================
# TEST
# =========================================================

if __name__ == "__main__":
    try:
        veriler = scrape_ziraat()

        print()
        print("=" * 70)
        print("ZİRAAT SCRAPER")
        print("=" * 70)
        print(
            f"Toplam çekilen ücret: "
            f"{len(veriler)}"
        )
        print()

        for i, row in enumerate(
            veriler[:50],
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
            f"[ziraat][HATA] "
            f"{exc}",
            file=sys.stderr,
        )
        sys.exit(1)
