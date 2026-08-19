"""
TEB "Ürün ve Hizmet Ücretleri" scraper.

Bu sürüm:
- Bireysel Ürün ve Hizmet Ücretleri sayfasını çeker.
- Ana sayfada ayrıca bağlantısı verilen Ticari/Tüzel Ürün ve Hizmet Ücretleri
  sayfasını da çeker.
- requests'i birincil kullanır; sayfa eksik görünürse Playwright fallback yapar.
- Nested table yapısında yalnız leaf tabloları parse eder.
- Gerçek ücret tablosunu header yapısından tanır.
- Tablo başlığını MASRAF alanında korur:
    ATM'den EFT Ücreti - 8.300 TL ve Altı
    FAST Ücreti - 8.300 TL ve Altı
    Havale - Şube - ...
- "Uluslararası Fon Transferi ve Mesajlaşma Ücreti" başlığını SWIFT olarak
  görünür hale getirir.
- Bireysel ve ticari sayfaların bütünlüğünü ayrı ayrı ve toplam olarak raporlar.
- Duplicate, tekrarlanan header, not/tek hücreli satır ve geçersiz satırları sayar.
- FAST / EFT / Havale / SWIFT kontrolü yapar.
- Fatura / HGS / Kiralık Kasa ek kontrolü yapar.
- main.py / update_excel.py ile uyumlu UcretSatiri listesi döndürür.
"""

import re
import sys
from dataclasses import dataclass
from typing import Dict, List, Optional, Set, Tuple

import requests
from bs4 import BeautifulSoup


SCRAPER_VERSION = "2026-08-19-v2-teb-bireysel-ticari-integrity"

TEB_URL = "https://www.teb.com.tr/urun-ve-hizmet-ucretleri/"
TEB_TICARI_URL = "https://www.teb.com.tr/tuzel-urun-ve-hizmet-ucretleri/"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "tr-TR,tr;q=0.9,en;q=0.8",
}

DATE_PATTERN = re.compile(
    r"G[üu]ncell[ei]nme\s*Tarihi\s*:\s*"
    r"(\d{1,2}[./]\d{1,2}[./]\d{4}"
    r"(?:\s+\d{1,2}:\d{2})?)",
    re.IGNORECASE,
)

DATE_VALUE_PATTERN = re.compile(
    r"^\d{1,2}[./]\d{1,2}[./]\d{4}"
    r"(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?$"
)


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

    text = re.sub(r"\s+", " ", text)
    return text.strip()


def _normalize_date(value: str) -> str:
    text = _normalize(value).replace("/", ".")

    if not text:
        return ""

    # dd.mm.yyyy HH:MM:SS -> saniyeyi kaldır.
    match = re.match(
        r"^(\d{1,2}\.\d{1,2}\.\d{4})"
        r"\s+(\d{1,2}:\d{2}):\d{2}$",
        text,
    )

    if match:
        return f"{match.group(1)} {match.group(2)}"

    return text


def _parse_aciklama(raw: str) -> Tuple[str, str]:
    text = _normalize(raw)

    if not text:
        return "", ""

    match = DATE_PATTERN.search(text)

    if not match:
        return text, ""

    date = _normalize_date(
        match.group(1)
    )

    clean = _normalize(
        DATE_PATTERN.sub(
            "",
            text,
        )
    ).strip(" .:-")

    return clean, date


# =========================================================
# TRANSFER KONTROLÜ
# =========================================================

def _has_transfer_term(
    text: str,
    term: str,
) -> bool:
    normalized = _normalize_key(text)

    patterns = {
        "fast": r"(?<![a-z0-9])fast(?![a-z0-9])",
        "eft": r"(?<![a-z0-9])eft(?![a-z0-9])",
        "havale": r"(?<![a-z0-9])havale(?![a-z0-9])",
        "swift": r"(?<![a-z0-9])swift(?![a-z0-9])",
    }

    pattern = patterns.get(term)

    if pattern:
        return re.search(
            pattern,
            normalized,
        ) is not None

    return term in normalized


# =========================================================
# TABLO GRID
# =========================================================

def _is_leaf_table(table) -> bool:
    return not bool(
        table.find(
            "table"
        )
    )


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
                row.append(text)

                if rowspan > 1:
                    active_spans[
                        col
                    ] = (
                        text,
                        rowspan - 1,
                    )

                col += 1

        if any(
            _normalize(x)
            for x in row
        ):
            rows.append(row)

    return rows


# =========================================================
# HEADER
# =========================================================

def _header_score(
    row: List[str],
) -> int:
    headers = [
        _normalize_key(x)
        for x in row
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

    # TEB ücret tablolarında en az Masraf + Asgari/Azami + Açıklama/Tarih
    # kolonları tanınmalı.
    if best_score < 3:
        return -1

    return best_index


def _find_col(
    headers: List[str],
    keywords: List[str],
) -> int:
    for i, header in enumerate(
        headers
    ):
        if all(
            keyword in header
            for keyword in keywords
        ):
            return i

    return -1


def _find_columns(
    header_row: List[str],
) -> Dict[str, int]:
    headers = [
        _normalize_key(x)
        for x in header_row
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

    # Güncel TEB tabloları standart 7 kolon.
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
            if cols[key] == -1:
                cols[key] = index

    return cols


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
# BAŞLIK / KATEGORİ
# =========================================================

IGNORE_HEADINGS = {
    "ürün ve hizmet ücretleri",
    "ticari ürün ve hizmet ücretleri",
    "kredi kartı masraf komisyon ve ücret listesi",
    "temel bankacılık ürün bilgilendirme formu",
    "müşteri ol",
    "ara",
    "kapat",
    "menü",
    "ana sayfa",
}

BIREYSEL_CATEGORY_ALIASES = {
    "atm kullanım": "ATM Kullanım",
    "bireysel krediler": "Bireysel Krediler",
    "kredi kartları": "Kredi Kartları ve Banka Kartları",
    "kredi kartları ve banka kartları": "Kredi Kartları ve Banka Kartları",
    "banka kartları": "Kredi Kartları ve Banka Kartları",
    "mevduat": "Mevduat Hesapları",
    "mevduat hesapları": "Mevduat Hesapları",
    "menkul kıymet işlemleri": "Menkul Kıymet İşlemleri",
    "çekler ve senetler": "Çekler ve Senetler",
    "diger": "Diğer",
    "diğer": "Diğer",
    "para aktarma": "Para Aktarma",
}

TICARI_CATEGORY_ALIASES = {
    "anlasmali kurumlar": "Ticari - Anlaşmalı Kurumlar",
    "anlaşmalı kurumlar": "Ticari - Anlaşmalı Kurumlar",
    "dis ticaret": "Ticari - Dış Ticaret",
    "dış ticaret": "Ticari - Dış Ticaret",
    "diger": "Ticari - Diğer",
    "diğer": "Ticari - Diğer",
    "krediler": "Ticari - Krediler",
    "nakit yonetimi": "Ticari - Nakit Yönetimi",
    "nakit yönetimi": "Ticari - Nakit Yönetimi",
    "para aktarma": "Ticari - Para Aktarma",
    "cekler ve senetler": "Ticari - Çekler ve Senetler",
    "çekler ve senetler": "Ticari - Çekler ve Senetler",
}


def _heading_text(
    element,
) -> str:
    return _normalize(
        element.get_text(
            " ",
            strip=True,
        )
    )


def _heading_candidates(
    table,
) -> List[str]:
    """
    Tabloya en yakın başlıkları yakın -> uzak sırasıyla döndürür.
    """
    candidates: List[str] = []

    for heading in table.find_all_previous(
        ["h1", "h2", "h3", "h4", "h5", "h6"],
        limit=20,
    ):
        text = _heading_text(
            heading
        )

        if not text:
            continue

        if _normalize_key(
            text
        ) in {
            _normalize_key(x)
            for x in IGNORE_HEADINGS
        }:
            continue

        if text not in candidates:
            candidates.append(
                text
            )

    return candidates


def _find_category(
    table,
    segment: str,
) -> str:
    candidates = _heading_candidates(
        table
    )

    aliases = (
        BIREYSEL_CATEGORY_ALIASES
        if segment == "bireysel"
        else TICARI_CATEGORY_ALIASES
    )

    for text in candidates:
        key = _normalize_key(
            text
        )

        if key in aliases:
            return aliases[key]

    if segment == "ticari":
        return "Ticari - Diğer"

    return "Genel"


def _find_table_title(
    table,
    category: str,
) -> str:
    """
    En yakın başlık genellikle ürün/kanal/tablo başlığıdır.
    Kategori başlığını başlık olarak tekrar etme.
    """
    for text in _heading_candidates(
        table
    ):
        if _normalize_key(
            text
        ) == _normalize_key(
            category.replace(
                "Ticari - ",
                "",
            )
        ):
            continue

        # Ana bölüm başlıkları MASRAF prefix'i olmasın.
        if (
            _normalize_key(text)
            in BIREYSEL_CATEGORY_ALIASES
            or _normalize_key(text)
            in TICARI_CATEGORY_ALIASES
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

    parts: List[str] = []

    if table_title:
        parts.append(
            table_title
        )

    raw_key = _normalize_key(
        raw_masraf
    )

    if not any(
        _normalize_key(part)
        == raw_key
        or raw_key in _normalize_key(part)
        for part in parts
    ):
        parts.append(
            raw_masraf
        )

    masraf = " - ".join(
        parts
    ) if parts else raw_masraf

    combined = _normalize_key(
        " ".join(
            [
                table_title,
                raw_masraf,
            ]
        )
    )

    # Resmî sayfada SWIFT kalemleri "Uluslararası Fon Transferi ve
    # Mesajlaşma Ücreti" başlığında yayınlanıyor.
    if (
        (
            "uluslararasi fon transfer" in combined
            or "yurtdisi doviz transfer" in combined
            or "yurt disi doviz transfer" in combined
        )
        and not _has_transfer_term(
            masraf,
            "swift",
        )
    ):
        masraf = (
            f"SWIFT - {masraf}"
        )

    return _normalize(
        masraf
    )


# =========================================================
# TEK SAYFA PARSER
# =========================================================

def _parse_page(
    html: str,
    segment: str,
    source_name: str,
) -> Tuple[
    List[UcretSatiri],
    Dict[str, int],
]:
    soup = BeautifulSoup(
        html,
        "lxml",
    )

    all_tables = soup.find_all(
        "table"
    )

    leaf_tables = [
        table
        for table in all_tables
        if _is_leaf_table(
            table
        )
    ]

    stats: Dict[str, int] = {
        "all_tables": len(
            all_tables
        ),
        "leaf_tables": len(
            leaf_tables
        ),
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
        leaf_tables
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

        category = _find_category(
            table,
            segment,
        )

        table_title = _find_table_title(
            table,
            category,
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

            site_tarihi = _normalize_date(
                get(
                    columns[
                        "tarih"
                    ]
                )
            )

            if not site_tarihi:
                site_tarihi = (
                    aciklama_tarihi
                )

            # Son hücre tarihse ama header mapping bir nedenle kaçtıysa.
            if (
                not site_tarihi
                and row
                and DATE_VALUE_PATTERN.match(
                    _normalize(
                        row[-1]
                    )
                )
            ):
                site_tarihi = (
                    _normalize_date(
                        row[-1]
                    )
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
                f"[teb][DEBUG][{source_name}] "
                f"segment={segment} | "
                f"tablo={table_index} | "
                f"kategori={category} | "
                f"başlık={table_title or '-'} | "
                f"header={header}",
                file=sys.stderr,
            )

    return rows_out, stats


# =========================================================
# SAYFA ÇEKME
# =========================================================

def _fetch_requests(
    url: str,
    session: requests.Session,
) -> str:
    response = session.get(
        url,
        headers=HEADERS,
        timeout=45,
    )
    response.raise_for_status()
    return response.text


def _fetch_playwright(
    url: str,
) -> str:
    from playwright.sync_api import (
        sync_playwright,
    )

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
                1200
            )

            # Kapalı accordion/tab varsa toplu aç.
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

            return page.content()

        finally:
            browser.close()


def _scrape_one_page(
    url: str,
    segment: str,
    session: requests.Session,
) -> Tuple[
    List[UcretSatiri],
    Dict[str, int],
    str,
]:
    try:
        html = _fetch_requests(
            url,
            session,
        )

        rows, stats = _parse_page(
            html,
            segment,
            "requests",
        )

        print(
            f"[teb][requests] "
            f"{segment}: "
            f"all={stats['all_tables']}, "
            f"leaf={stats['leaf_tables']}, "
            f"ücret={stats['fee_tables']}, "
            f"satır={len(rows)}",
            file=sys.stderr,
        )

        # TEB ücret sayfaları server-rendered. Çok küçük sonuç gelirse
        # requests cevabının eksik olduğu kabul edilir.
        if (
            stats["fee_tables"] >= 5
            and stats["candidate_rows"] >= 15
            and len(rows) >= 15
        ):
            return rows, stats, "requests"

        print(
            f"[teb][UYARI] "
            f"{segment} requests sonucu "
            "eksik görünüyor; "
            "Playwright fallback çalışacak.",
            file=sys.stderr,
        )

    except Exception as exc:
        print(
            f"[teb][UYARI] "
            f"{segment} requests başarısız: "
            f"{exc}",
            file=sys.stderr,
        )

    html = _fetch_playwright(
        url
    )

    rows, stats = _parse_page(
        html,
        segment,
        "playwright",
    )

    print(
        f"[teb][playwright] "
        f"{segment}: "
        f"all={stats['all_tables']}, "
        f"leaf={stats['leaf_tables']}, "
        f"ücret={stats['fee_tables']}, "
        f"satır={len(rows)}",
        file=sys.stderr,
    )

    return rows, stats, "playwright"


# =========================================================
# DEDUP
# =========================================================

def _deduplicate(
    rows: List[UcretSatiri],
) -> Tuple[
    List[UcretSatiri],
    int,
]:
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

    return result, duplicates


# =========================================================
# RAPORLAR
# =========================================================

def _print_page_report(
    label: str,
    stats: Dict[str, int],
    source: str,
    unique_count: int,
    duplicates: int,
) -> None:
    candidate = stats.get(
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
        f"[teb] ===== {label.upper()} "
        "BÜTÜNLÜK =====",
        file=sys.stderr,
    )

    print(
        f"[teb] Kaynak: {source}",
        file=sys.stderr,
    )

    print(
        f"[teb] Toplam tablo: "
        f"{stats.get('all_tables', 0)}",
        file=sys.stderr,
    )

    print(
        f"[teb] Leaf tablo: "
        f"{stats.get('leaf_tables', 0)}",
        file=sys.stderr,
    )

    print(
        f"[teb] Ücret tablosu: "
        f"{stats.get('fee_tables', 0)}",
        file=sys.stderr,
    )

    print(
        f"[teb] İlgisiz/atlanan tablo: "
        f"{stats.get('ignored_tables', 0)}",
        file=sys.stderr,
    )

    print(
        f"[teb] Ham ücret satırı adayı: "
        f"{candidate}",
        file=sys.stderr,
    )

    print(
        f"[teb] Parse edilen "
        f"(dedup öncesi): "
        f"{stats.get('parsed_before_dedup', 0)}",
        file=sys.stderr,
    )

    print(
        f"[teb] Duplicate: "
        f"{duplicates}",
        file=sys.stderr,
    )

    print(
        f"[teb] Tekrarlanan header: "
        f"{stats.get('repeated_headers', 0)}",
        file=sys.stderr,
    )

    print(
        f"[teb] Not / tek hücreli satır: "
        f"{stats.get('notes', 0)}",
        file=sys.stderr,
    )

    print(
        f"[teb] Geçersiz / boş veri satırı: "
        f"{stats.get('invalid_rows', 0)}",
        file=sys.stderr,
    )

    print(
        f"[teb] 0 kayıt üreten ücret tablosu: "
        f"{stats.get('zero_record_tables', 0)}",
        file=sys.stderr,
    )

    print(
        f"[teb] Benzersiz satır: "
        f"{unique_count}",
        file=sys.stderr,
    )

    if (
        candidate == explained
        and stats.get(
            "zero_record_tables",
            0,
        ) == 0
        and stats.get(
            "fee_tables",
            0,
        ) > 0
    ):
        print(
            f"[teb] {label} BÜTÜNLÜK: OK",
            file=sys.stderr,
        )
    else:
        print(
            f"[teb] {label} BÜTÜNLÜK: UYARI",
            file=sys.stderr,
        )

    print(
        "[teb] =============================",
        file=sys.stderr,
    )


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
        "[teb] ===== KATEGORİ RAPORU =====",
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
            f"[teb] "
            f"{category} -> "
            f"{count} kayıt",
            file=sys.stderr,
        )

    print(
        "[teb] ===========================",
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
        "[teb] ===== PARA AKTARMA KONTROLÜ =====",
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
            f"[teb] {label}: "
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
                f"[teb][UYARI] "
                f"{label} MASRAF alanında "
                "hiç bulunamadı.",
                file=sys.stderr,
            )

    print(
        "[teb] =================================",
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
        "[teb] ===== EK ÜRÜN KONTROLÜ =====",
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
            f"[teb] {label}: "
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
        "[teb] =============================",
        file=sys.stderr,
    )


# =========================================================
# ANA SCRAPER
# =========================================================

def scrape_teb(
    url: str = TEB_URL,
) -> List[UcretSatiri]:
    print(
        f"[teb] SÜRÜM: "
        f"{SCRAPER_VERSION}",
        file=sys.stderr,
    )

    session = requests.Session()
    session.headers.update(
        HEADERS
    )

    # -----------------------------------------------------
    # BİREYSEL
    # -----------------------------------------------------

    print(
        f"[teb] Bireysel sayfa: "
        f"{url}",
        file=sys.stderr,
    )

    (
        bireysel_raw,
        bireysel_stats,
        bireysel_source,
    ) = _scrape_one_page(
        url,
        "bireysel",
        session,
    )

    bireysel_rows, bireysel_dup = (
        _deduplicate(
            bireysel_raw
        )
    )

    # -----------------------------------------------------
    # TİCARİ / TÜZEL
    # -----------------------------------------------------

    print(
        f"[teb] Ticari sayfa: "
        f"{TEB_TICARI_URL}",
        file=sys.stderr,
    )

    (
        ticari_raw,
        ticari_stats,
        ticari_source,
    ) = _scrape_one_page(
        TEB_TICARI_URL,
        "ticari",
        session,
    )

    ticari_rows, ticari_dup = (
        _deduplicate(
            ticari_raw
        )
    )

    if not bireysel_rows:
        raise ScraperError(
            "TEB bireysel ücret sayfasından "
            "hiç kayıt çekilemedi."
        )

    if not ticari_rows:
        raise ScraperError(
            "TEB ticari/tüzel ücret sayfasından "
            "hiç kayıt çekilemedi."
        )

    _print_page_report(
        "Bireysel",
        bireysel_stats,
        bireysel_source,
        len(bireysel_rows),
        bireysel_dup,
    )

    _print_page_report(
        "Ticari",
        ticari_stats,
        ticari_source,
        len(ticari_rows),
        ticari_dup,
    )

    combined_raw = (
        bireysel_rows
        + ticari_rows
    )

    rows, cross_duplicates = (
        _deduplicate(
            combined_raw
        )
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
        "",
        file=sys.stderr,
    )

    print(
        "[teb] ===== TOPLAM BÜTÜNLÜK =====",
        file=sys.stderr,
    )

    print(
        f"[teb] Bireysel benzersiz: "
        f"{len(bireysel_rows)}",
        file=sys.stderr,
    )

    print(
        f"[teb] Ticari benzersiz: "
        f"{len(ticari_rows)}",
        file=sys.stderr,
    )

    print(
        f"[teb] Sayfalar arası duplicate: "
        f"{cross_duplicates}",
        file=sys.stderr,
    )

    print(
        f"[teb] Excel'e gidecek "
        f"toplam benzersiz: "
        f"{len(rows)}",
        file=sys.stderr,
    )

    print(
        "[teb] ===========================",
        file=sys.stderr,
    )

    _print_category_report(
        rows
    )

    _print_transfer_report(
        rows
    )

    _print_extra_product_report(
        rows
    )

    print(
        f"[teb] Toplam "
        f"{len(rows)} benzersiz satır bulundu.",
        file=sys.stderr,
    )

    return rows


# =========================================================
# TEST
# =========================================================

if __name__ == "__main__":
    try:
        veriler = scrape_teb()

        print()
        print("=" * 70)
        print("TEB SCRAPER")
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
            f"[teb][HATA] "
            f"{exc}",
            file=sys.stderr,
        )
        sys.exit(1)
