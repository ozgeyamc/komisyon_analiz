"""
Akbank "Ürün ve Hizmet Ücretleri" scraper.

Bu sürüm:
- Önce requests ile server-rendered tüm ücret tablolarını toplar.
- requests sonucu yetersiz görünürse Playwright fallback kullanır.
- Root tabloları işler; rowspan/colspan hücrelerini normalize eder.
- Ana kategori + tablo başlığı + satır adını koruyarak MASRAF üretir.
- EFT / FAST / Havale / SWIFT gibi başlıkların Excel filtresinde kaybolmasını önler.
- Duplicate kayıtları kontrollü temizler.
- GitHub Actions loguna kategori, bütünlük ve para aktarma kontrolü yazar.
- main.py / update_excel.py ile uyumlu UcretSatiri listesi döndürür.
"""

import re
import sys
from dataclasses import dataclass
from typing import Dict, List, Optional, Set, Tuple

import requests
from bs4 import BeautifulSoup, NavigableString


SCRAPER_VERSION = "2026-08-19-v4-akbank-context-fix"

AKBANK_URL = "https://www.akbank.com/urun-ve-hizmet-ucretleri"

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
    r"G[üu]ncellenme\s*Tarihi\s*:\s*"
    r"(\d{1,2}[./]\d{1,2}[./]\d{4}(?:\s+\d{1,2}:\d{2})?)",
    re.IGNORECASE,
)

DATE_PATTERN_TR = re.compile(
    r"(?:son\s+)?g[üu]ncellenme\s+tarihi\s*:?\s*"
    r"(\d{1,2})\s+"
    r"(ocak|şubat|mart|nisan|mayıs|haziran|temmuz|ağustos|eylül|ekim|kasım|aralık)\s+"
    r"(\d{4})",
    re.IGNORECASE,
)

DATE_PATTERN_ITIBAR = re.compile(
    r"(\d{1,2}[./]\d{1,2}[./]\d{4})\s+tarihi\s+itibar",
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
        temiz = _normalize(DATE_PATTERN.sub("", raw)).strip(" .:-")
        return temiz, tarih

    match_tr = DATE_PATTERN_TR.search(raw)
    if match_tr:
        gun = match_tr.group(1).zfill(2)
        ay = TURKCE_AYLAR.get(match_tr.group(2).lower(), "")
        yil = match_tr.group(3)

        if ay:
            tarih = f"{gun}.{ay}.{yil}"
            temiz = _normalize(DATE_PATTERN_TR.sub("", raw)).strip(" .:-")
            return temiz, tarih

    match_itibar = DATE_PATTERN_ITIBAR.search(raw)
    if match_itibar:
        tarih = match_itibar.group(1).replace("/", ".").strip()
        # Bilgi içeren açıklamayı silme.
        return raw, tarih

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
    Düz substring yerine kelime sınırı kullanır.
    Böylece örneğin 'defteri' yanlışlıkla EFT sayılmaz.
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
# KATEGORİ / TABLO BAŞLIĞI
# =========================================================

INVALID_TITLES = {
    "müşteri ol",
    "ara",
    "kapat",
    "menü",
    "ana sayfa",
    "bireysel",
    "kurumsal",
    "dijital bankacılık",
    "kampanyalar",
    "güvenlik",
    "ürün ve hizmet ücretleri",
    "şube & atm'ler",
    "masraf",
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

    if len(text) < 3 or len(text) > 220:
        return False

    return _normalize_key(text) not in {
        _normalize_key(x)
        for x in INVALID_TITLES
    }


def _direct_text(node) -> str:
    if not node:
        return ""

    pieces = []

    for child in node.find_all(string=True, recursive=False):
        value = _normalize(child)
        if value:
            pieces.append(value)

    return _normalize(" ".join(pieces))


def _nearest_text_between_tables(table) -> str:
    """
    Sadece mevcut tablo ile bir önceki tablo arasındaki metne bakar.

    Böylece bir önceki transfer tablosunun başlığı sonraki,
    ilgisiz tabloya taşınmaz.

    Örn.:
      "Akbank Fast Uluslararası ..." başlığı yalnızca kendi
      "Tüm Tutarlar İçin" tablosuna bağlanır; daha sonra gelen
      "Posta İle Aylık Hesap Özeti" tablosuna sızmaz.
    """
    node = table.previous_element
    visited = 0

    while node is not None and visited < 250:
        visited += 1

        # Bir önceki tabloya ulaştık: daha geriye gitme.
        if getattr(node, "name", None) == "table":
            break

        if isinstance(node, NavigableString):
            parent = getattr(node, "parent", None)

            # Önceki tablonun içindeki text node'larını alma.
            if parent is not None and parent.find_parent("table") is not None:
                node = node.previous_element
                continue

            text = _normalize(str(node))

            if (
                _is_valid_title(text)
                and len(text) <= 140
                and len(text.split()) <= 18
            ):
                key = _normalize_key(text)

                # Tarih, kolon başlığı ve açıklama benzeri metinleri ele.
                bad_tokens = (
                    "asgari tutar",
                    "asgari oran",
                    "azami tutar",
                    "azami oran",
                    "guncelleme tarihi",
                    "bsmv",
                    "kkdf",
                )

                if not any(token in key for token in bad_tokens):
                    return text

        node = node.previous_element

    return ""


def _find_context_titles(table) -> Tuple[str, str]:
    """
    Ana kategori için H2/H1 kullanır.
    Tablo başlığında yalnızca DOM olarak bu tabloya ait yakın metni
    kullanır; önceki tablonun transfer başlığını taşımayı engeller.
    """
    kategori = "Genel"

    h2 = table.find_previous("h2")
    if h2:
        text = _normalize(h2.get_text(" ", strip=True))
        if _is_valid_title(text):
            kategori = text
    else:
        h1 = table.find_previous("h1")
        if h1:
            text = _normalize(h1.get_text(" ", strip=True))
            if _is_valid_title(text):
                kategori = text

    table_title = _nearest_text_between_tables(table)

    # Local text bulunamadıysa yalnızca aynı H2 bölümündeki
    # klasik heading'i fallback olarak kullan.
    if not table_title:
        for node in table.find_all_previous(
            ["h3", "h4", "h5", "h6", "strong", "button"],
            limit=20,
        ):
            # Arada başka bir table varsa bu heading eski tabloya aittir.
            between_table = node.find_next("table")
            if between_table is not table:
                continue

            text = _normalize(node.get_text(" ", strip=True))

            if not _is_valid_title(text):
                continue

            if _same_text(text, kategori):
                continue

            table_title = text
            break

    return kategori, table_title


# =========================================================
# TABLO GRID NORMALİZASYONU
# =========================================================

def _root_tables(soup):
    return [
        table
        for table in soup.find_all("table")
        if table.find_parent("table") is None
    ]


def _table_to_rows(table) -> List[List[str]]:
    """
    rowspan / colspan hücrelerini genişleterek dikdörtgene yakın
    bir satır matrisi üretir.

    Bu sayede bir satırda rowspan nedeniyle eksik hücre olduğunda
    kolonların sola kayması engellenir.
    """

    rows: List[List[str]] = []

    # column_index -> (text, remaining_rows)
    active_spans: Dict[int, Tuple[str, int]] = {}

    trs = [
        tr
        for tr in table.find_all("tr")
        if tr.find_parent("table") is table
    ]

    for tr in trs:
        cells = tr.find_all(["th", "td"], recursive=False)

        if not cells and not active_spans:
            continue

        row: List[str] = []
        col = 0
        cell_index = 0

        max_guard = max(
            len(cells) * 8 + len(active_spans) + 20,
            30,
        )

        while (
            cell_index < len(cells)
            or active_spans
        ) and col < max_guard:

            if col in active_spans:
                text, remaining = active_spans[col]
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
                # Aktif span varsa sıradaki span kolonuna kadar boşluk.
                future_cols = [
                    c
                    for c in active_spans.keys()
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
                cell.get_text(" ", strip=True)
            )

            try:
                rowspan = max(
                    int(cell.get("rowspan", 1)),
                    1,
                )
            except Exception:
                rowspan = 1

            try:
                colspan = max(
                    int(cell.get("colspan", 1)),
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
    headers = [_normalize_key(cell) for cell in row]

    tests = [
        ["masraf"],
        ["asgari", "tutar"],
        ["asgari", "oran"],
        ["azami", "tutar"],
        ["azami", "oran"],
        ["aciklama"],
        ["guncelleme", "tarihi"],
        ["guncellenme", "tarihi"],
    ]

    score = 0

    for keywords in tests:
        if any(
            all(k in header for k in keywords)
            for header in headers
        ):
            score += 1

    return score


def _find_header_index(rows: List[List[str]]) -> int:
    best_index = -1
    best_score = 0

    for index, row in enumerate(rows[:10]):
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

    for index, header in enumerate(headers):
        if all(keyword in header for keyword in keywords):
            return index

    return -1


def _find_columns(header_row: List[str]) -> Dict[str, int]:
    headers = [_normalize_key(x) for x in header_row]

    result = {
        "masraf": _find_col(headers, ["masraf"]),
        "asgari_tutar": _find_col(headers, ["asgari", "tutar"]),
        "asgari_oran": _find_col(headers, ["asgari", "oran"]),
        "azami_tutar": _find_col(headers, ["azami", "tutar"]),
        "azami_oran": _find_col(headers, ["azami", "oran"]),
        "aciklama": _find_col(headers, ["aciklama"]),
        "tarih": -1,
    }

    result["tarih"] = _find_col(
        headers,
        ["guncelleme", "tarihi"],
    )

    if result["tarih"] == -1:
        result["tarih"] = _find_col(
            headers,
            ["guncellenme", "tarihi"],
        )

    if result["tarih"] == -1:
        result["tarih"] = _find_col(
            headers,
            ["tarih"],
        )

    # Akbank'ın güncel ücret tabloları çoğunlukla 7 kolon.
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

    left = [_normalize_key(x) for x in row]
    right = [_normalize_key(x) for x in header]

    while left and not left[-1]:
        left.pop()

    while right and not right[-1]:
        right.pop()

    return left == right


# =========================================================
# MASRAF
# =========================================================

def _build_masraf(
    raw_masraf: str,
    table_title: str,
    aciklama: str,
    kategori: str = "",
) -> str:
    """
    Transfer başlığını yalnızca gerçekten gerekli olduğunda MASRAF'a ekler.

    Bu kısıtlama:
      - "Anlık EFT/Havale İşlem Ücreti - Aval Kredileri"
      - "Kobi Giden Swift Paketi 350 - KKB Risk Raporu"
      - "Fast Uluslararası - Posta İle Aylık Hesap Özeti"
    gibi yanlış başlık taşmalarını engeller.
    """
    raw_masraf = _normalize(raw_masraf)
    table_title = _normalize(table_title)
    aciklama = _normalize(aciklama)
    kategori = _normalize(kategori)

    masraf = raw_masraf

    raw_has_transfer = _contains_transfer_term(raw_masraf)
    title_has_transfer = _contains_transfer_term(table_title)

    # Satırın kendisinde transfer terimi yoksa tablo başlığını yalnızca
    # satır gerçekten jenerik/kanal-kart satırıysa öne ekle.
    #
    # Böylece yanlış bir context başlığı gelse bile:
    #   Kobi Giden Swift Paketi 350 - KKB Risk Raporu
    # gibi anlamsız birleşimler oluşmaz.
    raw_key = _normalize_key(raw_masraf)

    generic_transfer_rows = {
        "tum tutarlar icin",
        "axess business",
        "wings business",
        "axess kobi",
    }

    allow_title_prefix = (
        raw_key in generic_transfer_rows
        or raw_key.startswith("tum tutarlar")
    )

    if (
        table_title
        and title_has_transfer
        and not raw_has_transfer
        and allow_title_prefix
        and not _same_text(table_title, raw_masraf)
    ):
        masraf = _normalize(f"{table_title} - {raw_masraf}")

    # Akbank yurtiçi FAST tarifesi ayrı satır değil;
    # EFT ücret satırının açıklamasında aynı ücretin FAST için de
    # uygulanacağı açıkça belirtiliyor.
    aciklama_key = _normalize_key(aciklama)

    fast_same_fee = (
        "fast sistemi uzerinden yapilan islemler icin de ayni ucret uygulanir"
        in aciklama_key
        or (
            _has_transfer_term(aciklama, "fast")
            and "ayni ucret" in aciklama_key
        )
    )

    if fast_same_fee and not _has_transfer_term(masraf, "fast"):
        masraf = _normalize(f"FAST / {masraf}")

    # Standart döviz havale/transfer kalemleri Akbank'ın SWIFT hizmetinin
    # ücret karşılığıdır. Fast Uluslararası ve Western Union'ı SWIFT diye
    # etiketlemiyoruz.
    combined = _normalize_key(
        " ".join([kategori, table_title, raw_masraf])
    )

    is_standard_fx_transfer = (
        (
            "uluslararasi fon transferi" in combined
            or "doviz havale" in combined
            or "doviz transfer" in combined
        )
        and "fast uluslararasi" not in combined
        and "western union" not in combined
    )

    if is_standard_fx_transfer and not _has_transfer_term(masraf, "swift"):
        masraf = _normalize(f"SWIFT - {masraf}")

    return masraf


# =========================================================
# PARSER
# =========================================================

def _parse_soup(
    soup: BeautifulSoup,
    source_name: str,
) -> Tuple[List[UcretSatiri], Dict[str, int]]:

    tables = _root_tables(soup)

    stats: Dict[str, int] = {
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

    results: List[UcretSatiri] = []

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

    kategori_sayilari: Dict[str, int] = {}

    for table_index, table in enumerate(tables):
        rows = _table_to_rows(table)

        if not rows:
            stats["ignored_tables"] += 1
            continue

        header_index = _find_header_index(rows)

        if header_index == -1:
            stats["ignored_tables"] += 1
            continue

        stats["fee_tables"] += 1

        header = [
            _normalize(x)
            for x in rows[header_index]
        ]

        columns = _find_columns(header)

        kategori, table_title = _find_context_titles(
            table
        )

        table_record_count = 0

        for raw_row in rows[header_index + 1:]:
            row = [
                _normalize(x)
                for x in raw_row
            ]

            if not row or not any(row):
                continue

            stats["candidate_rows"] += 1

            if _row_is_same_header(row, header):
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
                if index < 0 or index >= len(row):
                    return ""
                return _normalize(row[index])

            raw_masraf = get(columns["masraf"])

            if not raw_masraf:
                stats["invalid_rows"] += 1
                continue

            asgari_tutar = get(
                columns["asgari_tutar"]
            )
            asgari_oran = get(
                columns["asgari_oran"]
            )
            azami_tutar = get(
                columns["azami_tutar"]
            )
            azami_oran = get(
                columns["azami_oran"]
            )

            aciklama_raw = get(
                columns["aciklama"]
            )

            aciklama, aciklama_tarihi = (
                _parse_aciklama(aciklama_raw)
            )

            site_tarihi = get(
                columns["tarih"]
            ).replace("/", ".")

            if not site_tarihi:
                site_tarihi = aciklama_tarihi

            # Masraf dışında bütün alanlar boşsa
            # gerçek ücret kaydı olduğuna dair veri yok.
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
                table_title=table_title,
                aciklama=aciklama_raw,
                kategori=kategori,
            )

            stats["parsed_before_dedup"] += 1

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

            kategori_sayilari[kategori] = (
                kategori_sayilari.get(
                    kategori,
                    0,
                )
                + 1
            )

        if table_record_count == 0:
            stats["zero_record_tables"] += 1

            print(
                f"[akbank][DEBUG][{source_name}] "
                f"Ücret tablosu {table_index} 0 kayıt üretti | "
                f"Kategori: {kategori} | "
                f"Tablo başlığı: {table_title or '-'} | "
                f"Satır: {len(rows)} | "
                f"Header: {header}",
                file=sys.stderr,
            )

    print(
        f"[akbank][{source_name}] Toplam root tablo: "
        f"{stats['tables_total']}",
        file=sys.stderr,
    )
    print(
        f"[akbank][{source_name}] Ücret tablosu: "
        f"{stats['fee_tables']}",
        file=sys.stderr,
    )
    print(
        f"[akbank][{source_name}] İlgisiz/atlanan tablo: "
        f"{stats['ignored_tables']}",
        file=sys.stderr,
    )
    print(
        f"[akbank][{source_name}] 0 kayıt üreten ücret tablosu: "
        f"{stats['zero_record_tables']}",
        file=sys.stderr,
    )
    print(
        f"[akbank][{source_name}] Benzersiz sonuç: "
        f"{len(results)}",
        file=sys.stderr,
    )

    print("", file=sys.stderr)
    print(
        f"[akbank][{source_name}] ===== KATEGORİ RAPORU =====",
        file=sys.stderr,
    )

    for kategori, count in sorted(
        kategori_sayilari.items(),
        key=lambda item: (-item[1], item[0]),
    ):
        print(
            f"[akbank][{source_name}] "
            f"{kategori} -> {count} kayıt",
            file=sys.stderr,
        )

    print(
        f"[akbank][{source_name}] ===========================",
        file=sys.stderr,
    )

    return results, stats


# =========================================================
# REQUESTS
# =========================================================

def _scrape_with_requests(
    url: str = AKBANK_URL,
) -> Tuple[List[UcretSatiri], Dict[str, int]]:

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
    url: str = AKBANK_URL,
) -> Tuple[List[UcretSatiri], Dict[str, int]]:

    try:
        from playwright.sync_api import sync_playwright
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
                user_agent=HEADERS["User-Agent"],
                viewport={
                    "width": 1440,
                    "height": 1000,
                },
                locale="tr-TR",
                timezone_id="Europe/Istanbul",
                extra_http_headers={
                    "Accept-Language": (
                        "tr-TR,tr;q=0.9,en;q=0.8"
                    ),
                },
            )

            page = context.new_page()

            page.goto(
                url,
                timeout=120000,
                wait_until="domcontentloaded",
            )

            page.wait_for_timeout(2500)

            # Cookie
            for text in [
                "Tümünü Kabul Et",
                "Tümünü Kabul",
                "Kabul Et",
                "Kabul",
                "Kapat",
            ]:
                try:
                    locator = page.get_by_text(
                        text,
                        exact=True,
                    ).first

                    if locator.is_visible(
                        timeout=700
                    ):
                        locator.click(
                            timeout=2000
                        )
                        page.wait_for_timeout(300)
                        break
                except Exception:
                    pass

            # Sadece gerçekten kapalı aria-expanded kontrolleri.
            # Genel class* selector'larına tıklamıyoruz;
            # yanlış element/nav linki açıp sayfayı bozabilir.
            for round_no in range(8):
                opened = 0

                elements = page.locator(
                    "[aria-expanded='false']"
                )

                count = elements.count()

                for i in range(count):
                    try:
                        element = elements.nth(i)

                        if not element.is_visible(
                            timeout=150
                        ):
                            continue

                        element.scroll_into_view_if_needed(
                            timeout=1000
                        )

                        element.click(
                            timeout=1500,
                            force=True,
                        )

                        opened += 1
                        page.wait_for_timeout(60)

                    except Exception:
                        pass

                if opened == 0:
                    break

                print(
                    f"[akbank] Accordion turu "
                    f"{round_no + 1}: "
                    f"{opened} adet açıldı.",
                    file=sys.stderr,
                )

                page.wait_for_timeout(250)

            # Güvenli scroll.
            stable = 0
            previous_height = 0

            for _ in range(100):
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

                    page.wait_for_timeout(100)

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

            page.wait_for_timeout(800)
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
# RAPOR
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

    print("", file=sys.stderr)
    print(
        "[akbank] ===== BÜTÜNLÜK KONTROLÜ =====",
        file=sys.stderr,
    )

    print(
        f"[akbank] Toplam root tablo: "
        f"{stats.get('tables_total', 0)}",
        file=sys.stderr,
    )
    print(
        f"[akbank] Ücret tablosu: "
        f"{stats.get('fee_tables', 0)}",
        file=sys.stderr,
    )
    print(
        f"[akbank] Ham ücret satırı adayı: "
        f"{candidate_rows}",
        file=sys.stderr,
    )
    print(
        f"[akbank] Parse edilen (dedup öncesi): "
        f"{stats.get('parsed_before_dedup', 0)}",
        file=sys.stderr,
    )
    print(
        f"[akbank] Duplicate: "
        f"{stats.get('duplicates', 0)}",
        file=sys.stderr,
    )
    print(
        f"[akbank] Tekrarlanan header: "
        f"{stats.get('repeated_headers', 0)}",
        file=sys.stderr,
    )
    print(
        f"[akbank] Not / tek hücreli satır: "
        f"{stats.get('notes', 0)}",
        file=sys.stderr,
    )
    print(
        f"[akbank] Geçersiz / boş veri satırı: "
        f"{stats.get('invalid_rows', 0)}",
        file=sys.stderr,
    )
    print(
        f"[akbank] 0 kayıt üreten ücret tablosu: "
        f"{stats.get('zero_record_tables', 0)}",
        file=sys.stderr,
    )
    print(
        f"[akbank] Excel'e gidecek benzersiz satır: "
        f"{result_count}",
        file=sys.stderr,
    )

    if candidate_rows == explained:
        print(
            "[akbank] BÜTÜNLÜK: OK - "
            "aday satırların tamamı açıklandı.",
            file=sys.stderr,
        )
    else:
        print(
            "[akbank] BÜTÜNLÜK: UYARI - "
            f"{candidate_rows - explained} "
            "aday satır açıklanamadı.",
            file=sys.stderr,
        )

    print(
        "[akbank] ===============================",
        file=sys.stderr,
    )


def _print_transfer_report(
    rows: List[UcretSatiri],
) -> None:

    print("", file=sys.stderr)
    print(
        "[akbank] ===== PARA AKTARMA KONTROLÜ =====",
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
            f"[akbank] {label}: "
            f"{len(found)} kayıt",
            file=sys.stderr,
        )

        for row in found[:10]:
            print(
                f"    - {row.masraf}",
                file=sys.stderr,
            )

        if not found:
            print(
                f"[akbank][UYARI] "
                f"{label} MASRAF alanında hiç bulunamadı.",
                file=sys.stderr,
            )

    fast_from_description = [
        row
        for row in rows
        if (
            _has_transfer_term(row.masraf, "fast")
            and "ayni ucret" in _normalize_key(row.aciklama)
        )
    ]

    print(
        f"[akbank] FAST (EFT tarifesiyle aynı ücret): "
        f"{len(fast_from_description)} kayıt",
        file=sys.stderr,
    )

    print(
        "[akbank] =================================",
        file=sys.stderr,
    )


# =========================================================
# ANA SCRAPER
# =========================================================

def scrape_akbank(
    url: str = AKBANK_URL,
) -> List[UcretSatiri]:

    print(
        f"[akbank] SÜRÜM: {SCRAPER_VERSION}",
        file=sys.stderr,
    )

    print(
        f"[akbank] {url} adresinden veri çekiliyor...",
        file=sys.stderr,
    )

    rows: List[UcretSatiri] = []
    stats: Optional[Dict[str, int]] = None
    source = ""

    # -----------------------------------------------------
    # 1) REQUESTS
    # -----------------------------------------------------

    try:
        request_rows, request_stats = (
            _scrape_with_requests(url)
        )

        print(
            f"[akbank] requests sonucu: "
            f"{len(request_rows)} benzersiz satır.",
            file=sys.stderr,
        )

        # Akbank sayfasının server-rendered sürümünün
        # anlamlı büyüklükte veri içerdiğini doğrula.
        if (
            request_rows
            and request_stats.get(
                "fee_tables",
                0,
            ) >= 10
            and request_stats.get(
                "candidate_rows",
                0,
            ) >= 50
        ):
            rows = request_rows
            stats = request_stats
            source = "requests"

        else:
            print(
                "[akbank][UYARI] requests sonucu "
                "beklenenden küçük; "
                "Playwright deneniyor.",
                file=sys.stderr,
            )

    except Exception as exc:
        print(
            f"[akbank] requests başarısız: {exc}",
            file=sys.stderr,
        )

    # -----------------------------------------------------
    # 2) PLAYWRIGHT FALLBACK
    # -----------------------------------------------------

    if not rows:
        try:
            playwright_rows, playwright_stats = (
                _scrape_with_playwright(url)
            )

            print(
                f"[akbank] Playwright sonucu: "
                f"{len(playwright_rows)} "
                "benzersiz satır.",
                file=sys.stderr,
            )

            rows = playwright_rows
            stats = playwright_stats
            source = "playwright"

        except Exception as exc:
            print(
                f"[akbank] Playwright başarısız: {exc}",
                file=sys.stderr,
            )

    if not rows or stats is None:
        raise ScraperError(
            "Akbank sayfasından hiçbir ücret "
            "satırı çekilemedi."
        )

    rows = sorted(
        rows,
        key=lambda row: (
            _normalize_key(row.kategori),
            _normalize_key(row.masraf),
        ),
    )

    print(
        f"[akbank] Kullanılan kaynak: {source}",
        file=sys.stderr,
    )

    _print_integrity_report(
        stats,
        len(rows),
    )

    _print_transfer_report(rows)

    print(
        f"[akbank] Toplam "
        f"{len(rows)} benzersiz satır bulundu.",
        file=sys.stderr,
    )

    return rows


# =========================================================
# TEST
# =========================================================

if __name__ == "__main__":
    try:
        veriler = scrape_akbank()

        print()
        print("=" * 70)
        print("AKBANK SCRAPER")
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
                f"{i}. [{row.kategori}] "
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
            f"[akbank][HATA] {exc}",
            file=sys.stderr,
        )
        sys.exit(1)
