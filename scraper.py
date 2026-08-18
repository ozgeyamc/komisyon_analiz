"""
Garanti BBVA "Ürün ve Hizmet Ücretleri" scraper.

Amaç:
- Sayfadaki tüm ücret tablolarını mümkün olduğunca eksiksiz toplamak
- H2 başlığını ana kategori, tablo/accordion başlığını alt başlık olarak kullanmak
- EFT / FAST / Havale / SWIFT gibi ifadelerin Excel MASRAF filtresinde kaybolmasını önlemek
- Gerçek veri satırlarını yanlışlıkla header diye atlamamak
- Duplicate kayıtları kontrollü temizlemek
- GitHub Actions logunda bütünlük ve para transferi kontrolü üretmek
- main.py / update_excel.py ile uyumlu UcretSatiri listesi döndürmek
"""

import re
import sys
from dataclasses import dataclass
from typing import Dict, List, Optional, Set, Tuple

import requests
from bs4 import BeautifulSoup

SCRAPER_VERSION = "2026-08-19-v2-garanti-integrity"
GARANTI_URL = "https://www.garantibbva.com.tr/urun-ve-hizmet-ucretleri"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "tr-TR,tr;q=0.9,en;q=0.8",
}

DATE_PATTERN = re.compile(
    r"G[üu]ncellenme\s*Tarihi\s*:\s*[\s\xa0]*"
    r"(\d{1,2}[./]\d{1,2}[./]\d{4}(?:[\s\xa0]+\d{1,2}:\d{2})?)",
    re.IGNORECASE,
)
DATE_PATTERN_SON = re.compile(
    r"Son\s+g[üu]ncellenme\s+tarihi\s*:\s*[\s\xa0]*"
    r"(\d{1,2})\s+"
    r"(ocak|şubat|mart|nisan|mayıs|haziran|temmuz|ağustos|eylül|ekim|kasım|aralık)\s+"
    r"(\d{4})",
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
DATE_PATTERN_PAREN_TR = re.compile(
    r"\(?\s*(\d{1,2})\s+"
    r"(ocak|şubat|mart|nisan|mayıs|haziran|temmuz|ağustos|eylül|ekim|kasım|aralık)\s+"
    r"(\d{4})\s*\)?",
    re.IGNORECASE,
)

TURKCE_AYLAR = {
    "ocak": "01", "şubat": "02", "mart": "03", "nisan": "04",
    "mayıs": "05", "haziran": "06", "temmuz": "07", "ağustos": "08",
    "eylül": "09", "ekim": "10", "kasım": "11", "aralık": "12",
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


def _normalize(value: Optional[str]) -> str:
    if value is None:
        return ""
    value = str(value)
    value = value.replace("\xa0", " ").replace("\u200b", "")
    value = value.replace("\r", " ").replace("\n", " ")
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def _normalize_header(value: Optional[str]) -> str:
    value = _normalize(value).lower()
    replacements = {"ı": "i", "ğ": "g", "ü": "u", "ş": "s", "ö": "o", "ç": "c"}
    for old, new in replacements.items():
        value = value.replace(old, new)
    value = value.replace("%", " ")
    value = re.sub(r"\([^)]*\)", " ", value)
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def _same_text(a: str, b: str) -> bool:
    return _normalize_header(a) == _normalize_header(b)


def _parse_aciklama(raw_aciklama: str) -> Tuple[str, str]:
    raw = _normalize(raw_aciklama)
    if not raw:
        return "", ""

    match = DATE_PATTERN.search(raw)
    if match:
        tarih = match.group(1).replace("/", ".").strip()
        temiz = _normalize(DATE_PATTERN.sub("", raw)).strip(" .:-")
        return temiz, tarih

    match = DATE_PATTERN_SON.search(raw)
    if match:
        gun = match.group(1).zfill(2)
        ay = TURKCE_AYLAR.get(match.group(2).lower(), "")
        yil = match.group(3)
        if ay:
            tarih = f"{gun}.{ay}.{yil}"
            temiz = _normalize(DATE_PATTERN_SON.sub("", raw)).strip(" .:-")
            return temiz, tarih

    match = DATE_PATTERN_ITIBAR.search(raw)
    if match:
        tarih = match.group(1).replace("/", ".").strip()
        return raw, tarih

    match = DATE_PATTERN_TR.search(raw)
    if match:
        gun = match.group(1).zfill(2)
        ay = TURKCE_AYLAR.get(match.group(2).lower(), "")
        yil = match.group(3)
        if ay:
            tarih = f"{gun}.{ay}.{yil}"
            temiz = _normalize(DATE_PATTERN_TR.sub("", raw)).strip(" .:-")
            return temiz, tarih

    match = DATE_PATTERN_PAREN_TR.search(raw)
    if match:
        gun = match.group(1).zfill(2)
        ay = TURKCE_AYLAR.get(match.group(2).lower(), "")
        yil = match.group(3)
        if ay:
            return raw, f"{gun}.{ay}.{yil}"

    return raw, ""


INVALID_TITLES = {
    "müşteri ol", "ara", "kapat", "giriş yap", "daha fazla gör",
    "masraf", "asgari tutar", "asgari oran", "azami tutar", "azami oran", "açıklama",
}


def _is_valid_title(text: str) -> bool:
    text = _normalize(text)
    if not text or len(text) < 3 or len(text) > 180:
        return False
    normalized = _normalize_header(text)
    invalid_normalized = {_normalize_header(x) for x in INVALID_TITLES}
    if normalized in invalid_normalized:
        return False
    if normalized in {"bireysel", "ticari", "kurumsal", "anasayfa", "iletisim"}:
        return False
    return True


def _find_context_titles(table) -> Tuple[str, str]:
    kategori = "Genel"
    table_title = ""

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

    for node in table.find_all_previous(
        ["h2", "h3", "h4", "h5", "h6", "button", "a", "strong"],
        limit=30,
    ):
        if node.name == "h2":
            break
        text = _normalize(node.get_text(" ", strip=True))
        if not _is_valid_title(text):
            continue
        if _same_text(text, kategori):
            continue
        table_title = text
        break

    return kategori, table_title


def _header_score(values: List[str]) -> int:
    headers = [_normalize_header(x) for x in values]
    tests = [
        ["masraf"], ["asgari", "tutar"], ["asgari", "oran"],
        ["azami", "tutar"], ["azami", "oran"], ["aciklama"],
        ["guncelleme"], ["guncellenme"],
    ]
    score = 0
    for keywords in tests:
        if any(all(k in h for k in keywords) for h in headers):
            score += 1
    return score


def _find_header_index(rows: List[List[str]]) -> int:
    best_index = -1
    best_score = 0
    for index, row in enumerate(rows[:8]):
        score = _header_score(row)
        if score > best_score:
            best_score = score
            best_index = index
    return best_index if best_score >= 3 else -1


def _find_col(headers: List[str], keywords: List[str]) -> int:
    for index, header in enumerate(headers):
        if all(keyword in header for keyword in keywords):
            return index
    return -1


def _find_columns(header_row: List[str]) -> Dict[str, int]:
    headers = [_normalize_header(x) for x in header_row]
    result = {
        "masraf": _find_col(headers, ["masraf"]),
        "asgari_tutar": _find_col(headers, ["asgari", "tutar"]),
        "asgari_oran": _find_col(headers, ["asgari", "oran"]),
        "azami_tutar": _find_col(headers, ["azami", "tutar"]),
        "azami_oran": _find_col(headers, ["azami", "oran"]),
        "aciklama": _find_col(headers, ["aciklama"]),
        "tarih": -1,
    }
    result["tarih"] = _find_col(headers, ["guncelleme"])
    if result["tarih"] == -1:
        result["tarih"] = _find_col(headers, ["guncellenme"])
    if result["tarih"] == -1:
        result["tarih"] = _find_col(headers, ["tarih"])

    if len(header_row) >= 6:
        fallbacks = {
            "masraf": 0, "asgari_tutar": 1, "asgari_oran": 2,
            "azami_tutar": 3, "azami_oran": 4, "aciklama": 5,
        }
        for key, index in fallbacks.items():
            if result[key] == -1:
                result[key] = index
    return result


def _row_is_same_header(row: List[str], header: List[str]) -> bool:
    left = [_normalize_header(x) for x in row]
    right = [_normalize_header(x) for x in header]
    while left and not left[-1]:
        left.pop()
    while right and not right[-1]:
        right.pop()
    return left == right


def _build_masraf_name(raw_masraf: str, table_title: str, aciklama: str) -> str:
    raw_masraf = _normalize(raw_masraf)
    table_title = _normalize(table_title)
    parts: List[str] = []

    if table_title:
        title_core = re.sub(
            r"\b(ucreti|ucretleri|islemleri)\b",
            "",
            _normalize_header(table_title),
        ).strip()
        raw_norm = _normalize_header(raw_masraf)
        if title_core and title_core not in raw_norm:
            parts.append(table_title)

    parts.append(raw_masraf)
    masraf = " - ".join(p for p in parts if p)

    if "swift" in _normalize_header(aciklama) and "swift" not in _normalize_header(masraf):
        masraf = f"SWIFT - {masraf}"

    return _normalize(masraf)


def _table_to_rows(table) -> List[List[str]]:
    rows: List[List[str]] = []
    for tr in table.find_all("tr"):
        if tr.find_parent("table") is not table:
            continue
        cells = tr.find_all(["th", "td"], recursive=False)
        if not cells:
            continue
        values = [_normalize(cell.get_text(" ", strip=True)) for cell in cells]
        if any(values):
            rows.append(values)
    return rows


def _parse_tables(soup: BeautifulSoup, source_name: str) -> Tuple[List[UcretSatiri], Dict[str, int]]:
    all_tables = soup.find_all("table")
    tables = [table for table in all_tables if table.find_parent("table") is None]

    stats: Dict[str, int] = {
        "tables_total": len(tables), "fee_tables": 0, "ignored_tables": 0,
        "zero_record_tables": 0, "candidate_rows": 0,
        "parsed_before_dedup": 0, "duplicates": 0, "repeated_headers": 0,
        "notes": 0, "invalid_rows": 0,
    }

    results: List[UcretSatiri] = []
    seen: Set[Tuple[str, str, str, str, str, str, str, str]] = set()
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
        header = rows[header_index]
        column_map = _find_columns(header)
        kategori, table_title = _find_context_titles(table)
        table_record_count = 0

        for row in rows[header_index + 1:]:
            if not row or not any(row):
                continue

            stats["candidate_rows"] += 1

            if _row_is_same_header(row, header):
                stats["repeated_headers"] += 1
                continue

            meaningful_cells = sum(1 for value in row if _normalize(value))
            if meaningful_cells < 2:
                stats["notes"] += 1
                continue

            def get(index: int) -> str:
                if index < 0 or index >= len(row):
                    return ""
                return _normalize(row[index])

            raw_masraf = get(column_map["masraf"])
            if not raw_masraf:
                stats["invalid_rows"] += 1
                continue

            asgari_tutar = get(column_map["asgari_tutar"])
            asgari_oran = get(column_map["asgari_oran"])
            azami_tutar = get(column_map["azami_tutar"])
            azami_oran = get(column_map["azami_oran"])
            aciklama_raw = get(column_map["aciklama"])
            aciklama, aciklama_tarihi = _parse_aciklama(aciklama_raw)
            site_tarihi = get(column_map["tarih"]).replace("/", ".")
            if not site_tarihi:
                site_tarihi = aciklama_tarihi

            if not any([asgari_tutar, asgari_oran, azami_tutar, azami_oran, aciklama, site_tarihi]):
                stats["invalid_rows"] += 1
                continue

            masraf = _build_masraf_name(raw_masraf, table_title, aciklama_raw)
            stats["parsed_before_dedup"] += 1

            key = (
                kategori, masraf, asgari_tutar, asgari_oran,
                azami_tutar, azami_oran, aciklama, site_tarihi,
            )
            if key in seen:
                stats["duplicates"] += 1
                continue

            seen.add(key)
            results.append(UcretSatiri(
                kategori=kategori,
                masraf=masraf,
                asgari_tutar=asgari_tutar,
                asgari_oran=asgari_oran,
                azami_tutar=azami_tutar,
                azami_oran=azami_oran,
                aciklama=aciklama,
                site_guncelleme_tarihi=site_tarihi,
            ))
            table_record_count += 1
            kategori_sayilari[kategori] = kategori_sayilari.get(kategori, 0) + 1

        if table_record_count == 0:
            stats["zero_record_tables"] += 1
            print(
                f"[garanti][DEBUG][{source_name}] Ücret tablosu {table_index} 0 kayıt üretti | "
                f"Kategori: {kategori} | Tablo başlığı: {table_title or '-'} | "
                f"Satır: {len(rows)} | Header: {header}",
                file=sys.stderr,
            )

    print(f"[garanti][{source_name}] Toplam root tablo: {stats['tables_total']}", file=sys.stderr)
    print(f"[garanti][{source_name}] Ücret tablosu: {stats['fee_tables']}", file=sys.stderr)
    print(f"[garanti][{source_name}] İlgisiz/atlanan tablo: {stats['ignored_tables']}", file=sys.stderr)
    print(f"[garanti][{source_name}] 0 kayıt üreten ücret tablosu: {stats['zero_record_tables']}", file=sys.stderr)
    print(f"[garanti][{source_name}] Benzersiz sonuç: {len(results)}", file=sys.stderr)

    print("", file=sys.stderr)
    print(f"[garanti][{source_name}] ===== KATEGORİ RAPORU =====", file=sys.stderr)
    for kategori, count in sorted(kategori_sayilari.items(), key=lambda item: (-item[1], item[0])):
        print(f"[garanti][{source_name}] {kategori} -> {count} kayıt", file=sys.stderr)
    print(f"[garanti][{source_name}] ===========================", file=sys.stderr)

    return results, stats


def _scrape_with_playwright(url: str = GARANTI_URL) -> Tuple[List[UcretSatiri], Dict[str, int]]:
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise ScraperError("Playwright kurulu değil.") from exc

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        try:
            context = browser.new_context(
                user_agent=HEADERS["User-Agent"],
                locale="tr-TR",
                viewport={"width": 1440, "height": 1080},
                extra_http_headers={"Accept-Language": "tr-TR,tr;q=0.9,en;q=0.8"},
            )
            page = context.new_page()
            page.goto(url, timeout=120000, wait_until="domcontentloaded")
            page.wait_for_timeout(2500)

            for text in ["Tümünü Kabul Et", "Tümünü Kabul", "Kabul Et", "Kabul", "Kapat"]:
                try:
                    locator = page.get_by_text(text, exact=True).first
                    if locator.is_visible(timeout=800):
                        locator.click(timeout=2000)
                        page.wait_for_timeout(400)
                        break
                except Exception:
                    pass

            print("[garanti] Accordion'lar açılıyor...", file=sys.stderr)
            for round_no in range(8):
                opened = 0
                try:
                    elements = page.locator("[aria-expanded='false']")
                    count = elements.count()
                    for i in range(count):
                        try:
                            element = elements.nth(i)
                            if not element.is_visible(timeout=150):
                                continue
                            element.scroll_into_view_if_needed(timeout=1000)
                            element.click(timeout=1500, force=True)
                            opened += 1
                            page.wait_for_timeout(70)
                        except Exception:
                            pass
                except Exception:
                    pass
                if opened == 0:
                    break
                print(f"[garanti] Accordion turu {round_no + 1}: {opened} adet açıldı.", file=sys.stderr)
                page.wait_for_timeout(250)

            previous_height = 0
            stable = 0
            for _ in range(120):
                page.evaluate("window.scrollBy(0, Math.max(window.innerHeight * 0.8, 600));")
                page.wait_for_timeout(120)
                height = page.evaluate("document.body.scrollHeight")
                bottom = page.evaluate(
                    "window.innerHeight + window.scrollY >= document.body.scrollHeight - 100"
                )
                if bottom and height == previous_height:
                    stable += 1
                else:
                    stable = 0
                previous_height = height
                if stable >= 4:
                    break

            try:
                page.wait_for_load_state("networkidle", timeout=5000)
            except Exception:
                pass

            html = page.content()
        finally:
            browser.close()

    return _parse_tables(BeautifulSoup(html, "lxml"), source_name="playwright")


def _scrape_with_requests(url: str = GARANTI_URL) -> Tuple[List[UcretSatiri], Dict[str, int]]:
    try:
        response = requests.get(url, headers=HEADERS, timeout=40)
        response.raise_for_status()
    except requests.RequestException as exc:
        raise ScraperError(f"Sayfaya requests ile erişilemedi: {exc}") from exc

    return _parse_tables(BeautifulSoup(response.text, "lxml"), source_name="requests")


def _merge_results(*groups: List[UcretSatiri]) -> List[UcretSatiri]:
    merged: List[UcretSatiri] = []
    seen: Set[Tuple[str, str, str, str, str, str, str, str]] = set()

    for group in groups:
        for row in group:
            key = (
                row.kategori, row.masraf, row.asgari_tutar, row.asgari_oran,
                row.azami_tutar, row.azami_oran, row.aciklama, row.site_guncelleme_tarihi,
            )
            if key in seen:
                continue
            seen.add(key)
            merged.append(row)

    return merged


def _print_transfer_report(rows: List[UcretSatiri]) -> None:
    print("", file=sys.stderr)
    print("[garanti] ===== PARA AKTARMA KONTROLÜ =====", file=sys.stderr)

    for label, needle in [("FAST", "fast"), ("EFT", "eft"), ("Havale", "havale"), ("SWIFT", "swift")]:
        found = [
            row for row in rows
            if needle in _normalize_header(row.masraf) or needle in _normalize_header(row.aciklama)
        ]
        print(f"[garanti] {label}: {len(found)} kayıt", file=sys.stderr)
        for row in found[:8]:
            print(f"    - {row.masraf}", file=sys.stderr)

    print("[garanti] =================================", file=sys.stderr)


def _print_integrity_report(stats: Dict[str, int], result_count: int) -> None:
    candidate_rows = stats.get("candidate_rows", 0)
    parsed_before_dedup = stats.get("parsed_before_dedup", 0)
    repeated_headers = stats.get("repeated_headers", 0)
    notes = stats.get("notes", 0)
    invalid_rows = stats.get("invalid_rows", 0)
    explained = parsed_before_dedup + repeated_headers + notes + invalid_rows

    print("", file=sys.stderr)
    print("[garanti] ===== BÜTÜNLÜK KONTROLÜ =====", file=sys.stderr)
    print(f"[garanti] Ham ücret satırı adayı: {candidate_rows}", file=sys.stderr)
    print(f"[garanti] Parse edilen (dedup öncesi): {parsed_before_dedup}", file=sys.stderr)
    print(f"[garanti] Duplicate: {stats.get('duplicates', 0)}", file=sys.stderr)
    print(f"[garanti] Tekrarlanan header: {repeated_headers}", file=sys.stderr)
    print(f"[garanti] Not / tek hücreli satır: {notes}", file=sys.stderr)
    print(f"[garanti] Geçersiz / boş veri satırı: {invalid_rows}", file=sys.stderr)
    print(f"[garanti] Excel'e gidecek benzersiz satır: {result_count}", file=sys.stderr)

    if candidate_rows == explained:
        print("[garanti] BÜTÜNLÜK: OK - DOM'daki aday satırların tamamı açıklandı.", file=sys.stderr)
    else:
        print(
            f"[garanti] BÜTÜNLÜK: UYARI - {candidate_rows - explained} aday satır açıklanamadı.",
            file=sys.stderr,
        )

    print("[garanti] ===============================", file=sys.stderr)


def scrape_garanti_bbva(url: str = GARANTI_URL) -> List[UcretSatiri]:
    print(f"[garanti] SÜRÜM: {SCRAPER_VERSION}", file=sys.stderr)
    print(f"[garanti] {url} adresinden veri çekiliyor...", file=sys.stderr)

    playwright_rows: List[UcretSatiri] = []
    playwright_stats: Optional[Dict[str, int]] = None
    requests_rows: List[UcretSatiri] = []
    requests_stats: Optional[Dict[str, int]] = None

    try:
        playwright_rows, playwright_stats = _scrape_with_playwright(url)
        print(f"[garanti] Playwright sonucu: {len(playwright_rows)} benzersiz satır.", file=sys.stderr)
    except Exception as exc:
        print(f"[garanti] Playwright başarısız: {exc}", file=sys.stderr)

    try:
        requests_rows, requests_stats = _scrape_with_requests(url)
        print(f"[garanti] requests sonucu: {len(requests_rows)} benzersiz satır.", file=sys.stderr)
    except Exception as exc:
        print(f"[garanti] requests başarısız: {exc}", file=sys.stderr)

    if not playwright_rows and not requests_rows:
        raise ScraperError("Garanti sayfasından hiçbir ücret satırı çekilemedi.")

    rows = _merge_results(playwright_rows, requests_rows)
    print(f"[garanti] Birleştirilmiş toplam benzersiz ücret: {len(rows)}", file=sys.stderr)

    stats_candidates = [s for s in [playwright_stats, requests_stats] if s is not None]
    if stats_candidates:
        reference_stats = max(
            stats_candidates,
            key=lambda x: (x.get("candidate_rows", 0), x.get("fee_tables", 0)),
        )
        _print_integrity_report(reference_stats, len(rows))

    _print_transfer_report(rows)
    return rows


if __name__ == "__main__":
    try:
        veriler = scrape_garanti_bbva()
        print()
        print("=" * 70)
        print("GARANTİ BBVA SCRAPER")
        print("=" * 70)
        print(f"Toplam çekilen ücret: {len(veriler)}")
        print()

        for i, row in enumerate(veriler[:40], start=1):
            print(f"{i}. [{row.kategori}] {row.masraf}")
            print(f"   Asgari Tutar : {row.asgari_tutar}")
            print(f"   Asgari Oran  : {row.asgari_oran}")
            print(f"   Azami Tutar  : {row.azami_tutar}")
            print(f"   Azami Oran   : {row.azami_oran}")
            print(f"   Açıklama     : {row.aciklama}")
            print(f"   Tarih        : {row.site_guncelleme_tarihi}")
            print("-" * 70)

    except Exception as exc:
        print(f"[garanti][HATA] {exc}", file=sys.stderr)
        sys.exit(1)
