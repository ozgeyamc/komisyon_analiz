"""
İş Bankası "Ürün ve Hizmet Ücretleri" scraper.

Bu sürüm:
- h1, h2... grup numaralarında boşluk olsa bile sonraki grupları kaçırmaz.
- Sayfadaki tüm UHU_item_icerikC bloklarını sayar ve yapı dışında kalan blokları da fallback ile işler.
- EFT / FAST / Havale / SWIFT başlığını gerektiğinde MASRAF alanına taşır.
- Duplicate kayıtları temizler.
- GitHub Actions loguna kategori, bütünlük ve para aktarma kontrolü yazar.
- main.py / update_excel.py ile uyumlu UcretSatiri listesi döndürür.
"""

import re
import sys
from dataclasses import dataclass
from typing import Dict, List, Optional, Set, Tuple

ISBANK_URL = "https://www.isbank.com.tr/urun-ve-hizmet-ucretleri"

SCRAPER_VERSION = "2026-08-19-v3-isbank-transfer-boundaries"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "tr-TR,tr;q=0.9,en;q=0.8",
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


# =========================================================
# METİN YARDIMCILARI
# =========================================================

def _normalize(val) -> str:
    if val is None:
        return ""

    text = str(val)
    text = text.replace("\xa0", " ")
    text = text.replace("\u200b", "")
    text = text.replace("\r", " ")
    text = text.replace("\n", " ")
    text = re.sub(r"\s+", " ", text)

    return text.strip()


def _normalize_key(val) -> str:
    text = _normalize(val).lower()

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


def _meta(el, cls="UHU_icerik_meta") -> str:
    """
    İş Bankası sayfasında değerler çoğunlukla span içinde geliyor.
    Class yapısı değişirse kolonun kendi text'ine fallback yapar.
    """
    if not el:
        return ""

    span = el.find("span", class_=cls)
    if span:
        return _normalize(span.get_text(" ", strip=True))

    # Bazı bloklarda beklenen class olmayabilir.
    # Önce herhangi bir span'ın metnini dene.
    spans = el.find_all("span")
    for candidate in reversed(spans):
        text = _normalize(candidate.get_text(" ", strip=True))
        if text:
            return text

    return _normalize(el.get_text(" ", strip=True))


def _same_text(a: str, b: str) -> bool:
    return _normalize_key(a) == _normalize_key(b)


def _join_category(parts: List[str]) -> str:
    clean: List[str] = []

    for part in parts:
        part = _normalize(part)

        if not part:
            continue

        if clean and _same_text(clean[-1], part):
            continue

        clean.append(part)

    return " - ".join(clean) if clean else "Genel"


# =========================================================
# MASRAF ADI
# =========================================================

TRANSFER_TERMS = (
    "fast",
    "eft",
    "havale",
    "swift",
    "uftm",
    "fon transfer",
    "para transfer",
)


def _has_transfer_term(text: str, term: str) -> bool:
    """
    Transfer terimlerini kelime/sınır bazlı arar.

    Özellikle "EFT" için düz substring kullanmıyoruz;
    aksi halde "Defteri" kelimesinin içindeki "eft"
    yanlış pozitif üretir.
    """
    normalized = _normalize_key(text)

    patterns = {
        "fast": r"(?<![a-z0-9])fast(?![a-z0-9])",
        "eft": r"(?<![a-z0-9])eft(?![a-z0-9])",
        "havale": r"(?<![a-z0-9])havale(?![a-z0-9])",
        "swift": r"(?<![a-z0-9])swift(?![a-z0-9])",
        "uftm": r"(?<![a-z0-9])uftm(?![a-z0-9])",
        "fon transfer": r"fon\s+transfer",
        "para transfer": r"para\s+transfer",
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


def _build_masraf(
    raw_masraf: str,
    ana_kategori: str,
    alt_kategori: str,
    sub_kategori: str,
    aciklama: str,
) -> str:
    """
    Excel'de MASRAF filtresine EFT / FAST / Havale / SWIFT yazıldığında
    ilgili satırların bulunabilmesi için en yakın transfer başlığını
    gerektiğinde masrafın önüne ekler.
    """

    raw_masraf = _normalize(raw_masraf)
    raw_key = _normalize_key(raw_masraf)

    # En spesifik başlıktan en genele.
    title_candidates = [
        _normalize(sub_kategori),
        _normalize(alt_kategori),
        _normalize(ana_kategori),
    ]

    prefix = ""

    if not _contains_transfer_term(raw_masraf):
        for title in title_candidates:
            if not title:
                continue

            if not _contains_transfer_term(title):
                continue

            title_key = _normalize_key(title)

            if title_key in raw_key or raw_key in title_key:
                continue

            prefix = title
            break

    masraf = raw_masraf

    if prefix:
        masraf = f"{prefix} - {raw_masraf}"

    # SWIFT bazen yalnızca açıklama veya üst başlıkta yazabilir.
    if (
        "swift" in _normalize_key(aciklama)
        and "swift" not in _normalize_key(masraf)
    ):
        masraf = f"SWIFT - {masraf}"

    return _normalize(masraf)


# =========================================================
# DOM BAĞLAMI
# =========================================================

def _text_of(el) -> str:
    return _normalize(el.get_text(" ", strip=True)) if el else ""


def _get_group_divs(soup):
    """
    Eski kod:
        for hi in range(1, 20):
            grup_div = soup.find(id=f"h{hi}")
            if not grup_div:
                break

    Bir numara eksikse sonraki bütün grupları kaybedebilirdi.

    Burada gerçekten var olan h<number> id'lerini bulup sıralıyoruz.
    """

    groups = []

    for tag in soup.find_all(id=re.compile(r"^h\d+$")):
        match = re.fullmatch(r"h(\d+)", str(tag.get("id", "")))
        if not match:
            continue

        groups.append((int(match.group(1)), tag))

    groups.sort(key=lambda item: item[0])

    return [tag for _, tag in groups]


def _fallback_context(blok) -> Tuple[str, str, str]:
    """
    Normal hiyerarşi içinde yakalanmayan bir ücret bloğu varsa,
    kendisinden önceki en yakın group/item/sub başlıklarını kullan.
    """

    group_el = blok.find_previous(class_="UHU_group_header")
    item_el = blok.find_previous(class_="UHU_item_header")
    sub_el = blok.find_previous(class_="UHU_itemSub_header")

    return (
        _text_of(group_el) or "Genel",
        _text_of(item_el),
        _text_of(sub_el),
    )


# =========================================================
# TEK ÜCRET BLOĞUNU PARSE ET
# =========================================================

def _parse_block(
    blok,
    ana_kategori: str,
    alt_kategori: str,
    sub_kategori: str,
) -> Optional[UcretSatiri]:

    masraf_el = blok.find(class_="UHU_item_icerikH")
    raw_masraf = _text_of(masraf_el)

    if not raw_masraf:
        return None

    icerik1 = blok.find(class_="UHU_item_icerik1")
    icerik2 = blok.find(class_="UHU_item_icerik2")
    icerik3 = blok.find(class_="UHU_item_icerik3")
    icerik4 = blok.find(class_="UHU_item_icerik4")
    icerik5 = blok.find(class_="UHU_item_icerik5")
    aciklama_el = blok.find(class_="UHU_item_icerikF")

    asgari_tutar = _meta(icerik1)
    asgari_oran = _meta(icerik2)
    azami_tutar = _meta(icerik3)
    azami_oran = _meta(icerik4)
    tarih = _meta(icerik5, cls="UHU_icerik_meta2")
    aciklama = _text_of(aciklama_el)

    tarih = tarih.replace("/", ".")

    kategori = _join_category(
        [ana_kategori, alt_kategori, sub_kategori]
    )

    masraf = _build_masraf(
        raw_masraf=raw_masraf,
        ana_kategori=ana_kategori,
        alt_kategori=alt_kategori,
        sub_kategori=sub_kategori,
        aciklama=aciklama,
    )

    return UcretSatiri(
        kategori=kategori,
        masraf=masraf,
        asgari_tutar=asgari_tutar,
        asgari_oran=asgari_oran,
        azami_tutar=azami_tutar,
        azami_oran=azami_oran,
        aciklama=aciklama,
        site_guncelleme_tarihi=tarih,
    )


# =========================================================
# SOUP PARSER
# =========================================================

def _parse_soup(soup, source_name: str) -> Tuple[List[UcretSatiri], Dict[str, int]]:
    all_blocks = soup.find_all(class_="UHU_item_icerikC")
    group_divs = _get_group_divs(soup)

    stats: Dict[str, int] = {
        "groups": len(group_divs),
        "page_blocks": len(all_blocks),
        "structured_seen": 0,
        "fallback_seen": 0,
        "parsed_before_dedup": 0,
        "missing_masraf": 0,
        "duplicates": 0,
    }

    results: List[UcretSatiri] = []

    seen_rows: Set[
        Tuple[str, str, str, str, str, str, str, str]
    ] = set()

    processed_block_ids: Set[int] = set()
    kategori_sayilari: Dict[str, int] = {}

    def add_block(
        blok,
        ana_kategori: str,
        alt_kategori: str,
        sub_kategori: str,
        fallback: bool = False,
    ) -> None:
        block_id = id(blok)

        if block_id in processed_block_ids:
            return

        processed_block_ids.add(block_id)

        if fallback:
            stats["fallback_seen"] += 1
        else:
            stats["structured_seen"] += 1

        row = _parse_block(
            blok,
            ana_kategori=ana_kategori,
            alt_kategori=alt_kategori,
            sub_kategori=sub_kategori,
        )

        if row is None:
            stats["missing_masraf"] += 1
            return

        stats["parsed_before_dedup"] += 1

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

        if key in seen_rows:
            stats["duplicates"] += 1
            return

        seen_rows.add(key)
        results.append(row)

        ana = _normalize(ana_kategori) or "Genel"
        kategori_sayilari[ana] = kategori_sayilari.get(ana, 0) + 1

    # -----------------------------------------------------
    # NORMAL HİYERARŞİ
    # -----------------------------------------------------

    for grup_div in group_divs:
        ana_kategori_el = grup_div.find(class_="UHU_group_header")
        ana_kategori = _text_of(ana_kategori_el) or "Genel"

        item_headers = grup_div.find_all(class_="UHU_item_header")

        # Eğer group içinde item header yoksa blokları direkt fallback mantığıyla al.
        if not item_headers:
            for blok in grup_div.find_all(class_="UHU_item_icerikC"):
                add_block(
                    blok,
                    ana_kategori=ana_kategori,
                    alt_kategori="",
                    sub_kategori="",
                )
            continue

        for item_el in item_headers:
            alt_kategori = _text_of(item_el)

            item_sub_cover = item_el.find_next_sibling(
                id="UHU_itemSubCover"
            )

            if not item_sub_cover:
                item_sub_cover = item_el.parent

            if not item_sub_cover:
                continue

            sub_headers = item_sub_cover.find_all(
                class_="UHU_itemSub_header"
            )

            if sub_headers:
                for sub_el in sub_headers:
                    sub_kategori = _text_of(sub_el)

                    icerik_gc = sub_el.find_next_sibling(
                        id="UHU_item_icerik_GC"
                    )

                    if not icerik_gc:
                        icerik_gc = sub_el.parent

                    if not icerik_gc:
                        continue

                    for blok in icerik_gc.find_all(
                        class_="UHU_item_icerikC"
                    ):
                        add_block(
                            blok,
                            ana_kategori=ana_kategori,
                            alt_kategori=alt_kategori,
                            sub_kategori=sub_kategori,
                        )

            else:
                for blok in item_sub_cover.find_all(
                    class_="UHU_item_icerikC"
                ):
                    add_block(
                        blok,
                        ana_kategori=ana_kategori,
                        alt_kategori=alt_kategori,
                        sub_kategori="",
                    )

    # -----------------------------------------------------
    # ORPHAN / HİYERARŞİ DIŞI BLOKLAR
    # -----------------------------------------------------

    for blok in all_blocks:
        if id(blok) in processed_block_ids:
            continue

        ana, alt, sub = _fallback_context(blok)

        add_block(
            blok,
            ana_kategori=ana,
            alt_kategori=alt,
            sub_kategori=sub,
            fallback=True,
        )

    # -----------------------------------------------------
    # RAPOR
    # -----------------------------------------------------

    print(
        f"[isbank][{source_name}] Grup: {stats['groups']}",
        file=sys.stderr,
    )
    print(
        f"[isbank][{source_name}] Sayfadaki toplam ücret bloğu: "
        f"{stats['page_blocks']}",
        file=sys.stderr,
    )
    print(
        f"[isbank][{source_name}] Normal hiyerarşide görülen: "
        f"{stats['structured_seen']}",
        file=sys.stderr,
    )
    print(
        f"[isbank][{source_name}] Fallback ile bulunan: "
        f"{stats['fallback_seen']}",
        file=sys.stderr,
    )
    print(
        f"[isbank][{source_name}] Parse edilen (dedup öncesi): "
        f"{stats['parsed_before_dedup']}",
        file=sys.stderr,
    )
    print(
        f"[isbank][{source_name}] Duplicate: "
        f"{stats['duplicates']}",
        file=sys.stderr,
    )
    print(
        f"[isbank][{source_name}] Masraf adı olmayan blok: "
        f"{stats['missing_masraf']}",
        file=sys.stderr,
    )
    print(
        f"[isbank][{source_name}] Benzersiz sonuç: "
        f"{len(results)}",
        file=sys.stderr,
    )

    print("", file=sys.stderr)
    print(
        f"[isbank][{source_name}] ===== KATEGORİ RAPORU =====",
        file=sys.stderr,
    )

    for kategori, count in sorted(
        kategori_sayilari.items(),
        key=lambda item: (-item[1], item[0]),
    ):
        print(
            f"[isbank][{source_name}] {kategori} -> {count} kayıt",
            file=sys.stderr,
        )

    print(
        f"[isbank][{source_name}] ===========================",
        file=sys.stderr,
    )

    return results, stats


# =========================================================
# PLAYWRIGHT
# =========================================================

def _scrape_with_playwright(
    url: str = ISBANK_URL,
) -> Tuple[List[UcretSatiri], Dict[str, int]]:

    try:
        from playwright.sync_api import sync_playwright
        from bs4 import BeautifulSoup
    except ImportError as exc:
        raise ScraperError(
            "Playwright veya BeautifulSoup kurulu değil."
        ) from exc

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)

        try:
            context = browser.new_context(
                user_agent=HEADERS["User-Agent"],
                viewport={"width": 1440, "height": 1000},
                locale="tr-TR",
                timezone_id="Europe/Istanbul",
                extra_http_headers={
                    "Accept-Language": "tr-TR,tr;q=0.9,en;q=0.8",
                },
            )

            context.add_init_script(
                "Object.defineProperty("
                "navigator, 'webdriver', {get: () => undefined}"
                ")"
            )

            page = context.new_page()

            page.goto(
                url,
                timeout=120000,
                wait_until="domcontentloaded",
            )

            # Mevcut scraper 15 sn sabit bekliyordu.
            # Önce gerçek veri bloğunu bekleyelim; sonra kısa ek bekleme.
            try:
                page.wait_for_selector(
                    ".UHU_item_icerikC",
                    timeout=45000,
                )
            except Exception:
                pass

            page.wait_for_timeout(3000)

            # Sayfada lazy render varsa aşağı doğru ilerle.
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

                    if bottom and height == previous_height:
                        stable += 1
                    else:
                        stable = 0

                    previous_height = height

                    if stable >= 4:
                        break

                except Exception:
                    # Scroll debug işi scraper'ı düşürmesin.
                    break

            page.wait_for_timeout(1000)
            html = page.content()

        finally:
            browser.close()

    soup = BeautifulSoup(html, "lxml")

    return _parse_soup(
        soup,
        source_name="playwright",
    )


# =========================================================
# REQUESTS FALLBACK
# =========================================================

def _scrape_with_requests(
    url: str = ISBANK_URL,
) -> Tuple[List[UcretSatiri], Dict[str, int]]:

    import requests
    from bs4 import BeautifulSoup

    response = requests.get(
        url,
        headers=HEADERS,
        timeout=40,
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
# RAPORLAR
# =========================================================

def _print_integrity_report(
    stats: Dict[str, int],
    result_count: int,
) -> None:

    page_blocks = stats.get("page_blocks", 0)

    seen_total = (
        stats.get("structured_seen", 0)
        + stats.get("fallback_seen", 0)
    )

    parsed = stats.get("parsed_before_dedup", 0)
    missing = stats.get("missing_masraf", 0)

    print("", file=sys.stderr)
    print(
        "[isbank] ===== BÜTÜNLÜK KONTROLÜ =====",
        file=sys.stderr,
    )
    print(
        f"[isbank] Sayfadaki toplam ücret bloğu: {page_blocks}",
        file=sys.stderr,
    )
    print(
        f"[isbank] Parser'ın gördüğü toplam blok: {seen_total}",
        file=sys.stderr,
    )
    print(
        f"[isbank] Parse edilen (dedup öncesi): {parsed}",
        file=sys.stderr,
    )
    print(
        f"[isbank] Duplicate: {stats.get('duplicates', 0)}",
        file=sys.stderr,
    )
    print(
        f"[isbank] Masraf adı olmayan blok: {missing}",
        file=sys.stderr,
    )
    print(
        f"[isbank] Excel'e gidecek benzersiz satır: {result_count}",
        file=sys.stderr,
    )

    if page_blocks == seen_total and seen_total == parsed + missing:
        print(
            "[isbank] BÜTÜNLÜK: OK - "
            "sayfadaki tüm ücret blokları açıklandı.",
            file=sys.stderr,
        )
    else:
        print(
            "[isbank] BÜTÜNLÜK: UYARI - "
            "bazı ücret blokları açıklanamadı.",
            file=sys.stderr,
        )

    print(
        "[isbank] ==============================",
        file=sys.stderr,
    )


def _print_transfer_report(
    rows: List[UcretSatiri],
) -> None:

    print("", file=sys.stderr)
    print(
        "[isbank] ===== PARA AKTARMA KONTROLÜ =====",
        file=sys.stderr,
    )

    checks = [
        ("FAST", "fast"),
        ("EFT", "eft"),
        ("Havale", "havale"),
        ("SWIFT", "swift"),
    ]

    # Özellikle MASRAF sütununu kontrol ediyoruz.
    # Çünkü Excel'de kullanıcı bu sütundan filtre yapıyor.
    for label, needle in checks:
        found = [
            row
            for row in rows
            if _has_transfer_term(row.masraf, needle)
        ]

        print(
            f"[isbank] {label}: {len(found)} kayıt",
            file=sys.stderr,
        )

        for row in found[:10]:
            print(
                f"    - {row.masraf}",
                file=sys.stderr,
            )

    print(
        "[isbank] =================================",
        file=sys.stderr,
    )


# =========================================================
# ANA SCRAPER
# =========================================================

def scrape_isbank(
    url: str = ISBANK_URL,
) -> List[UcretSatiri]:

    print(
        f"[isbank] SÜRÜM: {SCRAPER_VERSION}",
        file=sys.stderr,
    )

    print(
        f"[isbank] {url} adresinden veri çekiliyor...",
        file=sys.stderr,
    )

    rows: List[UcretSatiri] = []
    stats: Optional[Dict[str, int]] = None

    # İş Bankası sayfasında mevcut çalışan yöntem Playwright olduğu için
    # onu birincil tutuyoruz.
    try:
        rows, stats = _scrape_with_playwright(url)

        if rows:
            print(
                f"[isbank] Playwright ile "
                f"{len(rows)} benzersiz satır bulundu.",
                file=sys.stderr,
            )

    except Exception as exc:
        print(
            f"[isbank] Playwright başarısız: {exc}",
            file=sys.stderr,
        )

    # Playwright veri üretmezse requests fallback.
    if not rows:
        try:
            rows, stats = _scrape_with_requests(url)

            if rows:
                print(
                    f"[isbank] requests ile "
                    f"{len(rows)} benzersiz satır bulundu.",
                    file=sys.stderr,
                )

        except Exception as exc:
            print(
                f"[isbank] requests başarısız: {exc}",
                file=sys.stderr,
            )

    if not rows:
        raise ScraperError(
            "İş Bankası sayfasında hiç veri satırı çekilemedi."
        )

    rows = sorted(
        rows,
        key=lambda s: (
            _normalize_key(s.kategori),
            _normalize_key(s.masraf),
        ),
    )

    if stats:
        _print_integrity_report(
            stats,
            len(rows),
        )

    _print_transfer_report(rows)

    print(
        f"[isbank] Toplam {len(rows)} benzersiz satır bulundu.",
        file=sys.stderr,
    )

    return rows


# =========================================================
# TEST
# =========================================================

if __name__ == "__main__":
    try:
        veriler = scrape_isbank()

        print()
        print("=" * 70)
        print("İŞ BANKASI SCRAPER")
        print("=" * 70)
        print(f"Toplam çekilen ücret: {len(veriler)}")
        print()

        for i, v in enumerate(veriler[:40], start=1):
            print(f"{i}. [{v.kategori}] {v.masraf}")
            print(f"   Asgari Tutar : {v.asgari_tutar}")
            print(f"   Asgari Oran  : {v.asgari_oran}")
            print(f"   Azami Tutar  : {v.azami_tutar}")
            print(f"   Azami Oran   : {v.azami_oran}")
            print(f"   Açıklama     : {v.aciklama}")
            print(f"   Tarih        : {v.site_guncelleme_tarihi}")
            print("-" * 70)

    except Exception as exc:
        print(
            f"[isbank][HATA] {exc}",
            file=sys.stderr,
        )
        sys.exit(1)
