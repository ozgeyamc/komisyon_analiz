"""
Yapı Kredi Bireysel Ürün ve Hizmet Ücretleri scraper.

Sayfadaki tüm HTML tablolarını Playwright ile toplar,
tablo başlıklarını analiz eder ve bütün ücret satırlarını
standart UcretSatiri formatına dönüştürür.
"""

import re
import sys
from dataclasses import dataclass
from typing import List, Dict, Set, Tuple, Optional


YAPIKREDI_URL = (
    "https://www.yapikredi.com.tr/"
    "bireysel-bankacilik/hesaplama-araclari/"
    "bireysel-urun-ve-hizmet-ucretleri"
)

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36"
)


# ---------------------------------------------------------
# VERİ YAPISI
# ---------------------------------------------------------

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


# ---------------------------------------------------------
# TARİH
# ---------------------------------------------------------

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


DATE_PATTERN = re.compile(
    r"(?:güncellenme|güncelleme)\s*tarihi\s*:?\s*"
    r"(\d{1,2}[./]\d{1,2}[./]\d{4}"
    r"(?:\s+\d{1,2}:\d{2})?)",
    re.IGNORECASE,
)


DATE_PATTERN_TR = re.compile(
    r"(?:son\s+)?"
    r"(?:güncellenme|güncelleme)\s+tarihi\s*:?\s*"
    r"(\d{1,2})\s+"
    r"(ocak|şubat|mart|nisan|mayıs|haziran|temmuz|"
    r"ağustos|eylül|ekim|kasım|aralık)\s+"
    r"(\d{4})",
    re.IGNORECASE,
)


def parse_aciklama(raw_aciklama: str):
    """
    Açıklamanın içindeki güncelleme tarihini ayırır.
    """

    raw = normalize(raw_aciklama)

    if not raw:
        return "", ""

    match = DATE_PATTERN.search(raw)

    if match:
        tarih = match.group(1)

        temiz = DATE_PATTERN.sub("", raw)
        temiz = normalize(temiz).strip(" .:-")

        return temiz, tarih

    match = DATE_PATTERN_TR.search(raw)

    if match:
        gun = match.group(1).zfill(2)
        ay = TURKCE_AYLAR.get(
            match.group(2).lower(),
            ""
        )
        yil = match.group(3)

        if ay:
            tarih = f"{gun}.{ay}.{yil}"

            temiz = DATE_PATTERN_TR.sub("", raw)
            temiz = normalize(temiz).strip(" .:-")

            return temiz, tarih

    return raw, ""


# ---------------------------------------------------------
# GENEL YARDIMCI FONKSİYONLAR
# ---------------------------------------------------------

def normalize(value: Optional[str]) -> str:
    """
    HTML'den gelen metni temizler.
    """

    if value is None:
        return ""

    value = str(value)

    value = value.replace("\xa0", " ")
    value = value.replace("\u200b", "")
    value = value.replace("\r", " ")
    value = value.replace("\n", " ")

    value = re.sub(r"\s+", " ", value)

    return value.strip()


def normalize_header(value: Optional[str]) -> str:
    """
    Header karşılaştırması için Türkçe karakterleri sadeleştirir.
    """

    value = normalize(value).lower()

    replacements = {
        "ı": "i",
        "ğ": "g",
        "ü": "u",
        "ş": "s",
        "ö": "o",
        "ç": "c",
    }

    for old, new in replacements.items():
        value = value.replace(old, new)

    value = value.replace("%", " ")

    # Parantez içlerini temizle
    value = re.sub(r"\([^)]*\)", " ", value)

    value = re.sub(r"\s+", " ", value)

    return value.strip()


def looks_like_amount(value: str) -> bool:
    """
    Hücrenin ücret/oran hücresi olup olmadığını kontrol eder.
    """

    value = normalize(value).lower()

    if not value:
        return False

    if value in {"-", "—", "–"}:
        return True

    if "tl" in value:
        return True

    if "%" in value:
        return True

    # 10
    # 10,50
    # 1.000
    # 0,275
    if re.fullmatch(
        r"[-+]?\d+(?:[.,]\d+)?",
        value
    ):
        return True

    return False


def get_cell(row: List[str], index: int) -> str:
    """
    Güvenli hücre alma.
    """

    if index < 0:
        return ""

    if index >= len(row):
        return ""

    return normalize(row[index])


# ---------------------------------------------------------
# HEADER ANALİZİ
# ---------------------------------------------------------

def find_col(
    headers: List[str],
    keywords: List[str]
) -> int:
    """
    Header içerisinde verilen kelimelerin tamamını arar.

    ÖNEMLİ:
    Bulamazsa -1 döner.
    0 index'i özel olarak korunur.
    """

    for index, header in enumerate(headers):

        if all(
            keyword in header
            for keyword in keywords
        ):
            return index

    return -1


def find_first(
    headers: List[str],
    candidates: List[List[str]]
) -> int:
    """
    Aday header kombinasyonlarını sırayla dener.
    """

    for keywords in candidates:

        index = find_col(
            headers,
            keywords
        )

        if index != -1:
            return index

    return -1


def find_header_index(
    rows: List[List[str]]
) -> int:
    """
    Tablodaki gerçek header satırını bulur.
    """

    for index, row in enumerate(rows[:8]):

        text = " ".join(
            normalize_header(cell)
            for cell in row
        )

        score = 0

        for keyword in [
            "asgari",
            "azami",
            "tutar",
            "oran",
            "aciklama",
            "guncelleme",
            "guncellenme",
            "kanal",
            "bsmv",
            "masraf",
            "ucret",
            "islem",
        ]:
            if keyword in text:
                score += 1

        if score >= 2:
            return index

    return 0


def find_columns(
    header_row: List[str]
) -> Dict[str, int]:
    """
    Yapı Kredi tablosundaki kolonları bulur.
    """

    headers = [
        normalize_header(h)
        for h in header_row
    ]

    result = {
        "masraf": -1,
        "asgari_tutar": -1,
        "asgari_oran": -1,
        "azami_tutar": -1,
        "azami_oran": -1,
        "aciklama": -1,
        "tarih": -1,
    }

    # -----------------------------------------------------
    # MASRAF
    # -----------------------------------------------------

    result["masraf"] = find_first(
        headers,
        [
            ["masraf"],
            ["ucret"],
            ["islem"],
            ["komisyon"],
        ]
    )

    # -----------------------------------------------------
    # ASGARİ TUTAR
    # -----------------------------------------------------

    result["asgari_tutar"] = find_first(
        headers,
        [
            ["asgari", "tutar"],
        ]
    )

    # -----------------------------------------------------
    # ASGARİ ORAN
    # -----------------------------------------------------

    result["asgari_oran"] = find_first(
        headers,
        [
            ["asgari", "oran"],
        ]
    )

    # -----------------------------------------------------
    # AZAMİ TUTAR
    # -----------------------------------------------------

    result["azami_tutar"] = find_first(
        headers,
        [
            ["azami", "tutar"],
        ]
    )

    # -----------------------------------------------------
    # AZAMİ ORAN
    # -----------------------------------------------------

    result["azami_oran"] = find_first(
        headers,
        [
            ["azami", "oran"],
        ]
    )

    # -----------------------------------------------------
    # AÇIKLAMA
    # -----------------------------------------------------

    result["aciklama"] = find_first(
        headers,
        [
            ["aciklama"],
        ]
    )

    # -----------------------------------------------------
    # TARİH
    # -----------------------------------------------------

    result["tarih"] = find_first(
        headers,
        [
            ["guncelleme", "tarihi"],
            ["guncellenme", "tarihi"],
            ["guncelleme"],
            ["guncellenme"],
        ]
    )

    return result


# ---------------------------------------------------------
# MASRAF BULMA
# ---------------------------------------------------------

def find_best_masraf(
    row: List[str],
    column_map: Dict[str, int]
) -> str:
    """
    Masraf kolonu bulunamazsa satırdaki en mantıklı
    açıklayıcı hücreyi bulur.
    """

    masraf_index = column_map["masraf"]

    if masraf_index >= 0:

        value = get_cell(
            row,
            masraf_index
        )

        if value and not looks_like_amount(value):
            return value

    # Kullanılmış kolonları çıkar
    used_indices = {
        index
        for index in column_map.values()
        if index >= 0
    }

    candidates = []

    for index, cell in enumerate(row):

        cell = normalize(cell)

        if not cell:
            continue

        if index in used_indices:
            continue

        if looks_like_amount(cell):
            continue

        header_like = normalize_header(cell)

        if header_like in {
            "asgari tutar",
            "asgari oran",
            "azami tutar",
            "azami oran",
            "aciklama",
            "guncelleme tarihi",
        }:
            continue

        candidates.append(cell)

    if candidates:
        # En uzun anlamlı metni seç
        return max(
            candidates,
            key=len
        )

    return ""


# ---------------------------------------------------------
# PLAYWRIGHT İLE TABLOLARI TOPLA
# ---------------------------------------------------------

def collect_tables(url: str):
    from playwright.sync_api import sync_playwright

    print(
        f"[yapikredi] Sayfa açılıyor: {url}",
        file=sys.stderr
    )

    with sync_playwright() as p:

        browser = p.chromium.launch(
            headless=True
        )

        context = browser.new_context(
            user_agent=USER_AGENT,
            viewport={
                "width": 1440,
                "height": 1080,
            },
            locale="tr-TR",
            extra_http_headers={
                "Accept-Language":
                    "tr-TR,tr;q=0.9,en;q=0.8",
            },
        )

        page = context.new_page()

        try:

            page.goto(
                url,
                timeout=120000,
                wait_until="domcontentloaded"
            )

            page.wait_for_timeout(2500)

            # -------------------------------------------------
            # COOKIE
            # -------------------------------------------------

            cookie_buttons = [
                "Tümünü Kabul Et",
                "Tümünü Kabul",
                "Kabul Et",
                "Kabul",
                "Kapat",
            ]

            for text in cookie_buttons:

                try:

                    locator = page.locator(
                        f"button:has-text('{text}'), "
                        f"a:has-text('{text}')"
                    ).first

                    if locator.is_visible(
                        timeout=1000
                    ):

                        locator.click(
                            timeout=2000
                        )

                        print(
                            f"[yapikredi] "
                            f"Cookie butonu: {text}",
                            file=sys.stderr
                        )

                        page.wait_for_timeout(500)

                except Exception:
                    pass

            # -------------------------------------------------
            # ARIA ACCORDION
            # -------------------------------------------------

            print(
                "[yapikredi] Accordion'lar açılıyor...",
                file=sys.stderr
            )

            for _ in range(3):

                try:

                    elements = page.locator(
                        "[aria-expanded='false']"
                    )

                    count = elements.count()

                    if count == 0:
                        break

                    for i in range(count):

                        try:

                            item = elements.nth(i)

                            if item.is_visible(
                                timeout=300
                            ):

                                item.click(
                                    timeout=1500
                                )

                                page.wait_for_timeout(
                                    150
                                )

                        except Exception:
                            pass

                except Exception:
                    break

            # -------------------------------------------------
            # BİLİNEN ACCORDION'LAR
            # -------------------------------------------------

            selectors = [
                ".accordion-title",
                ".accordionItem-title",
                ".collapsible-header",
                "a[data-toggle='collapse']",
                ".js-accordion-title",
                "h2.title",
                "h3.title",
            ]

            for selector in selectors:

                try:

                    elements = page.locator(
                        selector
                    )

                    count = elements.count()

                    for i in range(count):

                        try:

                            element = elements.nth(i)

                            if element.is_visible(
                                timeout=300
                            ):

                                element.click(
                                    timeout=1000
                                )

                                page.wait_for_timeout(
                                    150
                                )

                        except Exception:
                            pass

                except Exception:
                    pass

            # -------------------------------------------------
            # TAM SAYFA SCROLL
            # -------------------------------------------------

            print(
                "[yapikredi] Sayfa taranıyor...",
                file=sys.stderr
            )

            last_height = 0

            for _ in range(40):

                height = page.evaluate(
                    "document.body.scrollHeight"
                )

                if height == last_height:
                    break

                last_height = height

                page.evaluate(
                    "window.scrollTo("
                    "0, document.body.scrollHeight"
                    ")"
                )

                page.wait_for_timeout(500)

            page.evaluate(
                "window.scrollTo(0, 0)"
            )

            page.wait_for_timeout(1000)

            # -------------------------------------------------
            # TABLE + KATEGORİ
            # -------------------------------------------------

            javascript = r"""
            () => {

                function clean(value) {

                    return (value || "")
                        .replace(/\u00a0/g, " ")
                        .replace(/\u200b/g, "")
                        .replace(/\s+/g, " ")
                        .trim();
                }

                function getRows(table) {

                    const rows = [];

                    const trs =
                        Array.from(
                            table.querySelectorAll("tr")
                        );

                    for (const tr of trs) {

                        const cells =
                            Array.from(
                                tr.querySelectorAll(
                                    "th, td"
                                )
                            );

                        if (!cells.length)
                            continue;

                        rows.push(
                            cells.map(
                                cell =>
                                    clean(
                                        cell.innerText
                                    )
                            )
                        );
                    }

                    return rows;
                }

                function getCategory(table) {

                    let current = table;

                    for (let level = 0;
                         level < 8;
                         level++) {

                        if (!current.parentElement)
                            break;

                        current =
                            current.parentElement;

                        const headings =
                            current.querySelectorAll(
                                "h1,h2,h3,h4,h5,h6,"
                                + "button,"
                                + ".accordion-title,"
                                + ".title"
                            );

                        for (
                            const heading
                            of headings
                        ) {

                            const text =
                                clean(
                                    heading.innerText
                                );

                            if (
                                text &&
                                text.length < 150
                            ) {

                                const lower =
                                    text.toLowerCase();

                                if (
                                    !lower.includes(
                                        "asgari tutar"
                                    ) &&
                                    !lower.includes(
                                        "azami tutar"
                                    )
                                ) {
                                    return text;
                                }
                            }
                        }
                    }

                    return "Yapı Kredi";
                }

                const tables =
                    Array.from(
                        document.querySelectorAll(
                            "table"
                        )
                    );

                return tables.map(
                    (table, index) => {

                        return {
                            index: index,
                            kategori:
                                getCategory(table),
                            rows:
                                getRows(table)
                        };

                    }
                );
            }
            """

            tables = page.evaluate(
                javascript
            )

            print(
                f"[yapikredi] "
                f"{len(tables)} adet tablo bulundu.",
                file=sys.stderr
            )

            return tables

        finally:

            context.close()
            browser.close()


# ---------------------------------------------------------
# TABLOLARI PARSE ET
# ---------------------------------------------------------

def parse_tables(
    tables
) -> List[UcretSatiri]:

    sonuc = []

    # Daha detaylı duplicate anahtarı
    seen: Set[
        Tuple[
            str,
            str,
            str,
            str,
            str,
            str,
            str,
        ]
    ] = set()

    tablo_sayisi = 0
    atlanan_tablo = 0

    for table in tables:

        rows = table.get(
            "rows",
            []
        )

        if not rows:
            continue

        tablo_sayisi += 1

        kategori = normalize(
            table.get(
                "kategori",
                "Yapı Kredi"
            )
        )

        if not kategori:
            kategori = "Yapı Kredi"

        # -------------------------------------------------
        # HEADER
        # -------------------------------------------------

        header_index = find_header_index(
            rows
        )

        header = rows[header_index]

        # Gerçek header değilse tabloyu kaybetmeden
        # yine de kontrol ediyoruz.
        header_text = " ".join(
            normalize_header(x)
            for x in header
        )

        if not any(
            keyword in header_text
            for keyword in [
                "asgari",
                "azami",
                "tutar",
                "oran",
                "aciklama",
                "guncelleme",
                "kanal",
                "bsmv",
                "masraf",
                "ucret",
                "islem",
            ]
        ):

            atlanan_tablo += 1
            continue

        column_map = find_columns(
            header
        )

        # -------------------------------------------------
        # DATA
        # -------------------------------------------------

        for row in rows[header_index + 1:]:

            if not row:
                continue

            row = [
                normalize(cell)
                for cell in row
            ]

            if not any(row):
                continue

            # Header tekrar etmişse atla
            row_text = " ".join(
                normalize_header(x)
                for x in row
            )

            if (
                "asgari" in row_text
                and "azami" in row_text
                and "tutar" in row_text
            ):
                continue

            # -------------------------------------------------
            # MASRAF
            # -------------------------------------------------

            masraf = find_best_masraf(
                row,
                column_map
            )

            if not masraf:
                continue

            # -------------------------------------------------
            # ÜCRETLER
            # -------------------------------------------------

            asgari_tutar = get_cell(
                row,
                column_map["asgari_tutar"]
            )

            asgari_oran = get_cell(
                row,
                column_map["asgari_oran"]
            )

            azami_tutar = get_cell(
                row,
                column_map["azami_tutar"]
            )

            azami_oran = get_cell(
                row,
                column_map["azami_oran"]
            )

            # -------------------------------------------------
            # AÇIKLAMA
            # -------------------------------------------------

            aciklama_raw = get_cell(
                row,
                column_map["aciklama"]
            )

            aciklama, tarih = parse_aciklama(
                aciklama_raw
            )

            # -------------------------------------------------
            # TARİH
            # -------------------------------------------------

            tablo_tarihi = get_cell(
                row,
                column_map["tarih"]
            )

            if tablo_tarihi:
                tarih = tablo_tarihi

            # -------------------------------------------------
            # AÇIKLAMA KOLONU YOKSA
            # -------------------------------------------------

            if (
                not aciklama
                and column_map["aciklama"] == -1
            ):

                used = {
                    i
                    for i in column_map.values()
                    if i >= 0
                }

                extra = []

                for i, cell in enumerate(row):

                    if i in used:
                        continue

                    if not cell:
                        continue

                    if looks_like_amount(cell):
                        continue

                    extra.append(cell)

                if extra:

                    aciklama = " ".join(
                        extra
                    )

                    aciklama, extra_date = (
                        parse_aciklama(
                            aciklama
                        )
                    )

                    if not tarih:
                        tarih = extra_date

            # -------------------------------------------------
            # SADECE TAMAMEN BOŞ ÜCRET SATIRIYSA ATLA
            # -------------------------------------------------

            if not any([
                asgari_tutar,
                asgari_oran,
                azami_tutar,
                azami_oran,
                aciklama,
                tarih,
            ]):
                continue

            # -------------------------------------------------
            # DUPLICATE
            # -------------------------------------------------

            key = (
                kategori,
                masraf,
                asgari_tutar,
                asgari_oran,
                azami_tutar,
                azami_oran,
                aciklama,
            )

            if key in seen:
                continue

            seen.add(key)

            sonuc.append(
                UcretSatiri(
                    kategori=kategori,
                    masraf=masraf,
                    asgari_tutar=asgari_tutar,
                    asgari_oran=asgari_oran,
                    azami_tutar=azami_tutar,
                    azami_oran=azami_oran,
                    aciklama=aciklama,
                    site_guncelleme_tarihi=tarih,
                )
            )

    print(
        f"[yapikredi] İşlenen tablo: "
        f"{tablo_sayisi}",
        file=sys.stderr
    )

    print(
        f"[yapikredi] Atlanan tablo: "
        f"{atlanan_tablo}",
        file=sys.stderr
    )

    print(
        f"[yapikredi] Toplam benzersiz ücret: "
        f"{len(sonuc)}",
        file=sys.stderr
    )

    return sonuc


# ---------------------------------------------------------
# ANA FONKSİYON
# ---------------------------------------------------------

def scrape_yapikredi(
    url: str = YAPIKREDI_URL
) -> List[UcretSatiri]:

    tables = collect_tables(url)

    if not tables:
        raise ScraperError(
            "Yapı Kredi sayfasından tablo alınamadı."
        )

    sonuc = parse_tables(
        tables
    )

    if not sonuc:
        raise ScraperError(
            "Tablolar bulundu fakat ücret satırı "
            "oluşturulamadı."
        )

    return sonuc


# ---------------------------------------------------------
# TEST
# ---------------------------------------------------------

if __name__ == "__main__":

    try:

        sonuc = scrape_yapikredi()

        print()
        print("=" * 70)
        print("YAPI KREDİ SCRAPER")
        print("=" * 70)

        print(
            f"Toplam çekilen ücret: {len(sonuc)}"
        )

        print()

        for i, satir in enumerate(
            sonuc[:30],
            start=1
        ):

            print(
                f"{i}. "
                f"[{satir.kategori}] "
                f"{satir.masraf}"
            )

            print(
                f"   Asgari Tutar : "
                f"{satir.asgari_tutar}"
            )

            print(
                f"   Asgari Oran  : "
                f"{satir.asgari_oran}"
            )

            print(
                f"   Azami Tutar  : "
                f"{satir.azami_tutar}"
            )

            print(
                f"   Azami Oran   : "
                f"{satir.azami_oran}"
            )

            print(
                f"   Açıklama     : "
                f"{satir.aciklama}"
            )

            print(
                f"   Tarih        : "
                f"{satir.site_guncelleme_tarihi}"
            )

            print("-" * 70)

    except Exception as exc:

        print(
            f"[yapikredi][HATA] {exc}",
            file=sys.stderr
        )

        sys.exit(1)
