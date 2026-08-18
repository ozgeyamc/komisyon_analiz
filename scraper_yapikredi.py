"""
Yapı Kredi Bireysel Ürün ve Hizmet Ücretleri scraper.

Yapı Kredi ücret sayfasındaki tüm tabloları Playwright ile toplar.

Özellikler:
- Tüm accordion bölümlerini açar
- Lazy-load içerikleri tetikler
- rowspan / colspan destekler
- Çok satırlı header'ları normalize eder
- Tabloya en yakın kategori/accordion başlığını bulur
- Ücret satırlarını standart UcretSatiri formatına dönüştürür
- Duplicate kayıtları kontrol eder
- Kategori bazlı debug çıktısı üretir
"""

import re
import sys
from dataclasses import dataclass
from typing import List, Dict, Set, Tuple, Optional


# =========================================================
# SABİTLER
# =========================================================

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
# TARİH
# =========================================================

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

    raw = normalize(raw_aciklama)

    if not raw:
        return "", ""

    match = DATE_PATTERN.search(raw)

    if match:

        tarih = match.group(1)

        temiz = DATE_PATTERN.sub(
            "",
            raw
        )

        temiz = normalize(
            temiz
        ).strip(
            " .:-"
        )

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

            tarih = (
                f"{gun}.{ay}.{yil}"
            )

            temiz = DATE_PATTERN_TR.sub(
                "",
                raw
            )

            temiz = normalize(
                temiz
            ).strip(
                " .:-"
            )

            return temiz, tarih

    return raw, ""


# =========================================================
# NORMALİZASYON
# =========================================================

def normalize(
    value: Optional[str]
) -> str:

    if value is None:
        return ""

    value = str(value)

    value = value.replace(
        "\xa0",
        " "
    )

    value = value.replace(
        "\u200b",
        ""
    )

    value = value.replace(
        "\r",
        " "
    )

    value = value.replace(
        "\n",
        " "
    )

    value = re.sub(
        r"\s+",
        " ",
        value
    )

    return value.strip()


def normalize_header(
    value: Optional[str]
) -> str:

    value = normalize(
        value
    ).lower()

    replacements = {
        "ı": "i",
        "ğ": "g",
        "ü": "u",
        "ş": "s",
        "ö": "o",
        "ç": "c",
    }

    for old, new in replacements.items():

        value = value.replace(
            old,
            new
        )

    value = value.replace(
        "%",
        " "
    )

    value = re.sub(
        r"\([^)]*\)",
        " ",
        value
    )

    value = re.sub(
        r"\s+",
        " ",
        value
    )

    return value.strip()


def normalize_category(
    value: Optional[str]
) -> str:

    value = normalize(
        value
    )

    if not value:
        return "Yapı Kredi"

    # Kategoriye dönüşmemesi gereken başlıklar
    invalid = {
        "asgari tutar",
        "azami tutar",
        "açıklama",
        "açıklamalar",
        "masraf",
        "ücret",
        "oran",
        "tutar",
        "güncelleme tarihi",
        "güncellenme tarihi",
    }

    if normalize_header(
        value
    ) in {
        normalize_header(x)
        for x in invalid
    }:
        return "Yapı Kredi"

    return value


# =========================================================
# ÜCRET KONTROLÜ
# =========================================================

def looks_like_amount(
    value: str
) -> bool:

    value = normalize(
        value
    ).lower()

    if not value:
        return False

    if value in {
        "-",
        "—",
        "–",
    }:
        return True

    if "tl" in value:
        return True

    if "%" in value:
        return True

    if "usd" in value:
        return True

    if "eur" in value:
        return True

    if re.fullmatch(
        r"[-+]?\d+(?:[.,]\d+)?",
        value
    ):
        return True

    return False


def row_has_money_data(
    row: List[str]
) -> bool:

    for cell in row:

        cell = normalize(cell)

        if not cell:
            continue

        if looks_like_amount(
            cell
        ):
            return True

    return False


# =========================================================
# HÜCRE
# =========================================================

def get_cell(
    row: List[str],
    index: int
) -> str:

    if index < 0:
        return ""

    if index >= len(row):
        return ""

    return normalize(
        row[index]
    )


# =========================================================
# HEADER BULMA
# =========================================================

def find_col(
    headers: List[str],
    keywords: List[str]
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


def find_first(
    headers: List[str],
    candidates: List[List[str]]
) -> int:

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

    best_index = -1
    best_score = 0

    for index, row in enumerate(
        rows[:12]
    ):

        text = " ".join(
            normalize_header(
                cell
            )
            for cell in row
        )

        score = 0

        keywords = [
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
            "komisyon",
            "musteri",
        ]

        for keyword in keywords:

            if keyword in text:
                score += 1

        if score > best_score:

            best_score = score
            best_index = index

    if best_index == -1:
        return 0

    return best_index


def merge_header_rows(
    rows: List[List[str]],
    header_index: int
) -> List[str]:
    """
    Çok satırlı header'ları birleştirir.

    Örneğin:

    ["Asgari", "Azami", "Açıklama"]
    ["Tutar", "Oran", ""]

    ->

    ["Asgari Tutar", "Azami Oran", "Açıklama"]
    """

    if header_index < 0:
        return []

    header = [
        normalize(x)
        for x in rows[header_index]
    ]

    # Bir sonraki satır gerçekten header
    # devamı mı kontrol et.
    if (
        header_index + 1
        >= len(rows)
    ):
        return header

    next_row = rows[
        header_index + 1
    ]

    next_text = " ".join(
        normalize_header(x)
        for x in next_row
    )

    header_keywords = [
        "tutar",
        "oran",
        "tl",
        "%",
        "aciklama",
        "guncelleme",
        "guncellenme",
    ]

    if not any(
        keyword in next_text
        for keyword in header_keywords
    ):
        return header

    merged = []

    max_len = max(
        len(header),
        len(next_row)
    )

    for i in range(max_len):

        first = (
            header[i]
            if i < len(header)
            else ""
        )

        second = (
            normalize(next_row[i])
            if i < len(next_row)
            else ""
        )

        if first and second:

            merged.append(
                normalize(
                    f"{first} {second}"
                )
            )

        elif first:

            merged.append(
                first
            )

        else:

            merged.append(
                second
            )

    return merged


def find_columns(
    header_row: List[str]
) -> Dict[str, int]:

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
            ["ücret"],
            ["komisyon"],
            ["islem"],
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


# =========================================================
# MASRAF BUL
# =========================================================

def find_best_masraf(
    row: List[str],
    column_map: Dict[str, int]
) -> str:

    masraf_index = column_map[
        "masraf"
    ]

    # Gerçek masraf kolonu bulunduysa
    # doğrudan kullan.
    if masraf_index >= 0:

        value = get_cell(
            row,
            masraf_index
        )

        if value:
            return value

    used_indices = {
        index
        for index in column_map.values()
        if index >= 0
    }

    candidates = []

    for index, cell in enumerate(
        row
    ):

        cell = normalize(
            cell
        )

        if not cell:
            continue

        if index in used_indices:
            continue

        if looks_like_amount(
            cell
        ):
            continue

        candidates.append(
            cell
        )

    if not candidates:
        return ""

    # Açıklama çok uzunsa onu masraf olarak
    # seçmemek için ilk anlamlı hücreyi al.
    return candidates[0]


# =========================================================
# PLAYWRIGHT TABLO TOPLAMA
# =========================================================

def collect_tables(
    url: str
):

    from playwright.sync_api import (
        sync_playwright
    )

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

            # =================================================
            # SAYFA
            # =================================================

            page.goto(
                url,
                timeout=120000,
                wait_until="domcontentloaded"
            )

            page.wait_for_timeout(
                4000
            )

            # =================================================
            # COOKIE
            # =================================================

            cookie_buttons = [
                "Tümünü Kabul Et",
                "Tümünü Kabul",
                "Kabul Et",
                "Kabul",
            ]

            for text in cookie_buttons:

                try:

                    locator = page.get_by_text(
                        text,
                        exact=True
                    ).first

                    if locator.is_visible(
                        timeout=1000
                    ):

                        locator.click(
                            timeout=3000
                        )

                        print(
                            f"[yapikredi] "
                            f"Cookie butonu: {text}",
                            file=sys.stderr
                        )

                        page.wait_for_timeout(
                            1000
                        )

                        break

                except Exception:
                    pass

            # =================================================
            # ACCORDION
            # =================================================

            print(
                "[yapikredi] Accordion'lar açılıyor...",
                file=sys.stderr
            )

            accordion_selectors = [

                "[aria-expanded='false']",

                "button[data-toggle='collapse']",

                "a[data-toggle='collapse']",

                "[data-bs-toggle='collapse']",

                ".accordion-title",

                ".accordionItem-title",

                ".accordion-header",

                ".collapsible-header",

                ".js-accordion-title",

            ]

            for round_no in range(8):

                opened = 0

                for selector in accordion_selectors:

                    try:

                        elements = page.locator(
                            selector
                        )

                        count = elements.count()

                        for i in range(count):

                            try:

                                element = elements.nth(
                                    i
                                )

                                if not element.is_visible(
                                    timeout=200
                                ):
                                    continue

                                aria = (
                                    element.get_attribute(
                                        "aria-expanded"
                                    )
                                )

                                if aria == "true":
                                    continue

                                element.scroll_into_view_if_needed(
                                    timeout=1000
                                )

                                element.click(
                                    timeout=2000,
                                    force=True
                                )

                                opened += 1

                                page.wait_for_timeout(
                                    120
                                )

                            except Exception:
                                pass

                    except Exception:
                        pass

                if opened == 0:
                    break

                print(
                    f"[yapikredi] "
                    f"Accordion turu "
                    f"{round_no + 1}: "
                    f"{opened} adet açıldı.",
                    file=sys.stderr
                )

                page.wait_for_timeout(
                    500
                )

            # =================================================
            # LAZY LOAD
            # =================================================

            print(
                "[yapikredi] Sayfa taranıyor...",
                file=sys.stderr
            )

            stable = 0
            previous_height = 0

            for _ in range(150):

                page.evaluate(
                    """
                    window.scrollBy(
                        0,
                        Math.max(
                            window.innerHeight * 0.75,
                            500
                        )
                    );
                    """
                )

                page.wait_for_timeout(
                    250
                )

                height = page.evaluate(
                    "document.body.scrollHeight"
                )

                bottom = page.evaluate(
                    """
                    () =>
                        window.innerHeight +
                        window.scrollY >=
                        document.body.scrollHeight - 100
                    """
                )

                if (
                    bottom
                    and height == previous_height
                ):

                    stable += 1

                else:

                    stable = 0

                previous_height = height

                if stable >= 5:
                    break

            page.evaluate(
                """
                window.scrollTo(
                    0,
                    document.body.scrollHeight
                );
                """
            )

            page.wait_for_timeout(
                1500
            )

            # =================================================
            # TABLOLAR
            # =================================================

            javascript = r"""
            () => {

                function clean(value) {

                    return (value || "")
                        .replace(/\u00a0/g, " ")
                        .replace(/\u200b/g, "")
                        .replace(/\r/g, " ")
                        .replace(/\n/g, " ")
                        .replace(/\s+/g, " ")
                        .trim();
                }


                // =================================================
                // EN YAKIN KATEGORİYİ BUL
                // =================================================

                function getCategory(table) {

                    /*
                     * ÖNEMLİ:
                     *
                     * Eski kodda parent container'ın
                     * bütün heading'leri aranıyordu.
                     *
                     * Bu nedenle örneğin ilk accordion
                     * başlığı "Para Aktarma" ise,
                     * altındaki bütün tablolar Para Aktarma
                     * oluyordu.
                     *
                     * Burada sadece:
                     *
                     * 1. tabloya doğrudan bağlı başlıklar
                     * 2. tabloyu içeren accordion
                     * 3. accordion'ın kendi header'ı
                     *
                     * kontrol ediliyor.
                     */

                    let current = table;

                    for (
                        let level = 0;
                        level < 12;
                        level++
                    ) {

                        if (
                            !current.parentElement
                        ) {
                            break;
                        }

                        const parent =
                            current.parentElement;


                        // -------------------------------------------------
                        // Accordion container
                        // -------------------------------------------------

                        const accordion =
                            parent.closest(
                                ".accordion-item, " +
                                ".accordionItem, " +
                                ".accordion, " +
                                ".collapsible, " +
                                "[data-accordion]"
                            );

                        if (accordion) {

                            const header =
                                accordion.querySelector(
                                    ":scope > .accordion-header, " +
                                    ":scope > .accordion-title, " +
                                    ":scope > .accordionItem-title, " +
                                    ":scope > button, " +
                                    ":scope > h2, " +
                                    ":scope > h3, " +
                                    ":scope > h4"
                                );

                            if (header) {

                                const text =
                                    clean(
                                        header.innerText
                                    );

                                if (
                                    text &&
                                    text.length < 150
                                ) {

                                    return text;
                                }
                            }
                        }


                        // -------------------------------------------------
                        // Table'ın hemen üstündeki heading
                        // -------------------------------------------------

                        const directHeadings =
                            Array.from(
                                parent.children
                            ).filter(
                                element =>
                                    /^H[1-6]$/.test(
                                        element.tagName
                                    ) ||
                                    element.matches(
                                        ".title, " +
                                        ".accordion-title, " +
                                        ".accordionItem-title, " +
                                        ".section-title"
                                    )
                            );

                        if (
                            directHeadings.length
                        ) {

                            const heading =
                                directHeadings[
                                    directHeadings.length - 1
                                ];

                            const text =
                                clean(
                                    heading.innerText
                                );

                            if (
                                text &&
                                text.length < 150
                            ) {

                                return text;
                            }
                        }


                        // -------------------------------------------------
                        // aria-controls ile ilişkili accordion
                        // -------------------------------------------------

                        const id =
                            parent.id;

                        if (id) {

                            const controller =
                                document.querySelector(
                                    `[aria-controls="${id}"]`
                                );

                            if (controller) {

                                const text =
                                    clean(
                                        controller.innerText
                                    );

                                if (
                                    text &&
                                    text.length < 150
                                ) {

                                    return text;
                                }
                            }
                        }


                        current = parent;
                    }

                    return "Yapı Kredi";
                }


                // =================================================
                // ROWSPAN / COLSPAN
                // =================================================

                function getRows(table) {

                    const trs =
                        Array.from(
                            table.querySelectorAll(
                                "tr"
                            )
                        );

                    const matrix = [];

                    const pending = {};

                    let outputRow = 0;

                    for (
                        const tr of trs
                    ) {

                        if (
                            tr.closest("table")
                            !== table
                        ) {
                            continue;
                        }

                        const cells =
                            Array.from(
                                tr.children
                            ).filter(
                                element =>
                                    element.tagName === "TD" ||
                                    element.tagName === "TH"
                            );

                        if (!cells.length) {
                            continue;
                        }

                        if (!matrix[outputRow]) {
                            matrix[outputRow] = [];
                        }

                        let colIndex = 0;

                        for (
                            const cell of cells
                        ) {

                            while (
                                pending[
                                    `${outputRow}:${colIndex}`
                                ]
                            ) {

                                colIndex++;
                            }

                            const text =
                                clean(
                                    cell.innerText
                                );

                            const colspan =
                                Math.max(
                                    parseInt(
                                        cell.getAttribute(
                                            "colspan"
                                        ) || "1",
                                        10
                                    ),
                                    1
                                );

                            const rowspan =
                                Math.max(
                                    parseInt(
                                        cell.getAttribute(
                                            "rowspan"
                                        ) || "1",
                                        10
                                    ),
                                    1
                                );

                            for (
                                let r = 0;
                                r < rowspan;
                                r++
                            ) {

                                for (
                                    let c = 0;
                                    c < colspan;
                                    c++
                                ) {

                                    const targetRow =
                                        outputRow + r;

                                    const targetCol =
                                        colIndex + c;

                                    if (
                                        !matrix[targetRow]
                                    ) {
                                        matrix[targetRow] = [];
                                    }

                                    matrix[targetRow][
                                        targetCol
                                    ] = text;

                                    if (r > 0) {

                                        pending[
                                            `${targetRow}:${targetCol}`
                                        ] = true;
                                    }
                                }
                            }

                            colIndex += colspan;
                        }

                        outputRow++;
                    }

                    return matrix
                        .map(
                            row =>
                                row.map(
                                    cell =>
                                        clean(
                                            cell || ""
                                        )
                                )
                        )
                        .filter(
                            row =>
                                row.some(
                                    cell =>
                                        cell !== ""
                                )
                        );
                }


                // =================================================
                // ROOT TABLE
                // =================================================

                const allTables =
                    Array.from(
                        document.querySelectorAll(
                            "table"
                        )
                    );

                const tables =
                    allTables.filter(
                        table =>
                            !table.parentElement.closest(
                                "table"
                            )
                    );


                // =================================================
                // RESULT
                // =================================================

                return tables.map(
                    (table, index) => {

                        return {
                            index: index,

                            kategori:
                                getCategory(
                                    table
                                ),

                            rows:
                                getRows(
                                    table
                                )
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

            total_rows = sum(
                len(
                    x.get(
                        "rows",
                        []
                    )
                )
                for x in tables
            )

            print(
                f"[yapikredi] "
                f"Toplam tablo satırı: "
                f"{total_rows}",
                file=sys.stderr
            )

            return tables

        finally:

            context.close()
            browser.close()


# =========================================================
# TABLOLARI PARSE ET
# =========================================================

def parse_tables(
    tables
) -> List[UcretSatiri]:

    sonuc = []

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

    tablo_sayisi = 0
    atlanan_tablo = 0

    kategori_sayilari: Dict[
        str,
        int
    ] = {}

    kategori_tablolari: Dict[
        str,
        int
    ] = {}

    # =========================================================
    # TABLOLAR
    # =========================================================

    for table in tables:

        rows = table.get(
            "rows",
            []
        )

        if not rows:
            atlanan_tablo += 1
            continue

        tablo_sayisi += 1

        kategori = normalize_category(
            table.get(
                "kategori",
                "Yapı Kredi"
            )
        )

        kategori_tablolari[
            kategori
        ] = (
            kategori_tablolari.get(
                kategori,
                0
            ) + 1
        )

        # =====================================================
        # HEADER
        # =====================================================

        header_index = find_header_index(
            rows
        )

        if (
            header_index < 0
            or header_index >= len(rows)
        ):

            atlanan_tablo += 1

            print(
                f"[yapikredi][DEBUG] "
                f"Tablo {table.get('index')} "
                f"header bulunamadı. "
                f"Kategori: {kategori}",
                file=sys.stderr
            )

            continue

        header = merge_header_rows(
            rows,
            header_index
        )

        header_text = " ".join(
            normalize_header(
                x
            )
            for x in header
        )

        # =====================================================
        # HEADER KONTROLÜ
        # =====================================================

        header_keywords = [
            "asgari",
            "azami",
            "tutar",
            "oran",
            "aciklama",
            "guncelleme",
            "guncellenme",
            "masraf",
            "ucret",
            "komisyon",
            "islem",
            "kanal",
            "bsmv",
        ]

        if not any(
            keyword in header_text
            for keyword in header_keywords
        ):

            atlanan_tablo += 1

            print(
                f"[yapikredi][DEBUG] "
                f"Tablo {table.get('index')} "
                f"header ücret tablosu olarak "
                f"tanınmadı. "
                f"Kategori: {kategori}",
                file=sys.stderr
            )

            continue

        column_map = find_columns(
            header
        )

        # =====================================================
        # DATA
        # =====================================================

        tablo_kaydi = 0

        for row in rows[
            header_index + 1:
        ]:

            if not row:
                continue

            row = [
                normalize(cell)
                for cell in row
            ]

            if not any(row):
                continue

            # -------------------------------------------------
            # TEKRAR HEADER
            # -------------------------------------------------

            row_text = " ".join(
                normalize_header(
                    x
                )
                for x in row
            )

            if (
                "asgari" in row_text
                and "azami" in row_text
                and (
                    "tutar" in row_text
                    or "oran" in row_text
                )
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
                column_map[
                    "asgari_tutar"
                ]
            )

            asgari_oran = get_cell(
                row,
                column_map[
                    "asgari_oran"
                ]
            )

            azami_tutar = get_cell(
                row,
                column_map[
                    "azami_tutar"
                ]
            )

            azami_oran = get_cell(
                row,
                column_map[
                    "azami_oran"
                ]
            )

            # -------------------------------------------------
            # AÇIKLAMA
            # -------------------------------------------------

            aciklama_raw = get_cell(
                row,
                column_map[
                    "aciklama"
                ]
            )

            aciklama, tarih = (
                parse_aciklama(
                    aciklama_raw
                )
            )

            # -------------------------------------------------
            # TABLO TARİHİ
            # -------------------------------------------------

            tablo_tarihi = get_cell(
                row,
                column_map[
                    "tarih"
                ]
            )

            if tablo_tarihi:

                tarih = (
                    tablo_tarihi
                )

            # -------------------------------------------------
            # AÇIKLAMA KOLONU YOKSA
            # -------------------------------------------------

            if (
                not aciklama
                and column_map[
                    "aciklama"
                ] == -1
            ):

                used = {
                    i
                    for i in column_map.values()
                    if i >= 0
                }

                extra = []

                for i, cell in enumerate(
                    row
                ):

                    if i in used:
                        continue

                    if not cell:
                        continue

                    if looks_like_amount(
                        cell
                    ):
                        continue

                    extra.append(
                        cell
                    )

                if extra:

                    aciklama = " ".join(
                        extra
                    )

                    (
                        aciklama,
                        extra_date
                    ) = parse_aciklama(
                        aciklama
                    )

                    if not tarih:
                        tarih = extra_date

            # -------------------------------------------------
            # BU SATIR GERÇEKTEN VERİ Mİ?
            # -------------------------------------------------

            has_fee = any([
                asgari_tutar,
                asgari_oran,
                azami_tutar,
                azami_oran,
            ])

            has_description = bool(
                aciklama
            )

            has_date = bool(
                tarih
            )

            has_money = (
                row_has_money_data(
                    row
                )
            )

            if not (
                has_fee
                or has_description
                or has_date
                or has_money
            ):
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
                tarih,
            )

            if key in seen:
                continue

            seen.add(
                key
            )

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

            tablo_kaydi += 1

            kategori_sayilari[
                kategori
            ] = (
                kategori_sayilari.get(
                    kategori,
                    0
                ) + 1
            )

        # =====================================================
        # TABLO DEBUG
        # =====================================================

        if tablo_kaydi == 0:

            print(
                f"[yapikredi][DEBUG] "
                f"Tablo {table.get('index')} "
                f"0 kayıt üretti | "
                f"Kategori: {kategori} | "
                f"Satır: {len(rows)}",
                file=sys.stderr
            )

    # =========================================================
    # ÖZET
    # =========================================================

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

    # =========================================================
    # KATEGORİ RAPORU
    # =========================================================

    print(
        "",
        file=sys.stderr
    )

    print(
        "[yapikredi] ===== KATEGORİ RAPORU =====",
        file=sys.stderr
    )

    for kategori, count in sorted(
        kategori_sayilari.items(),
        key=lambda x: (
            -x[1],
            x[0]
        )
    ):

        tablo_count = (
            kategori_tablolari.get(
                kategori,
                0
            )
        )

        print(
            f"[yapikredi] "
            f"{kategori} -> "
            f"{count} kayıt / "
            f"{tablo_count} tablo",
            file=sys.stderr
        )

    print(
        "[yapikredi] ===========================",
        file=sys.stderr
    )

    return sonuc


# =========================================================
# ANA FONKSİYON
# =========================================================

def scrape_yapikredi(
    url: str = YAPIKREDI_URL
) -> List[UcretSatiri]:

    tables = collect_tables(
        url
    )

    if not tables:

        raise ScraperError(
            "Yapı Kredi sayfasından "
            "tablo alınamadı."
        )

    sonuc = parse_tables(
        tables
    )

    if not sonuc:

        raise ScraperError(
            "Tablolar bulundu fakat "
            "ücret satırı oluşturulamadı."
        )

    return sonuc


# =========================================================
# TEST
# =========================================================

if __name__ == "__main__":

    try:

        sonuc = scrape_yapikredi()

        print()
        print("=" * 70)
        print("YAPI KREDİ SCRAPER")
        print("=" * 70)

        print(
            f"Toplam çekilen ücret: "
            f"{len(sonuc)}"
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

            print(
                "-" * 70
            )

    except Exception as exc:

        print(
            f"[yapikredi][HATA] {exc}",
            file=sys.stderr
        )

        sys.exit(1)
