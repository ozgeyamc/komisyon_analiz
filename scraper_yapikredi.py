"""
Yapı Kredi Ürün ve Hizmet Ücretleri scraper.

- Playwright ile ücret sayfasındaki tüm tabloları toplar.
- Ana H2 başlığını kategori olarak kullanır.
- Accordion/alt başlık + tablo başlığı + satır adını MASRAF alanında birleştirir.
- FAST / EFT / Havale / SWIFT gibi başlıkların Excel filtresinde kaybolmasını engeller.
- Gerçek veri satırlarını yanlışlıkla tekrar header sanıp atlamaz.
- Aynı isimli ücretleri farklı alt kategorilerde yanlışlıkla duplicate saymaz.
- main.py / update_excel.py ile uyumlu UcretSatiri listesi döndürür.
"""

import re
import sys
from dataclasses import dataclass
from typing import Dict, List, Optional, Set, Tuple


SCRAPER_VERSION = "2026-08-18-v8-complete-fee-hierarchy"

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
    r"(?:son\s+)?(?:güncellenme|güncelleme)\s*tarihi\s*:?\s*"
    r"(\d{1,2}[./]\d{1,2}[./]\d{4}(?:\s+\d{1,2}:\d{2})?)",
    re.IGNORECASE,
)

DATE_PATTERN_TR = re.compile(
    r"(?:son\s+)?(?:güncellenme|güncelleme)\s+tarihi\s*:?\s*"
    r"(\d{1,2})\s+"
    r"(ocak|şubat|mart|nisan|mayıs|haziran|temmuz|ağustos|eylül|ekim|kasım|aralık)\s+"
    r"(\d{4})",
    re.IGNORECASE,
)


# =========================================================
# METİN YARDIMCILARI
# =========================================================


def normalize(value: Optional[str]) -> str:
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
    value = re.sub(r"\([^)]*\)", " ", value)
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def identity(value: Optional[str]) -> str:
    """Duplicate / başlık karşılaştırması için daha hafif normalize."""
    return normalize_header(value)


def parse_aciklama(raw_aciklama: str) -> Tuple[str, str]:
    raw = normalize(raw_aciklama)
    if not raw:
        return "", ""

    match = DATE_PATTERN.search(raw)
    if match:
        tarih = match.group(1)
        temiz = normalize(DATE_PATTERN.sub("", raw)).strip(" .:-")
        return temiz, tarih

    match = DATE_PATTERN_TR.search(raw)
    if match:
        gun = match.group(1).zfill(2)
        ay = TURKCE_AYLAR.get(match.group(2).lower(), "")
        yil = match.group(3)
        if ay:
            tarih = f"{gun}.{ay}.{yil}"
            temiz = normalize(DATE_PATTERN_TR.sub("", raw)).strip(" .:-")
            return temiz, tarih

    return raw, ""


def get_cell(row: List[str], index: int) -> str:
    if index < 0 or index >= len(row):
        return ""
    return normalize(row[index])


# =========================================================
# HEADER / KOLON TESPİTİ
# =========================================================


def find_col(headers: List[str], keywords: List[str]) -> int:
    for index, header in enumerate(headers):
        if all(keyword in header for keyword in keywords):
            return index
    return -1


def header_score(row: List[str]) -> int:
    """
    Yalnızca HEADER hücrelerinin kendisini puanlar.
    Data satırındaki uzun açıklamada geçen 'asgari tutar' vb.
    ifadeler yüzünden satırın header sayılmasını önlemek için bu fonksiyon
    sadece find_header_index sırasında kullanılır.
    """
    headers = [normalize_header(cell) for cell in row]

    tests = [
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
        if any(all(k in h for k in keywords) for h in headers):
            score += 1
    return score


def find_header_index(rows: List[List[str]]) -> int:
    best_index = -1
    best_score = 0

    for index, row in enumerate(rows[:10]):
        score = header_score(row)
        if score > best_score:
            best_index = index
            best_score = score

    # Standart ücret tablosu olduğunu anlamak için en az 3 header sinyali.
    if best_score < 3:
        return -1

    return best_index


def find_columns(header_row: List[str]) -> Dict[str, int]:
    headers = [normalize_header(h) for h in header_row]

    result = {
        "masraf": 0 if header_row else -1,
        "asgari_tutar": find_col(headers, ["asgari", "tutar"]),
        "asgari_oran": find_col(headers, ["asgari", "oran"]),
        "azami_tutar": find_col(headers, ["azami", "tutar"]),
        "azami_oran": find_col(headers, ["azami", "oran"]),
        "aciklama": find_col(headers, ["aciklama"]),
        "tarih": find_col(headers, ["guncelleme", "tarihi"]),
    }

    if result["tarih"] == -1:
        result["tarih"] = find_col(headers, ["guncellenme", "tarihi"])

    # Yapı Kredi standart ücret tablosunda kolonlar bu sırada.
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
        for key, fallback_index in fallbacks.items():
            if result[key] == -1:
                result[key] = fallback_index

    return result


def is_repeated_header(row: List[str], header: List[str]) -> bool:
    """
    Gerçek veri satırlarını yanlışlıkla header diye silmez.

    Eski mantık header_score(row) >= 3 idi. Bu, örneğin YP Teminat
    Mektubu açıklamasında 'azami oran' ve 'asgari tutar' geçtiği için
    gerçek veri satırını atabiliyordu.
    """
    row_norm = [identity(x) for x in row]
    header_norm = [identity(x) for x in header]

    # Birebir aynı satırsa kesin header tekrarı.
    if row_norm == header_norm:
        return True

    exact_header_cells = {
        "asgari tutar",
        "asgari oran",
        "azami tutar",
        "azami oran",
        "aciklama",
        "guncelleme tarihi",
        "guncellenme tarihi",
    }

    exact_matches = sum(1 for cell in row_norm if cell in exact_header_cells)
    return exact_matches >= 4


# =========================================================
# MASRAF ADI OLUŞTURMA
# =========================================================


def _add_unique_component(parts: List[str], value: str) -> None:
    value = normalize(value).strip(" -–—|")
    if not value:
        return

    value_id = identity(value)
    if not value_id:
        return

    # Aynı başlığı iki kez ekleme.
    for existing in parts:
        existing_id = identity(existing)
        if value_id == existing_id:
            return

        # Örn. alt başlık 'EFT', tablo başlığı 'EFT Gönderimi'.
        # Daha açıklayıcı olan uzun metin zaten EFT içeriyorsa kısa olanı
        # ayrıca eklemek zorunda değiliz.
        if value_id in existing_id and len(value_id) < len(existing_id):
            return

    # Yeni parça, var olan kısa parçayı kapsıyorsa kısa parçayı çıkar.
    parts[:] = [
        existing
        for existing in parts
        if not (
            identity(existing) in value_id
            and len(identity(existing)) < len(value_id)
        )
    ]

    parts.append(value)


def build_masraf_name(
    alt_kategori: str,
    tablo_basligi: str,
    satir_masrafi: str,
) -> str:
    """
    Excel'deki MASRAF alanına anlamlı hiyerarşi koyar.

    Örnekler:
      FAST Gönderim - İnternet/Mobil – 0 - 8.300 TL
      EFT Gönderimi - Geç - İnternet/Mobil – 0 - 8.300 TL
      Havale Gönderimi - ATM – 0 - 8.300 TL
      Uluslararası Fon Transferi ve Mesajlaşma Ücreti / SWIFT -
          Döviz Havale Gönderimi - Şubeden - Aynı Gün Valörlü
    """
    parts: List[str] = []

    _add_unique_component(parts, alt_kategori)
    _add_unique_component(parts, tablo_basligi)
    _add_unique_component(parts, satir_masrafi)

    return " - ".join(parts)


# =========================================================
# PLAYWRIGHT / DOM
# =========================================================


def collect_tables(url: str):
    from playwright.sync_api import sync_playwright

    print(f"[yapikredi] SÜRÜM: {SCRAPER_VERSION}", file=sys.stderr)
    print(f"[yapikredi] Sayfa açılıyor: {url}", file=sys.stderr)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent=USER_AGENT,
            viewport={"width": 1440, "height": 1080},
            locale="tr-TR",
            extra_http_headers={
                "Accept-Language": "tr-TR,tr;q=0.9,en;q=0.8",
            },
        )
        page = context.new_page()

        try:
            page.goto(url, timeout=120000, wait_until="domcontentloaded")
            page.wait_for_timeout(3500)

            # Cookie
            for text in [
                "Tümünü Kabul Et",
                "Tümünü Kabul",
                "Kabul Et",
                "Kabul",
            ]:
                try:
                    locator = page.get_by_text(text, exact=True).first
                    if locator.is_visible(timeout=800):
                        locator.click(timeout=2500)
                        print(
                            f"[yapikredi] Cookie butonu: {text}",
                            file=sys.stderr,
                        )
                        page.wait_for_timeout(500)
                        break
                except Exception:
                    pass

            print("[yapikredi] Accordion'lar açılıyor...", file=sys.stderr)

            # Sadece gerçekten kapalı accordion kontrollerini aç.
            for round_no in range(10):
                opened = 0
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
                        page.wait_for_timeout(80)
                    except Exception:
                        pass

                if opened == 0:
                    break

                print(
                    f"[yapikredi] Accordion turu {round_no + 1}: "
                    f"{opened} adet açıldı.",
                    file=sys.stderr,
                )
                page.wait_for_timeout(300)

            print("[yapikredi] Sayfa taranıyor...", file=sys.stderr)

            # Lazy load için kademeli scroll.
            stable = 0
            previous_height = 0

            for _ in range(180):
                page.evaluate(
                    "window.scrollBy(0, Math.max(window.innerHeight * 0.75, 500));"
                )
                page.wait_for_timeout(180)

                height = page.evaluate("document.body.scrollHeight")
                bottom = page.evaluate(
                    "window.innerHeight + window.scrollY >= "
                    "document.body.scrollHeight - 120"
                )

                if bottom and height == previous_height:
                    stable += 1
                else:
                    stable = 0

                previous_height = height
                if stable >= 5:
                    break

            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            page.wait_for_timeout(1200)

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

                function nearestPreviousHeading(table, tagName) {
                    const xpath = `preceding::${tagName}[1]`;
                    const result = document.evaluate(
                        xpath,
                        table,
                        null,
                        XPathResult.FIRST_ORDERED_NODE_TYPE,
                        null
                    );
                    return result.singleNodeValue;
                }

                function getCategory(table) {
                    const h2 = nearestPreviousHeading(table, "h2");
                    if (h2) {
                        const text = clean(h2.innerText);
                        if (text) return text;
                    }

                    const h1 = nearestPreviousHeading(table, "h1");
                    if (h1) {
                        const text = clean(h1.innerText);
                        if (text) return text;
                    }

                    return "Yapı Kredi";
                }

                function getSubcategory(table) {
                    // Tabloya en yakın önceki accordion başlığı / buton / h5.
                    const xpath =
                        "preceding::*[self::h5 or self::h4 or self::button][1]";
                    const result = document.evaluate(
                        xpath,
                        table,
                        null,
                        XPathResult.FIRST_ORDERED_NODE_TYPE,
                        null
                    );
                    const node = result.singleNodeValue;
                    return node ? clean(node.innerText) : "";
                }

                function getRows(table) {
                    const trs = Array.from(table.querySelectorAll("tr"))
                        .filter(tr => tr.closest("table") === table);

                    const matrix = [];
                    const spans = new Map();
                    let rowIndex = 0;

                    for (const tr of trs) {
                        const cells = Array.from(tr.children).filter(
                            el => el.tagName === "TD" || el.tagName === "TH"
                        );
                        if (!cells.length) continue;

                        if (!matrix[rowIndex]) matrix[rowIndex] = [];
                        let colIndex = 0;

                        function fillPending() {
                            while (spans.has(`${rowIndex}:${colIndex}`)) {
                                const val = spans.get(`${rowIndex}:${colIndex}`);
                                matrix[rowIndex][colIndex] = val;
                                colIndex++;
                            }
                        }

                        fillPending();

                        for (const cell of cells) {
                            fillPending();

                            const text = clean(cell.innerText);
                            const rowspan = Math.max(
                                parseInt(cell.getAttribute("rowspan") || "1", 10),
                                1
                            );
                            const colspan = Math.max(
                                parseInt(cell.getAttribute("colspan") || "1", 10),
                                1
                            );

                            for (let c = 0; c < colspan; c++) {
                                matrix[rowIndex][colIndex + c] = text;
                            }

                            if (rowspan > 1) {
                                for (let r = 1; r < rowspan; r++) {
                                    for (let c = 0; c < colspan; c++) {
                                        spans.set(
                                            `${rowIndex + r}:${colIndex + c}`,
                                            text
                                        );
                                    }
                                }
                            }

                            colIndex += colspan;
                        }

                        // Satırın sonundaki pending rowspan hücrelerini de ekle.
                        fillPending();
                        rowIndex++;
                    }

                    return matrix
                        .map(row => row.map(cell => clean(cell || "")))
                        .filter(row => row.some(cell => cell !== ""));
                }

                const allTables = Array.from(document.querySelectorAll("table"));
                const rootTables = allTables.filter(
                    table => !table.parentElement.closest("table")
                );

                return rootTables.map((table, index) => ({
                    index,
                    kategori: getCategory(table),
                    alt_kategori: getSubcategory(table),
                    rows: getRows(table),
                }));
            }
            """

            tables = page.evaluate(javascript)

            print(
                f"[yapikredi] {len(tables)} adet tablo bulundu.",
                file=sys.stderr,
            )

            total_rows = sum(len(t.get("rows", [])) for t in tables)
            print(
                f"[yapikredi] Toplam tablo satırı: {total_rows}",
                file=sys.stderr,
            )

            return tables

        finally:
            context.close()
            browser.close()


# =========================================================
# PARSE
# =========================================================


def parse_tables(tables) -> List[UcretSatiri]:
    sonuc: List[UcretSatiri] = []

    # alt_kategori + tablo başlığı duplicate anahtarına dahil.
    seen: Set[
        Tuple[str, str, str, str, str, str, str, str, str, str]
    ] = set()

    kategori_sayilari: Dict[str, int] = {}

    fee_table_count = 0
    ignored_table_count = 0
    zero_record_tables = 0

    candidate_rows = 0
    parsed_before_dedup = 0
    duplicate_rows = 0
    repeated_headers = 0
    note_rows = 0
    invalid_rows = 0

    for table in tables:
        rows = table.get("rows", []) or []
        table_index = table.get("index")
        kategori = normalize(table.get("kategori")) or "Yapı Kredi"
        alt_kategori = normalize(table.get("alt_kategori"))

        if not rows:
            ignored_table_count += 1
            continue

        header_index = find_header_index(rows)

        if header_index == -1:
            ignored_table_count += 1
            continue

        fee_table_count += 1

        header = [normalize(x) for x in rows[header_index]]
        column_map = find_columns(header)
        tablo_basligi = get_cell(header, 0)

        table_record_count = 0

        for row in rows[header_index + 1:]:
            row = [normalize(cell) for cell in row]

            if not row or not any(row):
                continue

            candidate_rows += 1

            # Sadece gerçek header tekrarıysa atla.
            if is_repeated_header(row, header):
                repeated_headers += 1
                continue

            satir_masrafi = get_cell(row, column_map["masraf"])

            asgari_tutar = get_cell(row, column_map["asgari_tutar"])
            asgari_oran = get_cell(row, column_map["asgari_oran"])
            azami_tutar = get_cell(row, column_map["azami_tutar"])
            azami_oran = get_cell(row, column_map["azami_oran"])
            aciklama_raw = get_cell(row, column_map["aciklama"])
            site_tarihi = get_cell(row, column_map["tarih"])

            aciklama, aciklama_tarihi = parse_aciklama(aciklama_raw)
            if not site_tarihi:
                site_tarihi = aciklama_tarihi

            meaningful_cells = sum(1 for cell in row if normalize(cell))
            if meaningful_cells < 2:
                note_rows += 1
                continue

            # İlk hücre rowspan/boşluk yüzünden boş gelebilirse tablo başlığını kullan.
            if not satir_masrafi:
                satir_masrafi = tablo_basligi

            if not satir_masrafi:
                invalid_rows += 1
                continue

            # Satır gerçekten ücret/veri taşıyor mu?
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
                invalid_rows += 1
                continue

            masraf = build_masraf_name(
                alt_kategori=alt_kategori,
                tablo_basligi=tablo_basligi,
                satir_masrafi=satir_masrafi,
            )

            if not masraf:
                invalid_rows += 1
                continue

            parsed_before_dedup += 1

            key = (
                kategori,
                alt_kategori,
                tablo_basligi,
                satir_masrafi,
                asgari_tutar,
                asgari_oran,
                azami_tutar,
                azami_oran,
                aciklama,
                site_tarihi,
            )

            if key in seen:
                duplicate_rows += 1
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
                    site_guncelleme_tarihi=site_tarihi,
                )
            )

            table_record_count += 1
            kategori_sayilari[kategori] = kategori_sayilari.get(kategori, 0) + 1

        if table_record_count == 0:
            zero_record_tables += 1
            print(
                f"[yapikredi][DEBUG] Ücret tablosu {table_index} 0 kayıt üretti | "
                f"Kategori: {kategori} | Alt başlık: {alt_kategori} | "
                f"Tablo başlığı: {tablo_basligi} | "
                f"Satır: {len(rows)} | Header: {header}",
                file=sys.stderr,
            )

    print(f"[yapikredi] Ücret tablosu: {fee_table_count}", file=sys.stderr)
    print(
        f"[yapikredi] İlgisiz/atlanan tablo: {ignored_table_count}",
        file=sys.stderr,
    )
    print(
        f"[yapikredi] 0 kayıt üreten ücret tablosu: {zero_record_tables}",
        file=sys.stderr,
    )
    print(
        f"[yapikredi] Toplam benzersiz ücret: {len(sonuc)}",
        file=sys.stderr,
    )

    print("", file=sys.stderr)
    print("[yapikredi] ===== KATEGORİ RAPORU =====", file=sys.stderr)
    for kategori, count in sorted(
        kategori_sayilari.items(),
        key=lambda item: (-item[1], item[0]),
    ):
        print(f"[yapikredi] {kategori} -> {count} kayıt", file=sys.stderr)
    print("[yapikredi] ===========================", file=sys.stderr)

    # Para aktarma özel kontrolü.
    print("", file=sys.stderr)
    print("[yapikredi] ===== PARA AKTARMA KONTROLÜ =====", file=sys.stderr)
    for keyword in ["FAST", "EFT", "Havale", "SWIFT"]:
        bulunan = [
            item
            for item in sonuc
            if keyword.casefold() in item.masraf.casefold()
        ]
        print(
            f"[yapikredi] {keyword}: {len(bulunan)} kayıt",
            file=sys.stderr,
        )
        for item in bulunan[:8]:
            print(f"    - {item.masraf}", file=sys.stderr)
    print("[yapikredi] =================================", file=sys.stderr)

    print("", file=sys.stderr)
    print("[yapikredi] ===== BÜTÜNLÜK KONTROLÜ =====", file=sys.stderr)
    print(f"[yapikredi] Ham veri satırı adayı: {candidate_rows}", file=sys.stderr)
    print(
        f"[yapikredi] Parse edilen (dedup öncesi): {parsed_before_dedup}",
        file=sys.stderr,
    )
    print(f"[yapikredi] Duplicate: {duplicate_rows}", file=sys.stderr)
    print(f"[yapikredi] Tekrarlanan header: {repeated_headers}", file=sys.stderr)
    print(f"[yapikredi] Not/dipnot satırı: {note_rows}", file=sys.stderr)
    print(f"[yapikredi] Geçersiz/değersiz satır: {invalid_rows}", file=sys.stderr)
    print(f"[yapikredi] Excel'e gidecek satır: {len(sonuc)}", file=sys.stderr)

    accounted = (
        parsed_before_dedup
        + repeated_headers
        + note_rows
        + invalid_rows
    )

    if accounted == candidate_rows:
        print(
            "[yapikredi] BÜTÜNLÜK: OK - aday satırların tamamı açıklandı.",
            file=sys.stderr,
        )
    else:
        print(
            f"[yapikredi] BÜTÜNLÜK: UYARI - "
            f"{candidate_rows - accounted} aday satır açıklanamıyor.",
            file=sys.stderr,
        )

    print("[yapikredi] ==============================", file=sys.stderr)

    return sonuc


# =========================================================
# ANA FONKSİYON
# =========================================================


def scrape_yapikredi(
    url: str = YAPIKREDI_URL,
) -> List[UcretSatiri]:
    tables = collect_tables(url)

    if not tables:
        raise ScraperError("Yapı Kredi sayfasından hiç tablo alınamadı.")

    sonuc = parse_tables(tables)

    if not sonuc:
        raise ScraperError(
            "Tablolar bulundu fakat hiçbir ücret satırı oluşturulamadı."
        )

    return sonuc


if __name__ == "__main__":
    try:
        sonuc = scrape_yapikredi()

        print()
        print("=" * 70)
        print("YAPI KREDİ SCRAPER")
        print(f"SÜRÜM: {SCRAPER_VERSION}")
        print("=" * 70)
        print(f"Toplam çekilen ücret: {len(sonuc)}")
        print()

        for i, satir in enumerate(sonuc[:40], start=1):
            print(f"{i}. [{satir.kategori}] {satir.masraf}")
            print(f"   Asgari Tutar : {satir.asgari_tutar}")
            print(f"   Asgari Oran  : {satir.asgari_oran}")
            print(f"   Azami Tutar  : {satir.azami_tutar}")
            print(f"   Azami Oran   : {satir.azami_oran}")
            print(f"   Açıklama     : {satir.aciklama}")
            print(f"   Tarih        : {satir.site_guncelleme_tarihi}")
            print("-" * 70)

    except Exception as exc:
        print(f"[yapikredi][HATA] {exc}", file=sys.stderr)
        sys.exit(1)
