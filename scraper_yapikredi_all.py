"""
scraper_yapikredi_all.py

Yapı Kredi "Bireysel Ürün ve Hizmet Ücretleri" sayfasından tüm tabloları çeker.
- Playwright ile sayfayı açar, cookie/dialog butonlarını kapatır.
- Accordion / sekme öğelerini açar, sayfayı kaydırarak dinamik tabloları yükler.
- JS ile her tablonun üst başlığını (Kategori) ve satırlarını çıkartır.
- Python tarafında header eşlemesi yapıp UcretSatiri nesneleri üretir.
- Sonucu Excel'e kaydeder.
"""

import re
import sys
from dataclasses import dataclass
from typing import List, Dict, Any, Set, Tuple

YAPIKREDI_URL = "https://www.yapikredi.com.tr/bireysel-bankacilik/hesaplama-araclari/bireysel-urun-ve-hizmet-ucretleri"

DATE_PATTERN = re.compile(
    r"G[üu]ncellenme\s*Tarihi\s*:\s*(\d{2}\.\d{2}\.\d{4}(?:\s+\d{2}:\d{2})?)",
    re.IGNORECASE,
)
DATE_PATTERN_TR = re.compile(
    r"(?:son\s+)?g[üu]ncellenme\s+tarihi\s*:?\s*"
    r"(\d{1,2})\s+(ocak|şubat|mart|nisan|mayıs|haziran|temmuz|ağustos|eylül|ekim|kasım|aralık)\s+(\d{4})",
    re.IGNORECASE,
)
TURKCE_AYLAR = {
    "ocak": "01", "şubat": "02", "mart": "03", "nisan": "04",
    "mayıs": "05", "haziran": "06", "temmuz": "07", "ağustos": "08",
    "eylül": "09", "ekim": "10", "kasım": "11", "aralık": "12",
}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
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


def _parse_aciklama(raw_aciklama: str):
    raw = (raw_aciklama or "").strip()
    match = DATE_PATTERN.search(raw)
    if match:
        tarih = match.group(1)
        temiz = DATE_PATTERN.sub("", raw).strip(" .")
        return temiz, tarih
    match_tr = DATE_PATTERN_TR.search(raw)
    if match_tr:
        gun = match_tr.group(1).zfill(2)
        ay = TURKCE_AYLAR.get(match_tr.group(2).lower(), "00")
        yil = match_tr.group(3)
        tarih = f"{gun}.{ay}.{yil}"
        temiz = DATE_PATTERN_TR.sub("", raw).strip(" .")
        return temiz, tarih
    return raw, ""


def _normalize(val: str) -> str:
    return (val or "").strip().replace("\xa0", " ").replace("\u200b", "").strip()


def _find_col_indices_from_headers(header_texts: List[str]) -> Dict[str, int]:
    def norm(h: str) -> str:
        h2 = re.sub(r"\(.*?\)", "", (h or ""))
        h2 = h2.replace("%", " ")
        h2 = re.sub(r"\s+", " ", h2).strip().lower()
        return h2

    headers_norm = [norm(h) for h in header_texts]

    def find_col(keywords: List[str]) -> int:
        for i, h in enumerate(headers_norm):
            if all(k in h for k in keywords):
                return i
        return -1

    col_masraf = find_col(["masraf"])
    if col_masraf == -1:
        col_masraf = find_col(["işlem"])
    if col_masraf == -1:
        col_masraf = find_col(["ücret"])
    if col_masraf == -1:
        for k in (["eft"], ["gönderim"], ["havale"], ["swift"], ["gelen eft"], ["gelen"], ["kanal"]):
            idx = find_col(k)
            if idx != -1:
                col_masraf = idx
                break

    col_asg_tutar = find_col(["asgari", "tutar"]) or find_col(["asgari"]) or find_col(["tutar"])
    col_asg_oran = find_col(["asgari", "oran"]) or find_col(["asgari oran"]) or find_col(["oran"])
    col_azm_tutar = find_col(["azami", "tutar"]) or find_col(["azami"]) or find_col(["tutar"])
    col_azm_oran = find_col(["azami", "oran"]) or find_col(["azami oran"]) or find_col(["oran"])
    col_aciklama = find_col(["açıklama"]) if find_col(["açıklama"]) != -1 else find_col(["aciklama"])
    col_tarih = find_col(["güncelleme"]) or find_col(["guncelleme"]) or find_col(["tarih"])

    if col_masraf == -1:
        col_masraf = 0

    return {
        "masraf": col_masraf,
        "asgari_tutar": col_asg_tutar,
        "asgari_oran": col_asg_oran,
        "azami_tutar": col_azm_tutar,
        "azami_oran": col_azm_oran,
        "aciklama": col_aciklama,
        "tarih": col_tarih,
    }


def scrape_yapikredi_all(url: str = YAPIKREDI_URL) -> List[UcretSatiri]:
    from playwright.sync_api import sync_playwright

    print(f"[yapikredi] {url} adresinden veri çekiliyor...", file=sys.stderr)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        try:
            context = browser.new_context(
                user_agent=HEADERS["User-Agent"],
                viewport={"width": 1440, "height": 1080},
                locale="tr-TR",
                extra_http_headers={
                    "Accept-Language": "tr-TR,tr;q=0.9,en;q=0.8",
                    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                }
            )
            page = context.new_page()
            page.goto(url, timeout=120000, wait_until="domcontentloaded")
            page.wait_for_timeout(2000)

            # Çerez onay butonlarını kapat
            for sel_text in ["Tümünü Kabul Et", "Tümünü Kabul", "Tümünü Reddet", "Kabul Et", "Kabul", "Kapat"]:
                try:
                    btn = page.locator(f"button:has-text(\"{sel_text}\"), a:has-text(\"{sel_text}\")").first
                    if btn.is_visible():
                        btn.click(timeout=1500)
                        page.wait_for_timeout(300)
                        break
                except Exception:
                    pass

            # Accordion, Tab ve gizli bölümleri aç
            try:
                accordion_selectors = [
                    "button[aria-expanded='false']",
                    "[role='button'][aria-expanded='false']",
                    ".accordion-header",
                    ".accordion-title",
                    ".collapsible-header",
                    "[data-toggle='collapse']"
                ]
                for selector in accordion_selectors:
                    elements = page.query_selector_all(selector)
                    for el in elements:
                        try:
                            if el.is_visible():
                                el.click()
                                page.wait_for_timeout(150)
                        except Exception:
                            pass
            except Exception:
                pass

            # Sayfayı kademeli aşağı kaydır (Lazy load/Render tetiklemesi)
            for _ in range(8):
                page.evaluate("window.scrollBy(0, window.innerHeight);")
                page.wait_for_timeout(300)

            try:
                page.wait_for_load_state("networkidle", timeout=5000)
            except Exception:
                pass

            # JS: Tabloları, üst başlıklarını ve satırlarını topla
            js_extract = r"""
() => {
    function findCategory(table) {
        let curr = table;
        // Yukarıya doğru hiyerarşik en yakın başlığı ara
        while (curr && curr !== document.body) {
            let prev = curr.previousElementSibling;
            while (prev) {
                if (['H1', 'H2', 'H3', 'H4', 'H5', 'H6'].includes(prev.tagName)) {
                    let txt = prev.innerText.trim();
                    if (txt) return txt;
                }
                let headingInside = prev.querySelector('h1, h2, h3, h4, h5, h6, .accordion-title, .title');
                if (headingInside) {
                    let txt = headingInside.innerText.trim();
                    if (txt) return txt;
                }
                prev = prev.previousElementSibling;
            }
            curr = curr.parentElement;
        }
        return "Yapı Kredi Masraflar";
    }

    function getHeaderText(table) {
        let thead = table.querySelector('thead');
        let row = null;
        if (thead) {
            const trs = thead.querySelectorAll('tr');
            row = trs[trs.length - 1] || trs[0];
        } else {
            row = table.querySelector('tr');
        }
        if (!row) return '';
        return Array.from(row.querySelectorAll('th, td')).map(c => c.innerText.trim()).join(' | ');
    }

    function rowsFromTable(table) {
        const trs = Array.from(table.querySelectorAll('tr'));
        const rows = [];
        for (const tr of trs) {
            const cells = Array.from(tr.querySelectorAll('th, td'));
            if (cells.length === 0) continue;
            rows.push(cells.map(c => c.innerText.trim()));
        }
        return rows;
    }

    const out = [];
    const tables = Array.from(document.querySelectorAll('table'));
    for (let i = 0; i < tables.length; i++) {
        const t = tables[i];
        const header = getHeaderText(t);
        const rows = rowsFromTable(t);
        const kategori = findCategory(t);
        out.push({ index: i, header: header, rows: rows, kategori: kategori });
    }
    return out;
}
"""
            tables_data = page.evaluate(js_extract)
            print(f"[yapikredi] Toplam {len(tables_data)} adet <table> bulundu.", file=sys.stderr)

        finally:
            browser.close()

    if not tables_data:
        raise ScraperError("Sayfada hiç tablo bulunamadı veya JS sonuç döndürmedi.")

    tum_satirlar: List[UcretSatiri] = []
    seen: Set[Tuple[str, str, str, str]] = set()

    for tinfo in tables_data:
        rows = tinfo.get("rows", []) or []
        kategori_baslik = _normalize(tinfo.get("kategori", "")) or "Genel"
        if not rows:
            continue

        header_idx = 0
        first_row = [c.lower() for c in rows[0]] if rows else []
        if not any(any(k in c for k in ["tutar", "asgari", "açıklama", "güncelle", "oran", "ücret", "masraf"]) for c in first_row):
            found = False
            for i in range(1, min(4, len(rows))):
                rowi = " ".join([c.lower() for c in rows[i]])
                if any(k in rowi for k in ["tutar", "asgari", "açıklama", "güncelle", "oran", "ücret", "masraf"]):
                    header_idx = i
                    found = True
                    break
            if not found:
                header_idx = 0

        header_row = rows[header_idx] if rows else []
        col_map = _find_col_indices_from_headers(header_row)

        for r in rows[header_idx + 1:]:
            def get_cell(idx):
                return _normalize(r[idx]) if 0 <= idx < len(r) else ""

            masraf = get_cell(col_map["masraf"])
            
            # Eğer masraf sütunu boşsa ya da sayı/tutar geldiyse alternatif metin hücresi bul
            if not masraf or (len(masraf) < 10 and any(tok in masraf.lower() for tok in ["tl", "%"])):
                for c in r:
                    c_norm = _normalize(c)
                    if c_norm and not any(tok in c_norm.lower() for tok in ["tl", "tutar", "asgari", "azami", "oran", "%"]):
                        masraf = c_norm
                        break

            if not masraf:
                continue

            asgari_tutar = get_cell(col_map["asgari_tutar"])
            asgari_oran = get_cell(col_map["asgari_oran"])
            azami_tutar = get_cell(col_map["azami_tutar"])
            azami_oran = get_cell(col_map["azami_oran"])
            aciklama_raw = get_cell(col_map["aciklama"])
            temiz_aciklama, aciklama_tarihi = _parse_aciklama(aciklama_raw)
            site_tarihi = get_cell(col_map["tarih"]) or aciklama_tarihi

            key = (kategori_baslik, masraf, asgari_tutar, azami_tutar)
            if key in seen:
                continue
            seen.add(key)

            tum_satirlar.append(UcretSatiri(
                kategori=kategori_baslik,
                masraf=masraf,
                asgari_tutar=asgari_tutar,
                asgari_oran=asgari_oran,
                azami_tutar=azami_tutar,
                azami_oran=azami_oran,
                aciklama=temiz_aciklama,
                site_guncelleme_tarihi=site_tarihi
            ))

    print(f"[yapikredi] Toplam {len(tum_satirlar)} satır komisyon/masraf verisi başarıyla işlendi.", file=sys.stderr)

    # Excel'e yaz
    try:
        import pandas as pd
        df = pd.DataFrame([s.__dict__ for s in tum_satirlar])
        out_fname = "yapikredi_all_komisyonlar.xlsx"
        df.to_excel(out_fname, index=False)
        print(f"[yapikredi] Excel dosyası oluşturuldu: {out_fname}", file=sys.stderr)
    except Exception as e:
        print(f"[yapikredi] Excel kaydedilirken hata oluştu: {e}", file=sys.stderr)

    return tum_satirlar


if __name__ == "__main__":
    sonuc = scrape_yapikredi_all()
    print(f"\nİşlem tamamlandı! Toplam çekilen satır sayısı: {len(sonuc)}")
