"""
scraper_yapikredi_all.py

Yapı Kredi "Bireysel Ürün ve Hizmet Ücretleri" sayfasından tüm tabloları çeker.
- Playwright ile sayfayı açar, cookie/dialog butonlarını denemeye tıklar,
  accordions'u açar, sayfayı kaydırır, sonra JS ile tüm tabloları toplar.
- Python tarafında header eşlemesi yapıp UcretSatiri nesneleri üretir.
- Yinelenenleri temizler ve sonucu Excel'e yazar (pandas yüklüyse).
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
    """
    header_texts: list of header cell texts (original)
    returns mapping of keys to column indices (or -1 if not found)
    """
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
        # try single words that often are used as section titles
        for k in (["eft"], ["gönderim"], ["havale"], ["swift"], ["gelen eft"], ["gelen"]):
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
                viewport={"width": 1280, "height": 1000},
                locale="tr-TR",
                extra_http_headers={
                    "Accept-Language": "tr-TR,tr;q=0.9,en;q=0.8",
                    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
                    "Accept-Encoding": "gzip, deflate, br",
                }
            )
            page = context.new_page()
            page.goto(url, timeout=120000, wait_until="domcontentloaded")

            # Try click common cookie buttons (Turkish variants)
            for sel_text in ["Tümünü Kabul Et", "Tümünü Kabul", "Tümünü Reddet", "Kabul Et", "Kabul"]:
                try:
                    page.locator(f"button:has-text(\"{sel_text}\")").first.click(timeout=1500)
                    print(f"[yapikredi][debug] Clicked cookie/button: {sel_text}", file=sys.stderr)
                    page.wait_for_timeout(300)
                    break
                except Exception:
                    pass

            # Try expand collapsed accordions and buttons that likely reveal sections
            try:
                # click any collapsed buttons
                toggles = page.query_selector_all("button[aria-expanded='false'], [role='button'][aria-expanded='false']")
                for t in toggles:
                    try:
                        t.click()
                        page.wait_for_timeout(120)
                    except Exception:
                        pass
                # click buttons containing keywords (eft/havale/gönderim)
                btns = page.query_selector_all("button")
                for b in btns:
                    try:
                        txt = (b.inner_text() or "").lower()
                        if any(k in txt for k in ["eft", "havale", "gönderim", "gönder"]):
                            try:
                                b.click()
                                page.wait_for_timeout(120)
                            except Exception:
                                pass
                    except Exception:
                        pass
            except Exception:
                pass

            # Scroll slowly to bottom to trigger lazy loading / rendering
            for _ in range(6):
                page.evaluate("window.scrollBy(0, window.innerHeight);")
                page.wait_for_timeout(400)

            # Wait a little for JS rendering
            try:
                page.wait_for_load_state("networkidle", timeout=8000)
            except Exception:
                pass
            page.wait_for_timeout(800)

            # JS: collect all tables -> header + rows arrays
            js = r"""
() => {
  function getHeaderText(table){
    let thead = table.querySelector('thead');
    let row = null;
    if(thead){
      const trs = thead.querySelectorAll('tr');
      row = trs[trs.length-1] || trs[0];
    } else {
      row = table.querySelector('tr');
    }
    if(!row) return '';
    return Array.from(row.querySelectorAll('th,td')).map(c=>c.innerText.trim()).join(' | ');
  }
  function rowsFromTable(table){
    const trs = Array.from(table.querySelectorAll('tr'));
    const rows = [];
    for(const tr of trs){
      const cells = Array.from(tr.querySelectorAll('th,td'));
      if(cells.length === 0) continue;
      rows.push(cells.map(c => c.innerText.trim()));
    }
    return rows;
  }
  const out = [];
  const tables = Array.from(document.querySelectorAll('table'));
  for(let i=0;i<tables.length;i++){
    const t = tables[i];
    const header = getHeaderText(t);
    const rows = rowsFromTable(t);
    out.push({index:i, header: header, rows: rows});
  }
  return out;
}
"""
            tables_data = page.evaluate(js)
            table_count = len(tables_data)
            print(f"[yapikredi] JS ile {table_count} <table> bulundu", file=sys.stderr)

        finally:
            browser.close()

    if not tables_data:
        raise ScraperError("Sayfada hiç tablo bulunamadı veya JS sonuç döndürmedi.")

    tum_satirlar: List[UcretSatiri] = []
    seen: Set[Tuple[str, str, str, str]] = set()
    processed_tables = 0

    for tinfo in tables_data:
        processed_tables += 1
        header_preview = tinfo.get("header", "") or ""
        rows = tinfo.get("rows", []) or []
        if not rows:
            continue

        # find header row index: prefer a row containing 'tutar'/'asgari'/'açıklama'/'güncelle'
        header_idx = 0
        first_row = [c.lower() for c in rows[0]] if rows else []
        if not any(any(k in c for k in ["tutar", "asgari", "açıklama", "güncelle", "oran"]) for c in first_row):
            found = False
            for i in range(1, min(4, len(rows))):
                rowi = " ".join([c.lower() for c in rows[i]])
                if any(k in rowi for k in ["tutar", "asgari", "açıklama", "güncelle", "oran"]):
                    header_idx = i
                    found = True
                    break
            if not found:
                header_idx = 0

        header_row = rows[header_idx] if rows else []
        col_map = _find_col_indices_from_headers(header_row)

        # Data rows are after header_idx
        for r in rows[header_idx+1:]:
            def get_cell(idx):
                return r[idx].strip() if 0 <= idx < len(r) else ""

            masraf = get_cell(col_map["masraf"])
            # If masraf seems like an amount (e.g., contains TL only) try to pick a textual cell
            if masraf:
                low = masraf.lower()
                if (("tl" in low or "%" in low or re.search(r"\d", low)) and len(masraf) < 10):
                    # try find a better descriptive cell
                    for c in r:
                        if c and not any(tok in c.lower() for tok in ["tl", "tutar", "asgari", "azami", "oran", "%"]):
                            masraf = c
                            break

            # fallback: pick first non-empty textual cell that is not an amount
            if not masraf:
                for c in r:
                    if c and not any(tok in c.lower() for tok in ["tl","tutar","asgari","azami","oran","%"]):
                        masraf = c
                        break

            if not masraf:
                # give up on this row
                continue

            asgari_tutar = get_cell(col_map["asgari_tutar"])
            asgari_oran = get_cell(col_map["asgari_oran"])
            azami_tutar = get_cell(col_map["azami_tutar"])
            azami_oran = get_cell(col_map["azami_oran"])
            aciklama_raw = get_cell(col_map["aciklama"])
            temiz_aciklama, aciklama_tarihi = _parse_aciklama(aciklama_raw)
            site_tarihi = get_cell(col_map["tarih"]) or aciklama_tarihi

            kategori = header_preview or "Yapikredi"

            key = (kategori, masraf, asgari_tutar or "", azami_tutar or "")
            if key in seen:
                continue
            seen.add(key)

            tum_satirlar.append(UcretSatiri(
                kategori=kategori,
                masraf=masraf,
                asgari_tutar=asgari_tutar,
                asgari_oran=asgari_oran,
                azami_tutar=azami_tutar,
                azami_oran=azami_oran,
                aciklama=temiz_aciklama,
                site_guncelleme_tarihi=site_tarihi
            ))

    print(f"[yapikredi] İşlenen tablo sayısı: {processed_tables}", file=sys.stderr)
    print(f"[yapikredi] Toplam {len(tum_satirlar)} satır bulundu.", file=sys.stderr)

    # Try write to Excel if pandas is available
    try:
        import pandas as pd
        df = pd.DataFrame([s.__dict__ for s in tum_satirlar])
        out_fname = "yapikredi_all_komisyonlar.xlsx"
        df.to_excel(out_fname, index=False)
        print(f"[yapikredi] Sonuç {out_fname} olarak kaydedildi.", file=sys.stderr)
    except Exception:
        pass

    return tum_satirlar


if __name__ == "__main__":
    sonuc = scrape_yapikredi_all()
    # print some examples
    for s in sonuc[:200]:
        print(s)
    print(f"Toplam {len(sonuc)} satır bulundu.")
