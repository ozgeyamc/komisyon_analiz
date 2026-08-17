"""
Yapı Kredi "Bireysel Ürün ve Hizmet Ücretleri" sayfasını çeken scraper modülü.
JS-evaluation tabanlı tablo okuma kullanır (colspan/rowspan/thead farklılıklarına karşı daha dayanıklı).
"""

import re
import sys
from dataclasses import dataclass
from typing import List, Tuple, Dict, Any, Set

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
    raw_aciklama = (raw_aciklama or "").strip()

    match = DATE_PATTERN.search(raw_aciklama)
    if match:
        tarih = match.group(1)
        temiz_aciklama = DATE_PATTERN.sub("", raw_aciklama).strip(" .")
        return temiz_aciklama, tarih

    match_tr = DATE_PATTERN_TR.search(raw_aciklama)
    if match_tr:
        gun = match_tr.group(1).zfill(2)
        ay = TURKCE_AYLAR.get(match_tr.group(2).lower(), "00")
        yil = match_tr.group(3)
        tarih = f"{gun}.{ay}.{yil}"
        temiz_aciklama = DATE_PATTERN_TR.sub("", raw_aciklama).strip(" .")
        return temiz_aciklama, tarih

    return raw_aciklama, ""


def _normalize(val: str) -> str:
    return (val or "").strip().replace("\xa0", " ").replace("\u200b", "").strip()


def _find_category_title(el, fallback: str) -> str:
    # Not used for JS extraction, kept for compatibility if needed
    parent = el.parent
    depth = 0
    while parent is not None and depth < 10:
        for sibling in parent.find_all_previous(["h1", "h2", "h3", "h4", "h5"], limit=3):
            text = _normalize(sibling.get_text())
            if len(text) > 5 and text not in ["Müşteri Ol", "Ara", "Kapat", "Menü", "Ana Sayfa"]:
                return text
        parent = parent.parent
        depth += 1
    return fallback


def _find_col_indices_from_headers(header_texts: List[str]) -> Dict[str, int]:
    # normalize headers
    def norm(h: str) -> str:
        h2 = re.sub(r"\(.*?\)", "", (h or ""))
        h2 = h2.replace("%", " ")
        h2 = re.sub(r"\s+", " ", h2).strip().lower()
        return h2
    headers_norm = [norm(h) for h in header_texts]

    def find_col(keywords):
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
        candidate = find_col(["eft"])
        if candidate != -1:
            col_masraf = candidate
        else:
            candidate2 = find_col(["gönderim"]) or find_col(["gönderimi"]) or find_col(["havale"]) or find_col(["swift"])
            if candidate2 != -1:
                col_masraf = candidate2

    col_asg_tutar = find_col(["asgari", "tutar"]) or find_col(["asgari"]) or find_col(["tutar"])
    col_asg_oran = find_col(["asgari", "oran"]) or find_col(["asgari oran"]) or find_col(["oran"])
    col_azm_tutar = find_col(["azami", "tutar"]) or find_col(["azami"]) or find_col(["tutar"])
    col_azm_oran = find_col(["azami", "oran"]) or find_col(["azami oran"]) or find_col(["oran"])
    col_aciklama = find_col(["açıklama"])
    if col_aciklama == -1:
        col_aciklama = find_col(["aciklama"])
    col_tarih = find_col(["güncelleme"])
    if col_tarih == -1:
        col_tarih = find_col(["guncelleme"]) or find_col(["tarih"])

    # fallback defaults
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


def scrape_yapikredi(url: str = YAPIKREDI_URL) -> List[UcretSatiri]:
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

            try:
                page.wait_for_selector("table", timeout=20000)
            except Exception:
                page.wait_for_timeout(8000)
            page.wait_for_timeout(1000)

            table_count = page.evaluate("() => document.querySelectorAll('table').length")
            print(f"[yapikredi] JS ile {table_count} <table> bulundu", file=sys.stderr)

            # Evaluate JS: find relevant tables and return header + rows (array of arrays)
            js = r"""
() => {
  const keywords = ['eft','gönderim','gönder','havale','swift','para gönderme','eft gönderimi','para gönder','eft gönderimi'];
  function isRelevantText(t){
    if(!t) return false;
    const s = t.toLowerCase();
    return keywords.some(k => s.includes(k));
  }
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
    // also check parent section text for keywords
    const parentText = (t.closest('section') || t.parentElement || {}).innerText || '';
    if(isRelevantText(header) || isRelevantText(parentText)){
      out.push({index: i, header: header, rows: rowsFromTable(t)});
    }
  }
  return out;
}
"""
            tables_data = page.evaluate(js)
        finally:
            browser.close()

    # tables_data is a list of dicts: {index, header, rows}
    print(f"[yapikredi][debug] Relevant tables found: {len(tables_data)}", file=sys.stderr)

    tum_satirlar: List[UcretSatiri] = []
    seen: Set[tuple] = set()

    for tinfo in tables_data:
        header_preview = tinfo.get("header", "")
        rows = tinfo.get("rows", [])
        if not rows:
            continue
        # determine header row: if first row contains 'asgari' or 'tutar' consider it header
        header_row = rows[0]
        header_norm = [h.lower() for h in header_row]
        # if first row seems to be data (no 'tutar' words), try to find a header row index
        header_idx = 0
        if not any("tutar" in h or "asgari" in h or "açıklama" in h or "güncelle" in h or "oran" in h for h in header_norm):
            # search first 3 rows for a header-like row
            found = False
            for i in range(1, min(4, len(rows))):
                r = rows[i]
                low = " ".join([c.lower() for c in r])
                if any(k in low for k in ["tutar", "asgari", "açıklama", "güncelle", "oran"]):
                    header_idx = i
                    header_row = r
                    found = True
                    break
            if not found:
                header_idx = 0
                header_row = rows[0]

        col_map = _find_col_indices_from_headers(header_row)

        # data rows are rows after header_idx
        for r in rows[header_idx+1:]:
            def get_cell(idx):
                return r[idx].strip() if 0 <= idx < len(r) else ""

            masraf = get_cell(col_map["masraf"])
            # If the masraf cell looks like an amount (e.g. contains 'TL' or digits) but not a description,
            # try to take the first non-empty textual cell instead
            if masraf and (masraf.lower().count("tl") > 0 and len(masraf) < 6):
                # try find a descriptive cell
                for c in r:
                    if c and not any(tok in c.lower() for tok in ["tl","%","tutar","asgari","azami"]):
                        masraf = c; break

            # fallback: if masraf empty, search row for any cell containing 'eft' or 'internet'
            if not masraf:
                for c in r:
                    if c and ("eft" in c.lower() or "internet" in c.lower() or "şube" in c.lower() or "mobil" in c.lower()):
                        masraf = c
                        break
            if not masraf:
                continue

            asgari_tutar = get_cell(col_map["asgari_tutar"])
            asgari_oran  = get_cell(col_map["asgari_oran"])
            azami_tutar  = get_cell(col_map["azami_tutar"])
            azami_oran   = get_cell(col_map["azami_oran"])
            aciklama_raw = get_cell(col_map["aciklama"])
            temiz_aciklama, aciklama_tarihi = _parse_aciklama(aciklama_raw)
            site_tarihi = get_cell(col_map["tarih"]) or aciklama_tarihi

            kategori = header_preview or "Yapikredi"
            key = (kategori, masraf, asgari_tutar or "", azami_tutar or "")
            if key in seen:
                continue
            seen.add(key)

            tum_satirlar.append(UcretSatiri(
                kategori=kategori, masraf=masraf,
                asgari_tutar=asgari_tutar, asgari_oran=asgari_oran,
                azami_tutar=azami_tutar, azami_oran=azami_oran,
                aciklama=temiz_aciklama, site_guncelleme_tarihi=site_tarihi
            ))

    print(f"[yapikredi] Toplam {len(tum_satirlar)} satır bulundu.", file=sys.stderr)
    if not tum_satirlar:
        raise ScraperError("Yapı Kredi sayfasında ilgili tablolar bulundu ama hiç veri satırı çekilemedi.")
    return tum_satirlar


if __name__ == "__main__":
    veriler = scrape_yapikredi()
    for v in veriler[:50]:
        print(v)
    print(f"Toplam {len(veriler)} satır bulundu.")
