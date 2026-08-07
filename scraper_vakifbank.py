"""
Vakıfbank "Ürün ve Hizmet Ücretleri" - Playwright route intercept ile API yanıtını yakalar.
"""

import re
import sys
import json
import time
from dataclasses import dataclass
from typing import List

VAKIFBANK_API_URL = "https://inbound.apigateway.vakifbank.com.tr:8443/getProductServicePrices"
VAKIFBANK_PAGE_URL = "https://www.vakifbank.com.tr/tr/urun-ve-hizmet-ucretleri"

DATE_PATTERN = re.compile(
    r"G[üu]ncellenme\s*Tarihi\s*:\s*[\s\xa0]*(\d{2}[./]\d{2}[./]\d{4}(?:[\s\xa0]+\d{2}:\d{2})?)",
    re.IGNORECASE,
)

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


def _normalize(val) -> str:
    if val is None:
        return ""
    return str(val).strip().replace("\xa0", " ").replace("\u200b", "").strip()


def _normalize_tutar(val) -> str:
    """Sayısal veya string 0 değerlerini boş döner, diğerlerini string yapar."""
    if val is None:
        return ""
    if isinstance(val, (int, float)):
        if val == 0:
            return ""
        if isinstance(val, float) and val == int(val):
            return str(int(val))
        return str(val)
    v = str(val).strip().replace("\xa0", " ").replace("\u200b", "").strip()
    if not v or v in ("0", "0.0", "0.00", "-"):
        return ""
    try:
        if float(v.replace(",", ".")) == 0:
            return ""
    except (ValueError, TypeError):
        pass
    return v


def _normalize_tarih(val) -> str:
    """
    Vakıfbank tarih formatlarını dd.mm.yyyy HH:MM'e çevirir.
    Desteklenen formatlar:
      "2026-04-28T10:10:41"   → "28.04.2026 10:10"
      "2026-04-28T10:10:41Z"  → "28.04.2026 10:10"
      "24.04.202609:27:39"    → "24.04.2026 09:27"
      "24.04.2026 09:27:39"   → "24.04.2026 09:27"
      "24.04.2026 09:27"      → "24.04.2026 09:27"
      "24.04.2026"            → "24.04.2026"
    """
    v = _normalize(val)
    if not v:
        return ""

    # ISO format: "2026-04-28T10:10:41" veya "2026-04-28T10:10:41Z"
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})", v)
    if m:
        return f"{m.group(3)}.{m.group(2)}.{m.group(1)} {m.group(4)}:{m.group(5)}"

    # "dd.mm.yyyyHH:MM" — yıl ile saat birleşik (boşluksuz)
    m = re.match(r"(\d{2}[./]\d{2}[./]\d{4})(\d{2}:\d{2})", v)
    if m:
        return f"{m.group(1)} {m.group(2)}"

    # "dd.mm.yyyy HH:MM:SS" — saniyeyi at
    m = re.match(r"(\d{2}[./]\d{2}[./]\d{4})\s+(\d{2}:\d{2}):\d{2}", v)
    if m:
        return f"{m.group(1)} {m.group(2)}"

    # "dd.mm.yyyy HH:MM" — zaten temiz
    m = re.match(r"(\d{2}[./]\d{2}[./]\d{4})\s+(\d{2}:\d{2})", v)
    if m:
        return f"{m.group(1)} {m.group(2)}"

    # "dd.mm.yyyy" — sadece tarih
    m = re.match(r"(\d{2}[./]\d{2}[./]\d{4})$", v)
    if m:
        return m.group(1)

    return v


def _parse_vakifbank_fee_list(fee_list: list) -> List[UcretSatiri]:
    satirlar = []
    for item in fee_list:
        if not isinstance(item, dict):
            continue
        masraf = _normalize(item.get("FeeName", ""))
        if not masraf:
            masraf = _normalize(item.get("ItemName", ""))
        if not masraf:
            continue
        satirlar.append(UcretSatiri(
            kategori=_normalize(item.get("MainTransactionGroupName", "Genel")),
            masraf=masraf,
            asgari_tutar=_normalize_tutar(item.get("MinimumAmount", "")),
            asgari_oran=_normalize_tutar(item.get("MinimumRate", "")),
            azami_tutar=_normalize_tutar(item.get("MaximumAmount", "")),
            azami_oran=_normalize_tutar(item.get("MaximumRate", "")),
            aciklama=_normalize(item.get("Description", "")),
            site_guncelleme_tarihi=_normalize_tarih(item.get("UpdateDate", "")),
        ))
    return satirlar


def scrape_vakifbank(page_url: str = VAKIFBANK_PAGE_URL) -> List[UcretSatiri]:
    from playwright.sync_api import sync_playwright

    print(f"[vakifbank] Playwright intercept başlıyor...", file=sys.stderr)
    captured = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        try:
            context = browser.new_context(
                user_agent=HEADERS["User-Agent"],
                viewport={"width": 1280, "height": 800},
                locale="tr-TR",
                extra_http_headers={
                    "Accept-Language": "tr-TR,tr;q=0.9,en;q=0.8",
                }
            )
            page = context.new_page()

            def handle_response(response):
                url = response.url
                if "getProductServicePrices" in url or "ProductServicePrice" in url:
                    print(f"[vakifbank] Hedef URL: {url} ({response.status})", file=sys.stderr)
                    if response.status == 200:
                        try:
                            body = response.body()
                            text = body.decode("utf-8")
                            if len(text) > 50:
                                captured.append(text)
                                print(f"[vakifbank] Yanıt yakalandı! ({len(text)} karakter)", file=sys.stderr)
                        except Exception as e:
                            print(f"[vakifbank] body() hatası: {e}", file=sys.stderr)

            page.on("response", handle_response)

            print(f"[vakifbank] Sayfa yükleniyor...", file=sys.stderr)
            page.goto(page_url, timeout=120000, wait_until="domcontentloaded")

            deadline = time.time() + 30
            while not captured and time.time() < deadline:
                page.wait_for_timeout(1000)
                print(f"[vakifbank] Bekleniyor... ({int(deadline - time.time())}sn kaldı)", file=sys.stderr)

            if captured:
                print(f"[vakifbank] Yanıt alındı!", file=sys.stderr)
            else:
                print(f"[vakifbank] 30sn içinde yanıt gelmedi.", file=sys.stderr)

        finally:
            browser.close()

    if not captured:
        raise ScraperError("Vakıfbank API yanıtı yakalanamadı.")

    full_response = captured[0]

    try:
        data = json.loads(full_response)
    except json.JSONDecodeError:
        print(f"[vakifbank] JSON parse hatası:\n{full_response[:2000]}", file=sys.stderr)
        raise ScraperError("Vakıfbank API yanıtı JSON değil.")

    print(f"[vakifbank] JSON parse edildi. Tip: {type(data).__name__}", file=sys.stderr)

    fee_list = None
    if isinstance(data, dict):
        data_block = data.get("Data") or data.get("data")
        if isinstance(data_block, dict):
            fee_list = data_block.get("Fee") or data_block.get("fee")
        elif isinstance(data_block, list):
            fee_list = data_block
    elif isinstance(data, list):
        fee_list = data

    if not fee_list:
        print(f"[vakifbank] Ham JSON (ilk 3000):\n{full_response[:3000]}", file=sys.stderr)
        raise ScraperError("Vakıfbank API'sinde Fee listesi bulunamadı.")

    print(f"[vakifbank] Fee listesi uzunluğu: {len(fee_list)}", file=sys.stderr)
    tum_satirlar = _parse_vakifbank_fee_list(fee_list)

    if not tum_satirlar:
        raise ScraperError("Vakıfbank Fee listesi parse edilemedi.")

    print(f"[vakifbank] Toplam {len(tum_satirlar)} satır bulundu.", file=sys.stderr)
    return tum_satirlar


if __name__ == "__main__":
    veriler = scrape_vakifbank()
    for v in veriler[:5]:
        print(v)
    print(f"Toplam {len(veriler)} satır bulundu.")
