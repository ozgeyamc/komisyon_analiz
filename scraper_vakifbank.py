"""
Vakıfbank "Ürün ve Hizmet Ücretleri" - direkt API endpoint'i kullanır.
"""

import re
import sys
import json
from dataclasses import dataclass
from typing import List

VAKIFBANK_API_URL = "https://inbound.apigateway.vakifbank.com.tr:8443/getProductServicePrices"
VAKIFBANK_PAGE_URL = "https://www.vakifbank.com.tr/tr/urun-ve-hizmet-ucretleri"

DATE_PATTERN = re.compile(
    r"G[üu]ncell[ei]nme\s*Tarihi\s*:\s*(\d{2}\.\d{2}\.\d{4}(?:\s+\d{2}:\d{2})?)",
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


def _parse_aciklama(raw_aciklama: str):
    raw_aciklama = raw_aciklama.strip()
    match = DATE_PATTERN.search(raw_aciklama)
    tarih = match.group(1) if match else ""
    temiz_aciklama = DATE_PATTERN.sub("", raw_aciklama).strip(" .")
    return temiz_aciklama, tarih


def _parse_json_response(data, satirlar: List[UcretSatiri]):
    if isinstance(data, list):
        for item in data:
            _parse_json_response(item, satirlar)
    elif isinstance(data, dict):
        keys = [k.lower() for k in data.keys()]
        has_fee = any(k in keys for k in [
            "masraf", "ucret", "fee", "price", "tutar", "amount",
            "servicename", "servisadi", "islemadi", "servisucretid",
            "islemturu", "hizmetadi", "hizmetucret"
        ])
        if has_fee:
            kategori = ""
            masraf = ""
            asg_tutar = ""
            asg_oran = ""
            azm_tutar = ""
            azm_oran = ""
            aciklama = ""
            tarih = ""

            for k, v in data.items():
                kl = k.lower()
                if any(x in kl for x in ["kategori", "category", "grup", "group", "baslik", "title"]):
                    kategori = _normalize(v)
                elif any(x in kl for x in ["masraf", "islemadi", "servicename", "servisadi", "hizmetadi", "islemturu"]):
                    if not masraf:
                        masraf = _normalize(v)
                elif kl in ["ad", "name"]:
                    if not masraf:
                        masraf = _normalize(v)
                elif any(x in kl for x in ["asgari", "min"]) and any(x in kl for x in ["tutar", "amount", "fiyat"]):
                    asg_tutar = _normalize(v)
                elif any(x in kl for x in ["asgari", "min"]) and any(x in kl for x in ["oran", "rate", "percent"]):
                    asg_oran = _normalize(v)
                elif any(x in kl for x in ["azami", "max"]) and any(x in kl for x in ["tutar", "amount", "fiyat"]):
                    azm_tutar = _normalize(v)
                elif any(x in kl for x in ["azami", "max"]) and any(x in kl for x in ["oran", "rate", "percent"]):
                    azm_oran = _normalize(v)
                elif any(x in kl for x in ["tutar", "amount", "fiyat", "ucret", "fee", "price"]):
                    if not asg_tutar:
                        asg_tutar = _normalize(v)
                elif any(x in kl for x in ["oran", "rate", "percent"]):
                    if not asg_oran:
                        asg_oran = _normalize(v)
                elif any(x in kl for x in ["aciklama", "description", "note", "bilgi"]):
                    aciklama = _normalize(v)
                elif any(x in kl for x in ["tarih", "date", "guncelleme"]):
                    tarih = _normalize(v)

            if masraf:
                temiz_aciklama, parsed_tarih = _parse_aciklama(aciklama)
                satirlar.append(UcretSatiri(
                    kategori=kategori or "Genel",
                    masraf=masraf,
                    asgari_tutar=asg_tutar,
                    asgari_oran=asg_oran,
                    azami_tutar=azm_tutar,
                    azami_oran=azm_oran,
                    aciklama=temiz_aciklama,
                    site_guncelleme_tarihi=tarih or parsed_tarih,
                ))
        else:
            for v in data.values():
                if isinstance(v, (dict, list)):
                    _parse_json_response(v, satirlar)


def _try_requests(api_url: str, cookies: dict = None) -> str:
    import requests
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    headers = {
        "User-Agent": HEADERS["User-Agent"],
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "tr-TR,tr;q=0.9,en;q=0.8",
        "Referer": VAKIFBANK_PAGE_URL,
        "Origin": "https://www.vakifbank.com.tr",
    }

    resp = requests.get(
        api_url,
        headers=headers,
        cookies=cookies or {},
        timeout=30,
        verify=False,
    )
    print(f"[vakifbank] requests status: {resp.status_code}", file=sys.stderr)
    print(f"[vakifbank] requests yanıt (ilk 500): {resp.text[:500]}", file=sys.stderr)
    resp.raise_for_status()
    return resp.text


def _get_cookies_via_playwright() -> dict:
    from playwright.sync_api import sync_playwright

    print("[vakifbank] Cookie'ler toplanıyor...", file=sys.stderr)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        try:
            context = browser.new_context(
                user_agent=HEADERS["User-Agent"],
                locale="tr-TR",
            )
            page = context.new_page()
            page.goto("https://www.vakifbank.com.tr", timeout=60000, wait_until="domcontentloaded")
            page.wait_for_timeout(3000)
            cookies = context.cookies()
            return {c["name"]: c["value"] for c in cookies}
        finally:
            browser.close()


def _intercept_api_via_playwright(api_url: str) -> str:
    from playwright.sync_api import sync_playwright

    print("[vakifbank] Playwright route intercept deneniyor...", file=sys.stderr)
    captured = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        try:
            context = browser.new_context(
                user_agent=HEADERS["User-Agent"],
                locale="tr-TR",
                extra_http_headers={
                    "Accept-Language": "tr-TR,tr;q=0.9,en;q=0.8",
                }
            )

            def handle_response(response):
                if "getProductServicePrices" in response.url:
                    try:
                        body = response.body()
                        captured.append(body.decode("utf-8"))
                        print(f"[vakifbank] API yanıtı yakalandı! ({len(body)} byte)", file=sys.stderr)
                    except Exception as e:
                        print(f"[vakifbank] Yanıt okunamadı: {e}", file=sys.stderr)

            page = context.new_page()
            page.on("response", handle_response)
            page.goto(VAKIFBANK_PAGE_URL, timeout=120000, wait_until="networkidle")
            page.wait_for_timeout(5000)
        finally:
            browser.close()

    if captured:
        return captured[0]
    return ""


def scrape_vakifbank(api_url: str = VAKIFBANK_API_URL) -> List[UcretSatiri]:
    print(f"[vakifbank] Başlıyor...", file=sys.stderr)

    full_response = ""

    # Yöntem 1: Cookie alıp direkt requests ile çağır
    try:
        cookies = _get_cookies_via_playwright()
        print(f"[vakifbank] {len(cookies)} cookie alındı", file=sys.stderr)
        full_response = _try_requests(api_url, cookies)
    except Exception as e:
        print(f"[vakifbank] requests yöntemi başarısız: {e}", file=sys.stderr)

    # Yöntem 2: Playwright route intercept
    if not full_response:
        try:
            full_response = _intercept_api_via_playwright(api_url)
        except Exception as e:
            print(f"[vakifbank] Playwright intercept başarısız: {e}", file=sys.stderr)

    if not full_response:
        raise ScraperError("Vakıfbank API'sine erişilemedi.")

    try:
        data = json.loads(full_response)
    except json.JSONDecodeError:
        print(f"[vakifbank] JSON parse hatası, ham yanıt:\n{full_response[:2000]}", file=sys.stderr)
        raise ScraperError("Vakıfbank API yanıtı JSON değil.")

    print(f"[vakifbank] JSON parse edildi. Tip: {type(data).__name__}", file=sys.stderr)
    if isinstance(data, dict):
        print(f"[vakifbank] JSON keys: {list(data.keys())[:20]}", file=sys.stderr)
    elif isinstance(data, list):
        print(f"[vakifbank] JSON liste uzunluğu: {len(data)}", file=sys.stderr)
        if data:
            print(f"[vakifbank] İlk eleman: {str(data[0])[:300]}", file=sys.stderr)

    tum_satirlar = []
    _parse_json_response(data, tum_satirlar)

    if not tum_satirlar:
        print(f"[vakifbank] Ham JSON (ilk 3000 karakter):\n{full_response[:3000]}", file=sys.stderr)
        raise ScraperError("Vakıfbank API'sinden veri parse edilemedi.")

    print(f"[vakifbank] Toplam {len(tum_satirlar)} satır bulundu.", file=sys.stderr)
    return tum_satirlar


if __name__ == "__main__":
    veriler = scrape_vakifbank()
    for v in veriler[:5]:
        print(v)
    print(f"Toplam {len(veriler)} satır bulundu.")
