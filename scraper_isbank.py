"""
İş Bankası "Ürün ve Hizmet Ücretleri" sayfasını çeken scraper modülü.
"""

import sys
from dataclasses import dataclass
from typing import List

ISBANK_URL = "https://www.isbank.com.tr/urun-ve-hizmet-ucretleri"

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


def _meta(el, cls="UHU_icerik_meta"):
    span = el.find("span", class_=cls) if el else None
    return _normalize(span.get_text()) if span else ""


def scrape_isbank(url: str = ISBANK_URL) -> List[UcretSatiri]:
    from playwright.sync_api import sync_playwright
    from bs4 import BeautifulSoup

    print(f"[isbank] {url} adresinden veri çekiliyor...", file=sys.stderr)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        try:
            context = browser.new_context(
                user_agent=HEADERS["User-Agent"],
                viewport={"width": 1280, "height": 800},
                locale="tr-TR",
                timezone_id="Europe/Istanbul",
            )
            context.add_init_script(
                "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
            )
            page = context.new_page()
            page.goto(url, timeout=90000, wait_until="domcontentloaded")
            page.wait_for_timeout(15000)
            html = page.content()
        finally:
            browser.close()

    soup = BeautifulSoup(html, "lxml")
    tum_satirlar = []

    for hi in range(1, 20):
        grup_div = soup.find(id=f"h{hi}")
        if not grup_div:
            break

        ana_kategori_el = grup_div.find(class_="UHU_group_header")
        ana_kategori = _normalize(ana_kategori_el.get_text()) if ana_kategori_el else f"Grup {hi}"

        item_headers = grup_div.find_all(class_="UHU_item_header")
        for item_el in item_headers:
            alt_kategori = _normalize(item_el.get_text())

            item_sub_cover = item_el.find_next_sibling(id="UHU_itemSubCover")
            if not item_sub_cover:
                item_sub_cover = item_el.parent

            sub_headers = item_sub_cover.find_all(class_="UHU_itemSub_header") if item_sub_cover else []

            for sub_el in (sub_headers if sub_headers else [None]):
                if sub_el is not None:
                    sub_kategori = _normalize(sub_el.get_text())
                    tam_kategori = f"{ana_kategori} - {alt_kategori} - {sub_kategori}"
                    icerik_gc = sub_el.find_next_sibling(id="UHU_item_icerik_GC")
                    if not icerik_gc:
                        icerik_gc = sub_el.parent
                else:
                    tam_kategori = f"{ana_kategori} - {alt_kategori}"
                    icerik_gc = item_sub_cover

                if not icerik_gc:
                    continue

                for blok in icerik_gc.find_all(class_="UHU_item_icerikC"):
                    masraf_el = blok.find(class_="UHU_item_icerikH")
                    masraf = _normalize(masraf_el.get_text()) if masraf_el else ""
                    if not masraf:
                        continue

                    icerik1 = blok.find(class_="UHU_item_icerik1")
                    icerik2 = blok.find(class_="UHU_item_icerik2")
                    icerik3 = blok.find(class_="UHU_item_icerik3")
                    icerik4 = blok.find(class_="UHU_item_icerik4")
                    icerik5 = blok.find(class_="UHU_item_icerik5")
                    aciklama_el = blok.find(class_="UHU_item_icerikF")

                    asgari_tutar = _meta(icerik1)
                    asgari_oran  = _meta(icerik2)
                    azami_tutar  = _meta(icerik3)
                    azami_oran   = _meta(icerik4)
                    tarih        = _meta(icerik5, cls="UHU_icerik_meta2")
                    aciklama     = _normalize(aciklama_el.get_text()) if aciklama_el else ""

                    tum_satirlar.append(UcretSatiri(
                        kategori=tam_kategori,
                        masraf=masraf,
                        asgari_tutar=asgari_tutar,
                        asgari_oran=asgari_oran,
                        azami_tutar=azami_tutar,
                        azami_oran=azami_oran,
                        aciklama=aciklama,
                        site_guncelleme_tarihi=tarih,
                    ))

    if not tum_satirlar:
        raise ScraperError("İş Bankası sayfasında hiç veri satırı çekilemedi.")

    tum_satirlar = sorted(tum_satirlar, key=lambda s: (s.kategori, s.masraf))
    print(f"[isbank] Toplam {len(tum_satirlar)} satır bulundu.", file=sys.stderr)
    return tum_satirlar


if __name__ == "__main__":
    veriler = scrape_isbank()
    for v in veriler[:10]:
        print(v)
    print(f"\nToplam {len(veriler)} satır bulundu.")
