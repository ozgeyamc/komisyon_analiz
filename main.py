"""
Banka komisyon ücretleri takip botu - ana çalıştırma script'i.

Kullanım:
    python main.py
"""

import sys

from scraper import scrape_garanti_bbva
from update_excel import EXCEL_DOSYA_ADI, excel_guncelle_coklu


def _try_scrape(banka_adi, fn):
    try:
        satirlar = fn()
        print(f"{banka_adi}: {len(satirlar)} satır bulundu.")
        return satirlar
    except Exception as exc:
        print(f"[HATA] {banka_adi} çekilemedi: {exc}", file=sys.stderr)
        return None


def main() -> int:
    print("=== Banka Komisyon Ücretleri Takip Botu ===")

    banka_verileri = {}

    scrapers = [
        ("GARANTİ", scrape_garanti_bbva),
    ]

    optional = [
        ("ZİRAAT",    "scraper_ziraat",    "scrape_ziraat"),
        ("HALKBANK",  "scraper_halkbank",  "scrape_halkbank"),
        ("AKBANK",    "scraper_akbank",    "scrape_akbank"),
        ("YAPIKREDI", "scraper_yapikredi", "scrape_yapikredi"),
        ("VAKIFBANK", "scraper_vakifbank", "scrape_vakifbank"),
        ("QNB",       "scraper_qnb",       "scrape_qnb"),
        ("DENİZBANK", "scraper_denizbank", "scrape_denizbank"),
        ("TEB",       "scraper_teb",       "scrape_teb"),
        ("İŞBANKASI", "scraper_isbank",    "scrape_isbank"),
    ]

    for label, module, func in optional:
        try:
            mod = __import__(module)
            scrapers.append((label, getattr(mod, func)))
        except ImportError:
            pass

    for banka_adi, fn in scrapers:
        print(f"\n--- {banka_adi} çekiliyor ---")
        satirlar = _try_scrape(banka_adi, fn)
        if satirlar:
            banka_verileri[banka_adi] = satirlar

    if not banka_verileri:
        print("[HATA] Hiçbir bankadan veri çekilemedi!", file=sys.stderr)
        return 1

    try:
        ozet = excel_guncelle_coklu(banka_verileri, EXCEL_DOSYA_ADI)
    except Exception as exc:
        print(f"[HATA] Excel yazılırken hata: {exc}", file=sys.stderr)
        return 1

    print(f"\nTamamlandı. Toplam {ozet['eklendi']} satır yazıldı.")
    print(f"Excel dosyası: {EXCEL_DOSYA_ADI}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
