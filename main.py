"""
Garanti BBVA + Ziraat + Halkbank + Akbank + Yapı Kredi komisyon ücretleri takip botu.
"""

import sys

from scraper import scrape_garanti_bbva
from scraper_ziraat import scrape_ziraat
from scraper_halkbank import scrape_halkbank
from scraper_akbank import scrape_akbank
from scraper_yapikredi import scrape_yapikredi
from update_excel import EXCEL_DOSYA_ADI, excel_guncelle_coklu


def main() -> int:
    print("=== Banka Komisyon Ücretleri Takip Botu ===")

    banka_verileri = {}

    print("\n--- Garanti BBVA çekiliyor ---")
    try:
        satirlar = scrape_garanti_bbva()
        banka_verileri["GARANTİ"] = satirlar
        print(f"Garanti: {len(satirlar)} satır bulundu.")
    except Exception as exc:
        print(f"[HATA] Garanti çekilemedi: {exc}", file=sys.stderr)

    print("\n--- Ziraat Bankası çekiliyor ---")
    try:
        satirlar = scrape_ziraat()
        banka_verileri["ZİRAAT"] = satirlar
        print(f"Ziraat: {len(satirlar)} satır bulundu.")
    except Exception as exc:
        print(f"[HATA] Ziraat çekilemedi: {exc}", file=sys.stderr)

    print("\n--- Halkbank çekiliyor ---")
    try:
        satirlar = scrape_halkbank()
        banka_verileri["HALKBANK"] = satirlar
        print(f"Halkbank: {len(satirlar)} satır bulundu.")
    except Exception as exc:
        print(f"[HATA] Halkbank çekilemedi: {exc}", file=sys.stderr)

    print("\n--- Akbank çekiliyor ---")
    try:
        satirlar = scrape_akbank()
        banka_verileri["AKBANK"] = satirlar
        print(f"Akbank: {len(satirlar)} satır bulundu.")
    except Exception as exc:
        print(f"[HATA] Akbank çekilemedi: {exc}", file=sys.stderr)

    print("\n--- Yapı Kredi çekiliyor ---")
    try:
        satirlar = scrape_yapikredi()
        banka_verileri["YAPIKREDI"] = satirlar
        print(f"Yapı Kredi: {len(satirlar)} satır bulundu.")
    except Exception as exc:
        print(f"[HATA] Yapı Kredi çekilemedi: {exc}", file=sys.stderr)

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
