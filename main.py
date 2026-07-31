"""
Garanti BBVA + Ziraat Bankası komisyon ücretleri takip botu.
"""

import sys

from scraper import ScraperError, scrape_garanti_bbva
from scraper_ziraat import ScraperError as ZiraatScraperError
from scraper_ziraat import scrape_ziraat
from update_excel import EXCEL_DOSYA_ADI, excel_guncelle_coklu


def main() -> int:
    print("=== Banka Komisyon Ücretleri Takip Botu ===")

    banka_verileri = {}

    # Garanti BBVA
    print("\n--- Garanti BBVA çekiliyor ---")
    try:
        garanti_satirlar = scrape_garanti_bbva()
        banka_verileri["GARANTİ"] = garanti_satirlar
        print(f"Garanti: {len(garanti_satirlar)} satır bulundu.")
    except (ScraperError, Exception) as exc:
        print(f"[HATA] Garanti çekilemedi: {exc}", file=sys.stderr)

    # Ziraat Bankası
    print("\n--- Ziraat Bankası çekiliyor ---")
    try:
        ziraat_satirlar = scrape_ziraat()
        banka_verileri["ZİRAAT"] = ziraat_satirlar
        print(f"Ziraat: {len(ziraat_satirlar)} satır bulundu.")
    except (ZiraatScraperError, Exception) as exc:
        print(f"[HATA] Ziraat çekilemedi: {exc}", file=sys.stderr)

    if not banka_verileri:
        print("[HATA] Hiçbir bankadan veri çekilemedi!", file=sys.stderr)
        return 1

    # Excel'e yaz
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
