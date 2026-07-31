"""
Garanti BBVA komisyon ücretleri takip botu - ana çalıştırma script'i.

Kullanım:
    python main.py
"""

import sys

from scraper import ScraperError, scrape_garanti_bbva
from update_excel import EXCEL_DOSYA_ADI, excel_guncelle


def main() -> int:
    print("=== Garanti BBVA Komisyon Ücretleri Takip Botu ===")

    try:
        satirlar = scrape_garanti_bbva()
    except ScraperError as exc:
        print(f"[HATA] Scraping başarısız oldu: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:
        print(f"[HATA] Beklenmeyen bir hata oluştu: {exc}", file=sys.stderr)
        return 1

    if not satirlar:
        print("[HATA] Hiçbir veri çekilemedi, Excel güncellenmedi.", file=sys.stderr)
        return 1

    try:
        ozet = excel_guncelle(satirlar, EXCEL_DOSYA_ADI)
    except Exception as exc:
        print(f"[HATA] Excel güncellenirken hata oluştu: {exc}", file=sys.stderr)
        return 1

    print(
        f"Tamamlandı. Eklenen: {ozet['eklendi']}, "
        f"Güncellenen: {ozet['guncellendi']}, "
        f"Değişmeyen: {ozet['degismedi']}"
    )
    print(f"Excel dosyası: {EXCEL_DOSYA_ADI}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
