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
        print(
            f"[HATA] {banka_adi} çekilemedi: {exc}",
            file=sys.stderr
        )
        return None


def main() -> int:

    print(
        "=== Banka Komisyon Ücretleri Takip Botu ==="
    )

    BANKA_SIRASI = [
        ("GARANTİ",   "scraper",             "scrape_garanti_bbva"),
        ("YAPIKREDI", "scraper_yapikredi",   "scrape_yapikredi"),
        ("İŞBANKASI", "scraper_isbank",      "scrape_isbank"),
        ("AKBANK",    "scraper_akbank",      "scrape_akbank"),
        ("QNB",       "scraper_qnb",         "scrape_qnb"),
        ("DENİZBANK", "scraper_denizbank",   "scrape_denizbank"),
        ("HALKBANK",  "scraper_halkbank",    "scrape_halkbank"),
        ("VAKIFBANK", "scraper_vakifbank",   "scrape_vakifbank"),
        ("TEB",       "scraper_teb",         "scrape_teb"),
        ("ZİRAAT",    "scraper_ziraat",      "scrape_ziraat"),
    ]

    banka_verileri = {}

    for banka_adi, module, func in BANKA_SIRASI:

        print(
            f"\n--- {banka_adi} çekiliyor ---"
        )

        try:

            mod = __import__(module)

            fn = getattr(
                mod,
                func
            )

            satirlar = _try_scrape(
                banka_adi,
                fn
            )

            if satirlar:
                banka_verileri[
                    banka_adi
                ] = satirlar

        except ImportError:

            print(
                f"[UYARI] {banka_adi} modülü bulunamadı, "
                f"atlanıyor.",
                file=sys.stderr
            )

        except Exception as exc:

            print(
                f"[HATA] {banka_adi}: {exc}",
                file=sys.stderr
            )

    if not banka_verileri:

        print(
            "[HATA] Hiçbir bankadan veri çekilemedi!",
            file=sys.stderr
        )

        return 1

    try:

        ozet = excel_guncelle_coklu(
            banka_verileri,
            EXCEL_DOSYA_ADI
        )

    except Exception as exc:

        print(
            f"[HATA] Excel yazılırken hata: {exc}",
            file=sys.stderr
        )

        return 1

    print(
        f"\nTamamlandı. "
        f"Toplam {ozet['eklendi']} satır yazıldı."
    )

    print(
        f"Excel dosyası: {EXCEL_DOSYA_ADI}"
    )

    return 0


if __name__ == "__main__":
    sys.exit(main())
