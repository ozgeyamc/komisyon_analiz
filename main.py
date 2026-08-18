"""
Banka komisyon ücretleri takip botu - ana çalıştırma script'i.

Kullanım:
    python main.py
"""

import sys

from update_excel import EXCEL_DOSYA_ADI, excel_guncelle_coklu


def _try_scrape(banka_adi, fn):
    try:
        satirlar = fn()

        print(
            f"{banka_adi}: {len(satirlar)} satır bulundu."
        )

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

    # ---------------------------------------------------------
    # BANKA SIRASI
    # ---------------------------------------------------------

    BANKA_SIRASI = [

        (
            "GARANTİ",
            "scraper",
            "scrape_garanti_bbva"
        ),

        # YAPI KREDİ
        # Dosya adı:
        # scraper_yapikredi.py
        #
        # Fonksiyon:
        # scrape_yapikredi()
        (
            "YAPIKREDI",
            "scraper_yapikredi",
            "scrape_yapikredi"
        ),

        (
            "İŞBANKASI",
            "scraper_isbank",
            "scrape_isbank"
        ),

        (
            "AKBANK",
            "scraper_akbank",
            "scrape_akbank"
        ),

        (
            "QNB",
            "scraper_qnb",
            "scrape_qnb"
        ),

        (
            "DENİZBANK",
            "scraper_denizbank",
            "scrape_denizbank"
        ),

        (
            "HALKBANK",
            "scraper_halkbank",
            "scrape_halkbank"
        ),

        (
            "VAKIFBANK",
            "scraper_vakifbank",
            "scrape_vakifbank"
        ),

        (
            "TEB",
            "scraper_teb",
            "scrape_teb"
        ),

        (
            "ZİRAAT",
            "scraper_ziraat",
            "scrape_ziraat"
        ),
    ]

    banka_verileri = {}

    # ---------------------------------------------------------
    # BANKALARI TEK TEK ÇALIŞTIR
    # ---------------------------------------------------------

    for banka_adi, module, func in BANKA_SIRASI:

        print()
        print(
            f"--- {banka_adi} çekiliyor ---"
        )

        try:

            # Modülü yükle
            mod = __import__(module)

            # -------------------------------------------------
            # YAPI KREDİ DEBUG
            # -------------------------------------------------

            if banka_adi == "YAPIKREDI":

                print(
                    "[yapikredi] Kullanılan modül:"
                )

                print(
                    f"[yapikredi] {mod.__file__}"
                )

                # scraper içinde SCRAPER_VERSION varsa göster
                version = getattr(
                    mod,
                    "SCRAPER_VERSION",
                    None
                )

                if version:

                    print(
                        f"[yapikredi] SÜRÜM: {version}"
                    )

            # -------------------------------------------------
            # SCRAPER FONKSİYONUNU BUL
            # -------------------------------------------------

            fn = getattr(
                mod,
                func
            )

            # -------------------------------------------------
            # SCRAPER ÇALIŞTIR
            # -------------------------------------------------

            satirlar = _try_scrape(
                banka_adi,
                fn
            )

            if satirlar:

                banka_verileri[
                    banka_adi
                ] = satirlar

        except ImportError as exc:

            print(
                f"[UYARI] {banka_adi} modülü "
                f"bulunamadı, atlanıyor: {exc}",
                file=sys.stderr
            )

        except AttributeError as exc:

            print(
                f"[HATA] {banka_adi} scraper "
                f"fonksiyonu bulunamadı: {exc}",
                file=sys.stderr
            )

        except Exception as exc:

            print(
                f"[HATA] {banka_adi}: {exc}",
                file=sys.stderr
            )

    # ---------------------------------------------------------
    # HİÇ VERİ YOKSA
    # ---------------------------------------------------------

    if not banka_verileri:

        print(
            "[HATA] Hiçbir bankadan veri çekilemedi!",
            file=sys.stderr
        )

        return 1

    # ---------------------------------------------------------
    # EXCEL'E YAZ
    # ---------------------------------------------------------

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

    # ---------------------------------------------------------
    # SONUÇ
    # ---------------------------------------------------------

    print()
    print("=" * 60)

    print(
        f"Tamamlandı."
    )

    print(
        f"Toplam {ozet['eklendi']} satır yazıldı."
    )

    print(
        f"Excel dosyası: {EXCEL_DOSYA_ADI}"
    )

    print("=" * 60)

    return 0


if __name__ == "__main__":
    sys.exit(main())
