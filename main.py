"""
Banka komisyon ücretleri takip botu - ana çalıştırma script'i.

Global güvenlik davranışı:
1) 10 scraper'ın tamamı çalıştırılır.
2) Scraper sürümleri doğrulanır.
3) Satır sayıları + temel veri kalitesi global guard'dan geçer.
4) TEK BİR banka bile kritik kontrolden kalırsa Excel'e hiç yazılmaz.
   Böylece son doğru komisyonlar_guncel.xlsx korunur.
5) Yalnız tüm kontrol OK ise excel_guncelle_coklu() çağrılır.

Kullanım:
    python main.py
"""

import sys

from safety_guard import print_guard_report, validate_run
from update_comparison import update_comparison_sheet
from update_excel import EXCEL_DOSYA_ADI, excel_guncelle_coklu


BANKA_SIRASI = [
    ("GARANTİ",   "scraper",            "scrape_garanti_bbva"),
    ("YAPIKREDI", "scraper_yapikredi",  "scrape_yapikredi"),
    ("İŞBANKASI", "scraper_isbank",     "scrape_isbank"),
    ("AKBANK",    "scraper_akbank",     "scrape_akbank"),
    ("QNB",       "scraper_qnb",        "scrape_qnb"),
    ("DENİZBANK", "scraper_denizbank",  "scrape_denizbank"),
    ("HALKBANK",  "scraper_halkbank",   "scrape_halkbank"),
    ("VAKIFBANK", "scraper_vakifbank",  "scrape_vakifbank"),
    ("TEB",       "scraper_teb",        "scrape_teb"),
    ("ZİRAAT",    "scraper_ziraat",     "scrape_ziraat"),
]


# Doğruladığımız/dondurduğumuz scraper sürümleri.
EXPECTED_VERSIONS = {
    "GARANTİ":   "2026-08-19-v2-garanti-integrity",
    "YAPIKREDI": "2026-08-18-v8-complete-fee-hierarchy",
    "İŞBANKASI": "2026-08-19-v3-isbank-transfer-boundaries",
    "AKBANK":    "2026-08-19-v4-akbank-context-fix",
    "QNB":       "2026-08-19-v5-qnb-category-fix",
    "DENİZBANK": "2026-08-19-v2-denizbank-integrity",
    "HALKBANK":  "2026-08-19-v3-halkbank-card-commercial-fix",
    "VAKIFBANK": "2026-08-19-v3-vakifbank-channel-currency-fix",
    "TEB":       "2026-08-19-v5-teb-category-normalization",
    "ZİRAAT":    "2026-08-19-v4-ziraat-section-boundary-fix",
}


def _try_scrape(banka_adi, fn):
    try:
        satirlar = fn()

        if satirlar is None:
            print(
                f"[HATA] {banka_adi}: scraper None döndürdü.",
                file=sys.stderr,
            )
            return None

        print(
            f"{banka_adi}: {len(satirlar)} satır bulundu."
        )
        return satirlar

    except Exception as exc:
        print(
            f"[HATA] {banka_adi} çekilemedi: {exc}",
            file=sys.stderr,
        )
        return None


def _version_ok(
    banka_adi: str,
    module,
) -> bool:
    expected = EXPECTED_VERSIONS[
        banka_adi
    ]

    actual = getattr(
        module,
        "SCRAPER_VERSION",
        None,
    )

    print(
        f"[{banka_adi}] modül: "
        f"{getattr(module, '__file__', '?')}",
        file=sys.stderr,
    )

    print(
        f"[{banka_adi}] sürüm: "
        f"{actual or 'SÜRÜM BİLGİSİ YOK'}",
        file=sys.stderr,
    )

    if actual != expected:
        print(
            f"[FATAL] {banka_adi} yanlış scraper sürümü. "
            f"Beklenen: {expected} | Gelen: {actual}",
            file=sys.stderr,
        )
        return False

    return True


def main() -> int:
    print(
        "=== Banka Komisyon Ücretleri Takip Botu ==="
    )

    banka_verileri = {}
    version_errors = []

    for (
        banka_adi,
        module_name,
        func_name,
    ) in BANKA_SIRASI:

        print(
            f"\n--- {banka_adi} çekiliyor ---"
        )

        try:
            module = __import__(
                module_name
            )

            if not _version_ok(
                banka_adi,
                module,
            ):
                version_errors.append(
                    banka_adi
                )
                continue

            fn = getattr(
                module,
                func_name,
            )

            satirlar = _try_scrape(
                banka_adi,
                fn,
            )

            if satirlar is not None:
                banka_verileri[
                    banka_adi
                ] = satirlar

        except ImportError as exc:
            print(
                f"[HATA] {banka_adi} modülü bulunamadı: {exc}",
                file=sys.stderr,
            )

        except AttributeError as exc:
            print(
                f"[HATA] {banka_adi} scraper fonksiyonu bulunamadı: {exc}",
                file=sys.stderr,
            )

        except Exception as exc:
            print(
                f"[HATA] {banka_adi}: {exc}",
                file=sys.stderr,
            )

    # -----------------------------------------------------
    # GLOBAL GÜVENLİK KAPISI
    # -----------------------------------------------------

    guard = validate_run(
        banka_verileri
    )

    # Yanlış scraper sürümü başlı başına fatal.
    if version_errors:
        for banka in version_errors:
            guard.errors.append(
                f"{banka}: doğrulanmış scraper sürümü yüklenmedi."
            )
        guard.ok = False

    print_guard_report(
        guard
    )

    if not guard.ok:
        print(
            "\n[SONUÇ] Excel güncellenmedi. "
            "Son doğru dosya korunuyor.",
            file=sys.stderr,
        )
        return 2

    # -----------------------------------------------------
    # EXCEL YAZIMI - SADECE GÜVENLİ KOŞUDA
    # -----------------------------------------------------

    try:
        ozet = excel_guncelle_coklu(
            banka_verileri,
            EXCEL_DOSYA_ADI,
        )

    except Exception as exc:
        print(
            f"[HATA] Excel yazılırken hata: {exc}",
            file=sys.stderr,
        )
        return 1

    # Ana veri Excel'i başarıyla yazıldıktan sonra aynı dosya içindeki
    # KARŞILAŞTIRMA sayfasını da yeniden üret.
    try:
        comparison = update_comparison_sheet(
            EXCEL_DOSYA_ADI
        )

    except Exception as exc:
        print(
            f"[HATA] KARŞILAŞTIRMA sayfası güncellenirken hata: {exc}",
            file=sys.stderr,
        )
        return 1

    print(
        f"Karşılaştırma sayfası: "
        f"{comparison['comparison_rows']} satır | "
        f"not korundu: {comparison['notes_preserved']}"
    )

    print()
    print("=" * 60)
    print("Tamamlandı.")
    print(
        f"Toplam {ozet['eklendi']} satır yazıldı."
    )
    print(
        f"Excel dosyası: {EXCEL_DOSYA_ADI}"
    )
    print("=" * 60)

    return 0


if __name__ == "__main__":
    sys.exit(
        main()
    )
