"""
Banka komisyon ücretleri takip botu - final güvenli ana akış.

Akış:
1) 10 doğrulanmış primary scraper çalışır.
2) Sürümler + global safety guard PRIMARY veri üzerinde doğrulanır.
3) Ana ücret sayfasında bulunmayan kritik/verifiye edilmiş resmî ek kaynaklar
   supplemental_sources.py ile taranır ve primary veri zenginleştirilir.
4) Kritik ek kaynaklardan biri bozulursa mevcut doğru Excel'e dokunulmaz.
5) Excel önce geçici dosyada üretilir.
6) Aynı geçici Excel üzerinde KARŞILAŞTIRMA sayfası üretilir/doğrulanır.
7) Her şey başarılıysa geçici dosya atomik olarak komisyonlar_guncel.xlsx olur.

Kullanım:
    python main.py
"""

from __future__ import annotations

import shutil
import sys
from pathlib import Path

from safety_guard import print_guard_report, validate_run
from supplemental_sources import enrich_all, print_supplemental_report
from update_comparison import update_comparison_sheet
from update_excel import EXCEL_DOSYA_ADI, excel_guncelle_coklu


MAIN_VERSION = "2026-08-20-v3-primary-plus-official-secondary-atomic"

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
            print(f"[HATA] {banka_adi}: scraper None döndürdü.", file=sys.stderr)
            return None
        print(f"{banka_adi}: {len(satirlar)} satır bulundu.")
        return satirlar
    except Exception as exc:
        print(f"[HATA] {banka_adi} çekilemedi: {exc}", file=sys.stderr)
        return None


def _version_ok(banka_adi: str, module) -> bool:
    expected = EXPECTED_VERSIONS[banka_adi]
    actual = getattr(module, "SCRAPER_VERSION", None)
    print(f"[{banka_adi}] modül: {getattr(module, '__file__', '?')}")
    print(f"[{banka_adi}] sürüm: {actual or 'SÜRÜM BİLGİSİ YOK'}")

    if actual != expected:
        print(
            f"[FATAL] {banka_adi} yanlış scraper sürümü. "
            f"Beklenen: {expected} | Gelen: {actual}",
            file=sys.stderr,
        )
        return False
    return True


def _prepare_temp_excel(final_path: Path) -> Path:
    temp_path = final_path.with_name(final_path.stem + ".pipeline.tmp" + final_path.suffix)
    if temp_path.exists():
        temp_path.unlink()

    # Mevcut dosyayı kopyalamak özellikle KARŞILAŞTIRMA/NOTLAR gibi korunacak
    # kullanıcı alanlarının update_excel tarafından korunabilmesini sağlar.
    if final_path.exists():
        shutil.copy2(final_path, temp_path)

    return temp_path


def _cleanup(path: Path) -> None:
    try:
        if path.exists():
            path.unlink()
    except Exception as exc:
        print(f"[UYARI] Geçici dosya silinemedi: {path} | {exc}", file=sys.stderr)


def main() -> int:
    print("=== Banka Komisyon Ücretleri Takip Botu ===")
    print(f"[main] SÜRÜM: {MAIN_VERSION}")

    primary_data = {}
    version_errors = []

    # ------------------------------------------------------------------
    # 1) PRIMARY SCRAPER'LAR
    # ------------------------------------------------------------------
    for banka_adi, module_name, func_name in BANKA_SIRASI:
        print(f"\n--- {banka_adi} çekiliyor ---")
        try:
            module = __import__(module_name)
            if not _version_ok(banka_adi, module):
                version_errors.append(banka_adi)
                continue

            fn = getattr(module, func_name)
            satirlar = _try_scrape(banka_adi, fn)
            if satirlar is not None:
                primary_data[banka_adi] = satirlar

        except ImportError as exc:
            print(f"[HATA] {banka_adi} modülü bulunamadı: {exc}", file=sys.stderr)
        except AttributeError as exc:
            print(f"[HATA] {banka_adi} scraper fonksiyonu bulunamadı: {exc}", file=sys.stderr)
        except Exception as exc:
            print(f"[HATA] {banka_adi}: {exc}", file=sys.stderr)

    # ------------------------------------------------------------------
    # 2) PRIMARY GLOBAL SAFETY GUARD
    # ------------------------------------------------------------------
    guard = validate_run(primary_data)

    if version_errors:
        for banka in version_errors:
            guard.errors.append(f"{banka}: doğrulanmış scraper sürümü yüklenmedi.")
        guard.ok = False

    print_guard_report(guard)

    if not guard.ok:
        print(
            "\n[SONUÇ] Primary güvenlik kontrolü geçmedi. "
            "Excel güncellenmedi; son doğru dosya korunuyor.",
            file=sys.stderr,
        )
        return 2

    primary_total = sum(len(v) for v in primary_data.values())

    # ------------------------------------------------------------------
    # 3) RESMÎ EK KAYNAKLAR
    # ------------------------------------------------------------------
    try:
        enriched_data, supplemental_report = enrich_all(primary_data)
    except Exception as exc:
        print(f"[FATAL] Ek resmî kaynak katmanı çalışamadı: {exc}", file=sys.stderr)
        return 3

    print_supplemental_report(supplemental_report)

    if not supplemental_report.ok:
        print(
            "\n[SONUÇ] Kritik ek resmî kaynak kontrolü geçmedi. "
            "Excel güncellenmedi; son doğru dosya korunuyor.",
            file=sys.stderr,
        )
        return 3

    enriched_total = sum(len(v) for v in enriched_data.values())
    print(
        f"[main] Primary toplam={primary_total} | "
        f"ek kaynak sonrası={enriched_total} | "
        f"net ek={enriched_total - primary_total}"
    )

    # ------------------------------------------------------------------
    # 4) ATOMİK EXCEL + KARŞILAŞTIRMA PIPELINE'I
    # ------------------------------------------------------------------
    final_path = Path(EXCEL_DOSYA_ADI)
    temp_path = _prepare_temp_excel(final_path)

    try:
        ozet = excel_guncelle_coklu(enriched_data, str(temp_path))
        comparison = update_comparison_sheet(str(temp_path))

        # Her şey başarılı olmadan final Excel'e dokunma.
        temp_path.replace(final_path)

    except Exception as exc:
        _cleanup(temp_path)
        print(
            f"[HATA] Excel/karşılaştırma pipeline'ı başarısız: {exc}\n"
            "[SONUÇ] Son doğru Excel korunuyor.",
            file=sys.stderr,
        )
        return 4

    print()
    print("=" * 64)
    print("Tamamlandı.")
    print(f"Primary veri: {primary_total} satır")
    print(f"Ek kaynak sonrası: {enriched_total} satır")
    print(f"Excel yazılan: {ozet.get('eklendi', enriched_total)} satır")
    print(
        "Karşılaştırma: "
        f"{comparison['comparison_rows']} satır | "
        f"çözülen {comparison.get('matched_cells', '?')}/"
        f"{comparison.get('possible_cells', '?')} | "
        f"N/A {comparison.get('missing_cells', '?')} | "
        f"not korundu {comparison['notes_preserved']}"
    )
    print(f"Excel dosyası: {EXCEL_DOSYA_ADI}")
    print("=" * 64)
    return 0


if __name__ == "__main__":
    sys.exit(main())
