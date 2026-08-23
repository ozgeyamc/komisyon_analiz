"""
Banka komisyon ücretleri takip botu - final güvenli ana akış.

Akış:
1) 10 doğrulanmış primary scraper çalışır.
2) Scraper sürümleri + global safety guard doğrulanır.
3) supplemental_sources.py resmî ikincil kaynaklarla eksikleri tamamlar.
4) Excel önce geçici dosyaya yazılır.
5) update_comparison.py aynı geçici dosyada KARŞILAŞTIRMA sayfasını üretir.
6) KARŞILAŞTIRMA sayfası gerçekten oluşmadan final Excel değiştirilmez.
7) Her şey başarılıysa geçici dosya atomik olarak komisyonlar_guncel.xlsx olur.

Kullanım:
    python main.py
"""

from __future__ import annotations

import shutil
import sys
import time
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from safety_guard import (
    BASELINE_COUNTS,
    MAX_RATIO,
    MIN_RATIO,
    print_guard_report,
    validate_run,
)
from supplemental_sources import (
    SUPPLEMENTAL_VERSION,
    enrich_all,
    print_supplemental_report,
)
import update_comparison as _comparison_module

# Eski/mixed update_comparison.py yanlışlıkla repoda kalırsa Python import
# aşamasında patlamak yerine aşağıdaki sürüm kontrolü anlaşılır hata verir.
COMPARISON_SHEET = getattr(
    _comparison_module,
    "COMPARISON_SHEET",
    "KARŞILAŞTIRMA",
)
COMPARISON_VERSION = getattr(
    _comparison_module,
    "COMPARISON_VERSION",
    "SÜRÜM_BİLGİSİ_YOK",
)
update_comparison_sheet = getattr(
    _comparison_module,
    "update_comparison_sheet",
    None,
)
from update_excel import (
    EXCEL_DOSYA_ADI,
    EXCEL_WRITER_VERSION,
    excel_guncelle_coklu,
    final_excel_gorunumunu_temizle,
)


MAIN_VERSION = "2026-08-23-v18-comparison-v22"
EXPECTED_SUPPLEMENTAL_VERSION = "2026-08-23-v11-user-audit-final"
EXPECTED_COMPARISON_VERSION = "2026-08-23-v22-final-gap-cleanup"
EXPECTED_EXCEL_WRITER_VERSION = "2026-08-21-v2-clean-supplemental-display"

# Primary scraper güvenlik retry ayarları.
# İlk deneme + 2 tekrar = toplam 3 deneme.
MAX_SCRAPE_ATTEMPTS = 3
RETRY_WAIT_SECONDS = 3


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


def _component_versions_ok() -> bool:
    ok = True
    print(f"[main] SÜRÜM: {MAIN_VERSION}")
    print(f"[main] supplemental: {SUPPLEMENTAL_VERSION}")
    print(f"[main] comparison: {COMPARISON_VERSION}")
    print(f"[main] excel writer: {EXCEL_WRITER_VERSION}")

    if SUPPLEMENTAL_VERSION != EXPECTED_SUPPLEMENTAL_VERSION:
        print(
            "[FATAL] supplemental_sources.py yanlış/eski sürüm. "
            f"Beklenen={EXPECTED_SUPPLEMENTAL_VERSION} | Gelen={SUPPLEMENTAL_VERSION}",
            file=sys.stderr,
        )
        ok = False

    if COMPARISON_VERSION != EXPECTED_COMPARISON_VERSION:
        print(
            "[FATAL] update_comparison.py yanlış/eski sürüm. "
            f"Beklenen={EXPECTED_COMPARISON_VERSION} | Gelen={COMPARISON_VERSION}",
            file=sys.stderr,
        )
        ok = False

    if not callable(update_comparison_sheet):
        print(
            "[FATAL] update_comparison.py içinde update_comparison_sheet() "
            "fonksiyonu bulunamadı.",
            file=sys.stderr,
        )
        ok = False

    if EXCEL_WRITER_VERSION != EXPECTED_EXCEL_WRITER_VERSION:
        print(
            "[FATAL] update_excel.py yanlış/eski sürüm. "
            f"Beklenen={EXPECTED_EXCEL_WRITER_VERSION} | Gelen={EXCEL_WRITER_VERSION}",
            file=sys.stderr,
        )
        ok = False

    return ok


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



def _bank_guard_problem_banks(guard) -> set[str]:
    """
    Guard raporundaki banka-bazlı hatalardan hangi scraper'ların yeniden
    denenmesi gerektiğini çıkarır.

    TOPLAM hatası tek başına varsa güvenli olmak için tüm bankalar tekrar
    denenir. Sürüm hataları burada çözülmez; onlar ayrıca bloklanır.
    """
    retry_banks: set[str] = set()

    for error in getattr(guard, "errors", ()):
        text = str(error)

        for bank in BASELINE_COUNTS:
            if text.startswith(f"{bank}:"):
                retry_banks.add(bank)

        if text.startswith("TOPLAM:") and not retry_banks:
            retry_banks.update(BASELINE_COUNTS)

    return retry_banks


def _count_status(bank: str, rows) -> str:
    baseline = BASELINE_COUNTS[bank]

    if rows is None:
        return f"veri yok | ref={baseline}"

    try:
        count = len(rows)
    except TypeError:
        rows = list(rows)
        count = len(rows)

    min_count = int(baseline * MIN_RATIO)
    max_count = int(baseline * MAX_RATIO)

    return (
        f"satır={count} | ref={baseline} | "
        f"güvenli_aralık={min_count}-{max_count}"
    )


def _retry_failed_primary_scrapers(
    primary_data: dict,
    scraper_functions: dict,
    version_errors: list[str],
):
    """
    İlk safety guard başarısız olduğunda yalnız problemli primary scraper'ları
    tekrar çalıştırır.

    Güvenlik sınırları DEĞİŞTİRİLMEZ.
    Bir tekrar başarılı olursa yeni veri kullanılır. Başarısız/None sonuç eski
    kullanılabilir sonucu silmez.
    """
    guard = validate_run(primary_data)

    if version_errors:
        for bank in version_errors:
            guard.errors.append(
                f"{bank}: doğrulanmış scraper sürümü yüklenmedi."
            )
        guard.ok = False

    if guard.ok:
        return primary_data, guard

    retry_banks = _bank_guard_problem_banks(guard)

    # Sürüm uyuşmazlığı retry ile çözülmez.
    retry_banks.difference_update(version_errors)

    if not retry_banks:
        return primary_data, guard

    print()
    print("=" * 68)
    print("PRIMARY OTOMATİK TEKRAR KONTROLÜ")
    print("=" * 68)
    print(
        "[retry] İlk güvenlik kontrolü geçmedi. "
        "Yalnız problemli bankalar tekrar çekilecek."
    )

    for bank in sorted(retry_banks):
        print(f"[retry] {bank}: {_count_status(bank, primary_data.get(bank))}")

    # İlk scrape zaten 1. denemeydi. Burada 2 ve 3. denemeleri yapıyoruz.
    for attempt in range(2, MAX_SCRAPE_ATTEMPTS + 1):
        if not retry_banks:
            break

        print()
        print(
            f"[retry] Tur {attempt}/{MAX_SCRAPE_ATTEMPTS} "
            f"| bankalar={', '.join(sorted(retry_banks))}"
        )

        if RETRY_WAIT_SECONDS:
            time.sleep(RETRY_WAIT_SECONDS)

        for bank in list(sorted(retry_banks)):
            fn = scraper_functions.get(bank)

            if fn is None:
                print(
                    f"[retry][ATLANDI] {bank}: scraper fonksiyonu hazır değil.",
                    file=sys.stderr,
                )
                continue

            print(f"\n--- {bank} yeniden çekiliyor ({attempt}/{MAX_SCRAPE_ATTEMPTS}) ---")
            rows = _try_scrape(bank, fn)

            # Başarısız tekrar, elimizdeki önceki veriyi silmesin.
            if rows is not None:
                primary_data[bank] = rows

        guard = validate_run(primary_data)

        if version_errors:
            for bank in version_errors:
                guard.errors.append(
                    f"{bank}: doğrulanmış scraper sürümü yüklenmedi."
                )
            guard.ok = False

        if guard.ok:
            print()
            print(
                f"[retry] BAŞARILI - güvenlik kontrolü "
                f"{attempt}. denemede geçti."
            )
            print("=" * 68)
            return primary_data, guard

        retry_banks = _bank_guard_problem_banks(guard)
        retry_banks.difference_update(version_errors)

        if retry_banks:
            for bank in sorted(retry_banks):
                print(
                    f"[retry] hâlâ sorunlu: {bank} | "
                    f"{_count_status(bank, primary_data.get(bank))}"
                )

    print()
    print(
        "[retry] Tekrarlar tamamlandı; güvenlik kontrolü hâlâ geçmiyor. "
        "Yanlış/eksik veri yazılmaması için Excel bloke kalacak."
    )
    print("=" * 68)

    return primary_data, guard


def _prepare_temp_excel(final_path: Path) -> Path:
    temp_path = final_path.with_name(final_path.stem + ".pipeline.tmp" + final_path.suffix)
    if temp_path.exists():
        temp_path.unlink()
    if final_path.exists():
        shutil.copy2(final_path, temp_path)
    return temp_path


def _cleanup(path: Path) -> None:
    try:
        if path.exists():
            path.unlink()
    except Exception as exc:
        print(f"[UYARI] Geçici dosya silinemedi: {path} | {exc}", file=sys.stderr)


def _xlsx_sheet_names(path: Path) -> list[str]:
    """openpyxl'e bağımlı olmadan final xlsx içindeki sheet isimlerini doğrular."""
    ns = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(path) as zf:
        root = ET.fromstring(zf.read("xl/workbook.xml"))
    sheets = root.find("x:sheets", ns)
    if sheets is None:
        return []
    return [node.attrib.get("name", "") for node in sheets]


def _verify_comparison_file(path: Path, comparison: dict) -> None:
    if not path.exists() or path.stat().st_size < 10_000:
        raise RuntimeError("Geçici Excel oluşmadı veya dosya beklenenden küçük.")

    names = _xlsx_sheet_names(path)
    if COMPARISON_SHEET not in names:
        raise RuntimeError(
            f"{COMPARISON_SHEET} sayfası Excel içinde bulunamadı. Sheetler={names}"
        )

    rows = int(comparison.get("comparison_rows", 0) or 0)
    possible = int(comparison.get("possible_cells", 0) or 0)
    missing = int(comparison.get("missing_cells", 0) or 0)
    source_gaps = int(comparison.get("source_gap_cells", 0) or 0)
    ambiguous = int(comparison.get("ambiguous_cells", 0) or 0)

    if rows < 40 or possible < 200:
        raise RuntimeError(
            "Karşılaştırma sayfası beklenenden küçük görünüyor: "
            f"satır={rows}, mantıksal_hücre={possible}"
        )

    # N/A artık eşleştirme mantığı hatası kabul edilir. Veri gerçekten bankada
    # ayrı tarife olarak yayımlanmıyorsa SOURCE_GAP / resmî durum metni kullanılır.
    if missing != 0:
        raise RuntimeError(
            f"Karşılaştırmada açıklamasız N/A kaldı: {missing}. Final Excel yayınlanmadı."
        )

    print(
        f"[main] KARŞILAŞTIRMA doğrulandı: sheet var | "
        f"satır={rows} | mantıksal_hücre={possible} | "
        f"kaynak_boşluğu={source_gaps} | belirsiz={ambiguous} | N/A={missing}"
    )


def main() -> int:
    print("=== Banka Komisyon Ücretleri Takip Botu ===")

    # En başta yanlış dosya/sürüm kullanımı yakalanır. Böylece geçmişteki
    # v1 update_comparison veya comparison çağırmayan eski main sessizce çalışamaz.
    if not _component_versions_ok():
        return 5

    primary_data = {}
    version_errors = []
    scraper_functions = {}

    # 1) PRIMARY SCRAPER'LAR
    for banka_adi, module_name, func_name in BANKA_SIRASI:
        print(f"\n--- {banka_adi} çekiliyor ---")
        try:
            module = __import__(module_name)
            if not _version_ok(banka_adi, module):
                version_errors.append(banka_adi)
                continue
            fn = getattr(module, func_name)
            scraper_functions[banka_adi] = fn
            satirlar = _try_scrape(banka_adi, fn)
            if satirlar is not None:
                primary_data[banka_adi] = satirlar
        except ImportError as exc:
            print(f"[HATA] {banka_adi} modülü bulunamadı: {exc}", file=sys.stderr)
        except AttributeError as exc:
            print(f"[HATA] {banka_adi} scraper fonksiyonu bulunamadı: {exc}", file=sys.stderr)
        except Exception as exc:
            print(f"[HATA] {banka_adi}: {exc}", file=sys.stderr)

    # 2) PRIMARY GLOBAL SAFETY GUARD + OTOMATİK RETRY
    #
    # İlk koşu geçmezse güvenlik sınırlarını gevşetmek yerine yalnız sorunlu
    # scraper'lar iki kez daha denenir. Böylece geçici site/timeout/DOM yükleme
    # sorunları toparlanabilir; kalıcı veya şüpheli veri değişiminde Excel yine
    # korunur.
    primary_data, guard = _retry_failed_primary_scrapers(
        primary_data,
        scraper_functions,
        version_errors,
    )

    print_guard_report(guard)

    if not guard.ok:
        print(
            "\n[SONUÇ] Primary güvenlik kontrolü, otomatik tekrarların ardından "
            "da geçmedi. Excel güncellenmedi; son doğru dosya korunuyor.",
            file=sys.stderr,
        )
        return 2

    primary_total = sum(len(v) for v in primary_data.values())

    # 3) RESMÎ EK KAYNAKLAR
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
        f"ek kaynak sonrası={enriched_total} | net ek={enriched_total - primary_total}"
    )

    # 4) ATOMİK EXCEL + KARŞILAŞTIRMA
    final_path = Path(EXCEL_DOSYA_ADI)
    temp_path = _prepare_temp_excel(final_path)

    try:
        ozet = excel_guncelle_coklu(enriched_data, str(temp_path))

        # Teknik supplemental marker'ları KARŞILAŞTIRMA eşleştirmesi için burada
        # hâlâ korunur. Önce karşılaştırma üretilir.
        comparison = update_comparison_sheet(str(temp_path))

        # Karşılaştırma tamamlandıktan sonra yalnız final KOMİSYONLAR görünümü
        # temizlenir. Böylece eşleştirme mantığı bozulmaz, kullanıcı teknik
        # SERVICE/CHANNEL/SUPPLEMENTAL metinlerini görmez.
        display_cleanup = final_excel_gorunumunu_temizle(str(temp_path))
        print(
            "[main] Final görünüm temizlendi: "
            f"açıklama={display_cleanup.get('temizlenen_aciklama', 0)} | "
            f"kaynak_linki={display_cleanup.get('kaynak_linki', 0)}"
        )

        _verify_comparison_file(temp_path, comparison)

        # Yalnız yukarıdaki tüm kontroller başarılıysa final dosya değiştirilir.
        temp_path.replace(final_path)

        # Replace sonrasında da sheet varlığını bir kez daha doğrula.
        final_names = _xlsx_sheet_names(final_path)
        if COMPARISON_SHEET not in final_names:
            raise RuntimeError(
                f"Final Excel'e geçişten sonra {COMPARISON_SHEET} kayboldu: {final_names}"
            )

    except Exception as exc:
        _cleanup(temp_path)
        print(
            f"[HATA] Excel/karşılaştırma pipeline'ı başarısız: {exc}\n"
            "[SONUÇ] Son doğru Excel korunuyor.",
            file=sys.stderr,
        )
        return 4

    print()
    print("=" * 68)
    print("Tamamlandı.")
    print(f"Primary veri: {primary_total} satır")
    print(f"Ek kaynak sonrası: {enriched_total} satır")
    print(f"Excel yazılan: {ozet.get('eklendi', enriched_total)} satır")
    print(
        "Karşılaştırma: "
        f"{comparison['comparison_rows']} satır | "
        f"doğrulanmış {comparison.get('matched_cells', '?')}/"
        f"{comparison.get('possible_cells', '?')} | "
        f"kaynak boşluğu {comparison.get('source_gap_cells', '?')} | "
        f"belirsiz {comparison.get('ambiguous_cells', '?')} | "
        f"N/A {comparison.get('missing_cells', '?')} | "
        f"not korundu {comparison['notes_preserved']}"
    )
    print(f"Sheetler: {_xlsx_sheet_names(final_path)}")
    print(f"Excel dosyası: {EXCEL_DOSYA_ADI}")
    print("=" * 68)
    return 0


if __name__ == "__main__":
    sys.exit(main())
