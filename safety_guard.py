"""
Global güvenlik katmanı.

Amaç:
- Bir scraper site değişikliği / hata nedeniyle çok az satır döndürürse
  mevcut doğru Excel'in üzerine eksik veri yazılmasını engellemek.
- Bir banka tamamen başarısız olursa Excel yazımını iptal etmek.
- Aşırı satır artışlarını (örn. duplicate patlaması) da durdurmak.
- Temel veri kalitesini (kategori/masraf boşluğu) kontrol etmek.

Bu modül Excel'e dokunmaz. Sadece "bu koşu güvenli mi?" kararını verir.
main.py, yalnızca bu kontrol OK ise excel_guncelle_coklu() çağırır.
"""

from dataclasses import dataclass
import re
import unicodedata
from typing import Dict, Iterable, List, Mapping, Sequence, Tuple


# 24.08.2026 resmî kaynak envanteriyle doğrulanan primary referanslar.
# Halkbank: 367 dinamik API + dört sayfalık ticari PDF'den 135 hizmet = 502.
BASELINE_COUNTS: Dict[str, int] = {
    "GARANTİ": 715,
    "YAPIKREDI": 398,
    "İŞBANKASI": 946,
    "AKBANK": 461,
    "QNB": 390,
    "DENİZBANK": 492,
    "HALKBANK": 502,
    "VAKIFBANK": 378,
    "TEB": 542,
    "ZİRAAT": 586,
}

BASELINE_TOTAL = sum(BASELINE_COUNTS.values())  # 5410

# Banka bazında referansın %85'inden azı gelirse yazmayı durdur.
# %70, bir alt sayfanın tamamen kaybolmasını dahi geçirebildiği için fazla
# gevşekti. Banka sitesinde gerçek ve büyük bir değişiklik varsa manuel olarak
# doğrulanmadan eski doğru Excel'in üzerine yazılmaması tercih edilir.
MIN_RATIO = 0.85

# Duplicate / DOM patlaması gibi durumları yakalamak için aşırı artışı da durdur.
MAX_RATIO = 1.75

# Toplam satır sayısı için ikinci emniyet kemeri.
TOTAL_MIN_RATIO = 0.80
TOTAL_MAX_RATIO = 1.40

# Bu seviyeyi aşan ama hâlâ güvenli aralıkta kalan değişimler logda UYARI olsun.
WARN_DELTA_RATIO = 0.08

# kategori veya masrafı boş satırlar bu orandan fazlaysa koşu güvenli sayılmaz.
MAX_INVALID_CORE_RATIO = 0.01

# Her banka için temel ürün ailelerinin kaybolmadığını doğrular. Bu kontrol
# sayısal satır eşiğinin tek başına yakalayamadığı bölüm-bazlı scraper
# eksiklerini bloke eder.
COMMON_REQUIRED_TERMS = (
    "fast",
    "eft",
    "havale",
    "swift",
    "atm",
    "cek",
    "fatura",
    "menkul",
)

BANK_REQUIRED_TERMS = {
    "HALKBANK": (
        "mkk hesap bakim ucreti",
        "mkk yatirimci sicil numarasi ve sifre gonderim ucreti",
        "mkk dibs alim satim islemleri ucreti",
        "mkk osba alim-satim ucreti",
        "uyelerarasi menkul kiymet transferi",
        "uye ici hesaplararasi",
        "nakit yonetimi",
        "uluslararasi fon transferi",
        "fast islemleri",
        "belge ve bilgilendirme",
        "cek defteri",
        "uye isyeri",
        "paraf klasik",
        "paraf gold",
        "paraf platinum",
    ),
}

MIN_DISTINCT_CATEGORIES = 5


@dataclass
class GuardResult:
    ok: bool
    errors: List[str]
    warnings: List[str]
    counts: Dict[str, int]
    total: int


def _row_signature(row) -> Tuple[str, ...]:
    """
    Exact duplicate kontrolü için scraper'ların ortak UcretSatiri alanlarını kullanır.
    Alan yoksa boş kabul edilir.
    """
    fields = (
        "kategori",
        "masraf",
        "asgari_tutar",
        "asgari_oran",
        "azami_tutar",
        "azami_oran",
        "aciklama",
        "site_guncelleme_tarihi",
    )
    return tuple(str(getattr(row, field, "") or "").strip() for field in fields)


def _core_is_valid(row) -> bool:
    kategori = str(getattr(row, "kategori", "") or "").strip()
    masraf = str(getattr(row, "masraf", "") or "").strip()
    return bool(kategori and masraf)


def _normalize_key(value) -> str:
    text = str(value or "").lower()
    text = text.translate(
        str.maketrans(
            {
                "ı": "i",
                "ğ": "g",
                "ü": "u",
                "ş": "s",
                "ö": "o",
                "ç": "c",
            }
        )
    )
    text = unicodedata.normalize("NFKD", text)
    text = "".join(
        char
        for char in text
        if not unicodedata.combining(char)
    )
    return re.sub(r"\s+", " ", text).strip()


def validate_run(
    banka_verileri: Mapping[str, Sequence],
) -> GuardResult:
    errors: List[str] = []
    warnings: List[str] = []
    counts: Dict[str, int] = {}

    expected_banks = set(BASELINE_COUNTS)
    received_banks = set(banka_verileri)

    missing = sorted(expected_banks - received_banks)
    unexpected = sorted(received_banks - expected_banks)

    for banka in missing:
        errors.append(
            f"{banka}: veri yok / scraper başarısız. "
            "Eksik banka varken Excel yazılamaz."
        )

    for banka in unexpected:
        warnings.append(
            f"{banka}: referans listesinde olmayan banka geldi; "
            "global toplam hesabına dahil edilecek ama ayrıca kontrol edilmeli."
        )

    for banka, baseline in BASELINE_COUNTS.items():
        rows = banka_verileri.get(banka)

        if rows is None:
            counts[banka] = 0
            continue

        try:
            count = len(rows)
        except TypeError:
            rows = list(rows)
            count = len(rows)

        counts[banka] = count

        min_count = int(baseline * MIN_RATIO)
        max_count = int(baseline * MAX_RATIO)

        if count == 0:
            errors.append(
                f"{banka}: 0 satır geldi. Referans {baseline}. "
                "Excel güncellemesi durdurulmalı."
            )
            continue

        if count < min_count:
            errors.append(
                f"{banka}: {count} satır geldi; referans {baseline}. "
                f"Alt güvenlik sınırı {min_count} (%{MIN_RATIO * 100:.0f})."
            )

        if count > max_count:
            errors.append(
                f"{banka}: {count} satır geldi; referans {baseline}. "
                f"Üst güvenlik sınırı {max_count} (%{MAX_RATIO * 100:.0f}). "
                "Duplicate/DOM patlaması olabilir."
            )

        delta_ratio = abs(count - baseline) / baseline

        if (
            delta_ratio >= WARN_DELTA_RATIO
            and min_count <= count <= max_count
        ):
            direction = "artış" if count > baseline else "düşüş"
            warnings.append(
                f"{banka}: referansa göre %{delta_ratio * 100:.1f} {direction} "
                f"({baseline} -> {count})."
            )

        # Temel satır kalitesi.
        invalid_core = sum(
            1
            for row in rows
            if not _core_is_valid(row)
        )

        allowed_invalid = max(
            1,
            int(count * MAX_INVALID_CORE_RATIO),
        )

        if invalid_core > allowed_invalid:
            errors.append(
                f"{banka}: {invalid_core}/{count} satırda KATEGORİ veya MASRAF boş. "
                f"İzin verilen en fazla {allowed_invalid}."
            )

        invalid_zero_dates = [
            row
            for row in rows
            if _normalize_key(getattr(row, "site_guncelleme_tarihi", ""))
            in {"30.12.1899", "30.12.1899 00:00", "30.12.1899 00:00:00"}
        ]
        if invalid_zero_dates:
            errors.append(
                f"{banka}: {len(invalid_zero_dates)} satırda geçersiz "
                "30.12.1899 seri-tarih artefaktı var."
            )

        if banka == "HALKBANK":
            malformed_fragments = {")", "de)", "erde)", "ilerde)"}
            malformed_rows = []
            for row in rows:
                numeric_fields = (
                    getattr(row, "asgari_tutar", ""),
                    getattr(row, "asgari_oran", ""),
                    getattr(row, "azami_tutar", ""),
                    getattr(row, "azami_oran", ""),
                )
                if any(
                    _normalize_key(value) in malformed_fragments
                    for value in numeric_fields
                ):
                    malformed_rows.append(row)
                    continue
                if any(
                    re.search(r"\d\s+[.,]\s*\d|\d\s+\d", str(value or ""))
                    for value in numeric_fields
                ):
                    malformed_rows.append(row)
            if malformed_rows:
                errors.append(
                    f"HALKBANK: {len(malformed_rows)} ticari PDF satırında "
                    "metin parçası tutar/oran kolonuna taşmış."
                )

        distinct_categories = {
            _normalize_key(
                getattr(row, "kategori", "")
            )
            for row in rows
            if _normalize_key(
                getattr(row, "kategori", "")
            )
        }

        if len(distinct_categories) < MIN_DISTINCT_CATEGORIES:
            errors.append(
                f"{banka}: yalnız {len(distinct_categories)} farklı kategori geldi; "
                f"beklenen en az {MIN_DISTINCT_CATEGORIES}. Alt sayfa/bölüm kaybı olabilir."
            )

        searchable = "\n".join(
            _normalize_key(
                f"{getattr(row, 'kategori', '')} "
                f"{getattr(row, 'masraf', '')}"
            )
            for row in rows
        )
        required_terms = (
            COMMON_REQUIRED_TERMS
            + BANK_REQUIRED_TERMS.get(banka, ())
        )
        missing_terms = [
            term
            for term in required_terms
            if term not in searchable
        ]

        if missing_terms:
            errors.append(
                f"{banka}: kritik ürün/işlem kalemleri eksik: "
                + ", ".join(missing_terms)
            )

        # Yüksek exact duplicate oranı DOM/API tekrarına işaret eder ve yanlış
        # satırların Excel'e yazılmasına izin verilmez.
        signatures = [_row_signature(row) for row in rows]
        duplicate_count = len(signatures) - len(set(signatures))

        if count and duplicate_count / count >= 0.05:
            errors.append(
                f"{banka}: {duplicate_count} exact duplicate satır var "
                f"(%{duplicate_count / count * 100:.1f})."
            )

    total = sum(counts.values())

    total_min = int(BASELINE_TOTAL * TOTAL_MIN_RATIO)
    total_max = int(BASELINE_TOTAL * TOTAL_MAX_RATIO)

    if total < total_min:
        errors.append(
            f"TOPLAM: {total} satır; referans {BASELINE_TOTAL}. "
            f"Alt toplam güvenlik sınırı {total_min}."
        )

    if total > total_max:
        errors.append(
            f"TOPLAM: {total} satır; referans {BASELINE_TOTAL}. "
            f"Üst toplam güvenlik sınırı {total_max}."
        )

    return GuardResult(
        ok=not errors,
        errors=errors,
        warnings=warnings,
        counts=counts,
        total=total,
    )


def print_guard_report(result: GuardResult) -> None:
    print()
    print("=" * 68)
    print("GLOBAL EXCEL GÜVENLİK KONTROLÜ")
    print("=" * 68)

    for banka, baseline in BASELINE_COUNTS.items():
        count = result.counts.get(banka, 0)
        delta = count - baseline

        if baseline:
            pct = (delta / baseline) * 100
        else:
            pct = 0.0

        status = "OK"

        min_count = int(baseline * MIN_RATIO)
        max_count = int(baseline * MAX_RATIO)

        if count < min_count or count > max_count:
            status = "BLOKE"

        print(
            f"{banka:<12} "
            f"{count:>5} | ref {baseline:>5} | "
            f"{pct:+7.1f}% | {status}"
        )

    print("-" * 68)
    total_delta_pct = (
        (result.total - BASELINE_TOTAL) / BASELINE_TOTAL * 100
        if BASELINE_TOTAL
        else 0.0
    )
    print(
        f"{'TOPLAM':<12} "
        f"{result.total:>5} | ref {BASELINE_TOTAL:>5} | "
        f"{total_delta_pct:+7.1f}%"
    )

    if result.warnings:
        print()
        print("[UYARILAR]")
        for warning in result.warnings:
            print(f" - {warning}")

    if result.errors:
        print()
        print("[BLOKE EDEN HATALAR]")
        for error in result.errors:
            print(f" - {error}")

    print()

    if result.ok:
        print(
            "[GÜVENLİK] OK - 10 banka kontrolü geçti. "
            "Excel yazımına izin verildi."
        )
    else:
        print(
            "[GÜVENLİK] BLOKE - Excel yazılmayacak. "
            "Mevcut komisyonlar_guncel.xlsx olduğu gibi korunacak."
        )

    print("=" * 68)
