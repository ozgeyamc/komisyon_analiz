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
from typing import Dict, Iterable, List, Mapping, Sequence, Tuple


# 20.08.2026 tarihinde 10 banka birlikte başarılı çalıştırılarak doğrulanan referanslar.
BASELINE_COUNTS: Dict[str, int] = {
    "GARANTİ": 715,
    "YAPIKREDI": 398,
    "İŞBANKASI": 946,
    "AKBANK": 461,
    "QNB": 390,
    "DENİZBANK": 492,
    "HALKBANK": 395,
    "VAKIFBANK": 378,
    "TEB": 542,
    "ZİRAAT": 586,
}

BASELINE_TOTAL = sum(BASELINE_COUNTS.values())  # 5303

# Banka bazında referansın %70'inden azı gelirse yazmayı durdur.
MIN_RATIO = 0.70

# Duplicate / DOM patlaması gibi durumları yakalamak için aşırı artışı da durdur.
MAX_RATIO = 1.75

# Toplam satır sayısı için ikinci emniyet kemeri.
TOTAL_MIN_RATIO = 0.80
TOTAL_MAX_RATIO = 1.40

# Bu seviyeyi aşan ama hâlâ güvenli aralıkta kalan değişimler logda UYARI olsun.
WARN_DELTA_RATIO = 0.15

# kategori veya masrafı boş satırlar bu orandan fazlaysa koşu güvenli sayılmaz.
MAX_INVALID_CORE_RATIO = 0.01


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

        # Exact duplicate oranı. Bu kontrol sadece uyarıdır; banka sitelerinde
        # nadiren birebir tekrarlar görülebilir.
        signatures = [_row_signature(row) for row in rows]
        duplicate_count = len(signatures) - len(set(signatures))

        if count and duplicate_count / count >= 0.05:
            warnings.append(
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
