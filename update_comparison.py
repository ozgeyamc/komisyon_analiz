"""
komisyonlar_guncel.xlsx içindeki banka ücretlerini ortak bir masraf sözlüğüne
çevirip KARŞILAŞTIRMA sayfasını otomatik üretir.

Temel fark:
- Bankaların MASRAF adlarını birebir eşleştirmez.
- Önce hizmeti (EFT, Havale, FAST, Fatura vb.), kanalı (Mobil/Şube/ATM)
  ve tutar aralığını semantik kurallarla çıkarır.
- Aynı hizmet + aynı kanal + aynı tutar aralığı olan farklı banka satırlarını
  tek bir ORTAK MASRAF ADI altında toplar.

Örnek:
  Garanti : "Elektronik Fon Transferi (EFT) Ücreti - EFT Garanti BBVA
             Mobil / İnternet (0-8300 TL)"
  Akbank  : "FAST / Akbank Mobil'den, İnternet'ten ... EFT ...
             (8.300 TL ve Altı)"
  YKB     : "EFT Gönderimi - İnternet/Mobil – 0 - 8.300 TL"
  İşbank  : "EFT Gönderilmesi - 0-8.300 TL"

hepsi şu anahtara bağlanır:
  EFT Gönderimi | Mobil | 0 TRY - 8.300 TRY

Ana Excel her gün güncellendiğinde KARŞILAŞTIRMA sayfası da aynı veriyle
tekrar üretilir. NOTLAR sütunundaki manuel notlar ortak satır anahtarına göre
korunur.
"""

import re
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from functools import lru_cache
from pathlib import Path
from typing import Dict, Iterable, List, Mapping, Optional, Sequence, Set, Tuple

from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


COMPARISON_VERSION = "2026-08-21-v16-precision-first-stable"
COMPARISON_SHEET = "KARŞILAŞTIRMA"
PREVIEW_LAYOUT_SIGNATURE = "4BANKS|A:I|J:L_EMPTY|M_NOTES|SCREENSHOT_ORDER|PRECISION_FIRST|LOGICAL_GENERAL_CELLS"

STATUS_AVAILABLE = "[SUPPLEMENTAL][AVAILABLE_NO_SEPARATE_FEE]"
STATUS_EMPTY = "[SUPPLEMENTAL][PUBLISHED_EMPTY]"
STATUS_NUMERIC = "[SUPPLEMENTAL][OFFICIAL_FEE]"
STATUS_NOT_APPLICABLE = "[SUPPLEMENTAL][NOT_APPLICABLE]"


# Eski denemelerde oluşmuş karşılaştırma sayfaları kullanıcıyı yanıltmasın.
# v9 çalıştığında KARŞILAŞTIRMA ile başlayan eski sayfaların tamamı silinir
# ve sadece preview ile aynı kompakt sayfa yeniden oluşturulur.

# Kullanıcının karşılaştırma şablonundaki banka seti.
BANKS = [
    "GARANTİ",
    "İŞBANKASI",
    "AKBANK",
    "YAPIKREDI",
]

DISPLAY_BANKS = {
    "GARANTİ": "Garanti BBVA",
    "İŞBANKASI": "İş Bankası",
    "AKBANK": "Akbank",
    "YAPIKREDI": "Yapı ve Kredi Bankası",
}

BANK_COLORS = {
    "GARANTİ": "70AD47",
    "İŞBANKASI": "2E75B6",
    "AKBANK": "E31E24",
    "YAPIKREDI": "17365D",
}


PRIMARY_SOURCE_URLS = {
    "GARANTİ": "https://www.garantibbva.com.tr/urun-ve-hizmet-ucretleri",
    "İŞBANKASI": "https://www.isbank.com.tr/urun-ve-hizmet-ucretleri",
    "AKBANK": "https://www.akbank.com/urun-ve-hizmet-ucretleri",
    "YAPIKREDI": "https://www.yapikredi.com.tr/bireysel-bankacilik/hesaplama-araclari/bireysel-urun-ve-hizmet-ucretleri",
}

HEADER_ALIASES = {
    "banka": "BANKA",
    "kategori": "KATEGORİ",
    "masraf": "MASRAF",
    "asgari tutar": "ASGARİ TUTAR",
    "asgari oran": "ASGARİ ORAN",
    "azami tutar": "AZAMİ TUTAR",
    "azami oran": "AZAMİ ORAN",
    "aciklama": "AÇIKLAMA",
    "site guncelleme tarihi": "GÜNCELLEME TARİHİ",
    "komisyon guncelleme tarihi": "GÜNCELLEME TARİHİ",
    "son kontrol": "SON KONTROL",
    "calistirilma tarihi": "SON KONTROL",
}


@dataclass(frozen=True)
class FeeRow:
    banka: str
    kategori: str
    masraf: str
    asgari_tutar: str = ""
    asgari_oran: str = ""
    azami_tutar: str = ""
    azami_oran: str = ""
    aciklama: str = ""
    guncelleme_tarihi: str = ""

    @property
    def text(self) -> str:
        return " | ".join(
            x for x in (self.kategori, self.masraf, self.aciklama) if x
        )


@dataclass(frozen=True)
class Band:
    low: Optional[float]
    high: Optional[float]


@dataclass(frozen=True)
class RowSpec:
    """Karşılaştırmada görünen ORTAK satır tanımı."""

    label: str
    service: str
    band_key: Optional[str] = None
    detail: Optional[str] = None
    split_channel: bool = True


# ---------------------------------------------------------------------------
# ORTAK MASRAF SÖZLÜĞÜ
# ---------------------------------------------------------------------------
# Buradaki label'lar artık banka sitesindeki ham MASRAF adı değildir.
# Tüm bankalar bu ortak isimlere eşleştirilir.
#
# SECTION satırı: ("SECTION", "başlık")
# ROW satırı    : ("ROW", RowSpec(...))
# ---------------------------------------------------------------------------

LAYOUT = [
    # Referans görseldeki iş akışına yakın sıra:
    # EFT -> Şans Oyunu -> Para Çekme/ATM -> Fatura -> Havale -> FAST/SWIFT
    # -> SGK -> Kasa -> Çek -> KKB -> HGS -> Vergi -> Senet -> diğer transferler.

    ("SECTION", "EFT"),
    ("ROW", RowSpec("0 TRY - 8.300 TRY", "EFT", "TRANSFER_1")),
    ("ROW", RowSpec("8.300,01 TRY - 399.000 TRY", "EFT", "TRANSFER_2")),
    ("ROW", RowSpec("399.000,01 TRY -", "EFT", "TRANSFER_3")),

    ("ROW", RowSpec("Şans Oyunu Ödemeleri", "SANS_OYUNU")),

    ("SECTION", "PARA ÇEKME / ATM"),
    ("ROW", RowSpec("Günlük Limit Üzeri Para Çekme", "LIMIT_UZERI_PARA_CEKME", split_channel=False)),
    ("ROW", RowSpec("Ortak ATM Para Çekme", "ORTAK_ATM_PARA_CEKME", split_channel=False)),
    ("ROW", RowSpec("Ortak ATM Bakiye Sorgulama", "BAKIYE_ATM_YURTICI", split_channel=False)),

    ("SECTION", "FATURA"),
    ("ROW", RowSpec("0 TRY - 149,99 TRY", "FATURA", "FATURA_1")),
    ("ROW", RowSpec("150 TRY -", "FATURA", "FATURA_2")),

    ("SECTION", "HAVALE"),
    ("ROW", RowSpec("0 TRY - 8.300 TRY", "HAVALE", "TRANSFER_1")),
    ("ROW", RowSpec("8.300,01 TRY - 399.000 TRY", "HAVALE", "TRANSFER_2")),
    ("ROW", RowSpec("399.000,01 TRY -", "HAVALE", "TRANSFER_3")),

    ("SECTION", "FAST"),
    ("ROW", RowSpec("0 TRY - 8.300 TRY", "FAST", "TRANSFER_1")),
    ("ROW", RowSpec("8.300,01 TRY - 399.000 TRY", "FAST", "TRANSFER_2")),
    ("ROW", RowSpec("399.000,01 TRY -", "FAST", "TRANSFER_3")),

    ("ROW", RowSpec("SWIFT - Gelen", "SWIFT_GELEN", split_channel=False)),
    ("ROW", RowSpec("SWIFT - Giden", "SWIFT_GIDEN")),
    ("ROW", RowSpec("Yurt Dışı FAST / Global FAST", "YURT_DISI_FAST")),
    ("ROW", RowSpec("Visa YP Direct Transfer", "VISA_YP_DIRECT")),

    ("SECTION", "SGK TAHSİLAT"),
    ("ROW", RowSpec("0 TRY - 99,99 TRY", "SGK", "SGK_1")),
    ("ROW", RowSpec("100 TRY -", "SGK", "SGK_2")),

    ("SECTION", "KİRALIK KASA"),
    ("ROW", RowSpec("Büyük Kasa", "KASA", detail="BUYUK", split_channel=False)),
    ("ROW", RowSpec("Orta Kasa", "KASA", detail="ORTA", split_channel=False)),
    ("ROW", RowSpec("Küçük Kasa", "KASA", detail="KUCUK", split_channel=False)),
    ("ROW", RowSpec("Özel / Süper Kasa", "KASA", detail="OZEL", split_channel=False)),

    ("SECTION", "ÇEK"),
    ("ROW", RowSpec("Çek Defteri / Karekodlu Çek Karnesi", "CEK_DEFTERI", split_channel=False)),
    ("ROW", RowSpec("Çek Düzenleme", "CEK_DUZENLEME", split_channel=False)),
    ("ROW", RowSpec("Özel Nitelikli / Dövizli Çek Düzenleme", "CEK_OZEL", split_channel=False)),
    ("ROW", RowSpec("Çek İade", "CEK_IADE", split_channel=False)),
    ("ROW", RowSpec("Çek Tahsil - Aynı Banka", "CEK_TAHSIL", detail="AYNI", split_channel=False)),
    ("ROW", RowSpec("Çek Tahsil - Diğer Banka", "CEK_TAHSIL", detail="DIGER", split_channel=False)),
    ("ROW", RowSpec("Çek Tahsil - Döviz Çeki", "CEK_TAHSIL", detail="DOVIZ", split_channel=False)),
    ("ROW", RowSpec("Karşılıksız Çek Belgelendirme", "CEK_KARSILIKSIZ", split_channel=False)),
    ("ROW", RowSpec("Çek Düzeltme Hakkı", "CEK_DUZELTME_HAKKI", split_channel=False)),

    ("ROW", RowSpec("KKB Risk Raporu", "KREDI_RISK", split_channel=False)),
    ("ROW", RowSpec("HGS Etiket / Kart Bedeli", "HGS", split_channel=False)),
    ("ROW", RowSpec("Vergi Tahsilat Komisyonu", "VERGI")),

    ("SECTION", "SENET"),
    ("ROW", RowSpec("Senet İade", "SENET_IADE", split_channel=False)),
    ("ROW", RowSpec("Senet Protesto", "SENET_PROTESTO", split_channel=False)),
    ("ROW", RowSpec("Senet Protesto Kaldırma", "SENET_PROTESTO_KALDIRMA", split_channel=False)),
    ("ROW", RowSpec("Senet Tahsil - Aynı Banka", "SENET_TAHSIL", detail="AYNI", split_channel=False)),
    ("ROW", RowSpec("Senet Tahsil - Muhabir / Diğer Banka", "SENET_TAHSIL", detail="DIGER", split_channel=False)),

    ("SECTION", "DÜZENLİ TRANSFERLER"),
    ("ROW", RowSpec("Düzenli EFT - 0 TRY - 8.300 TRY", "DUZENLI_EFT", "TRANSFER_1")),
    ("ROW", RowSpec("Düzenli EFT - 8.300,01 TRY - 399.000 TRY", "DUZENLI_EFT", "TRANSFER_2")),
    ("ROW", RowSpec("Düzenli EFT - 399.000,01 TRY -", "DUZENLI_EFT", "TRANSFER_3")),
    ("ROW", RowSpec("Düzenli Havale - 0 TRY - 8.300 TRY", "DUZENLI_HAVALE", "TRANSFER_1")),
    ("ROW", RowSpec("Düzenli Havale - 8.300,01 TRY - 399.000 TRY", "DUZENLI_HAVALE", "TRANSFER_2")),
    ("ROW", RowSpec("Düzenli Havale - 399.000,01 TRY -", "DUZENLI_HAVALE", "TRANSFER_3")),

    ("ROW", RowSpec("Ortak ATM / Başka Kuruluş Para Yatırma", "PARA_YATIRMA", split_channel=False)),
    ("ROW", RowSpec("KKB Çek Bilgileri / Çek Risk Raporu", "CEK_RISK", split_channel=False)),
    ("ROW", RowSpec("Elektronik Altın / Altın Transferi", "ALTIN_TRANSFER", split_channel=False)),
    ("ROW", RowSpec("Kıymetli Maden Fiziki Teslimi", "KIYMETLI_MADEN_TESLIM", split_channel=False)),

    ("SECTION", "Aidat Ödemeleri"),
    ("ROW", RowSpec("0 TRY - 149,99 TRY", "AIDAT", "FATURA_1")),
    ("ROW", RowSpec("150 TRY -", "AIDAT", "FATURA_2")),
    ("SECTION", "Özel Okul Ödeme"),
    ("ROW", RowSpec("0 TRY - 149,99 TRY", "OZEL_OKUL", "FATURA_1")),
    ("ROW", RowSpec("150 TRY -", "OZEL_OKUL", "FATURA_2")),
    ("ROW", RowSpec("Telefon Ödemeleri", "TELEFON")),

    ("SECTION", "BELGE / RAPOR / ARAŞTIRMA"),
    ("ROW", RowSpec("Arşiv Araştırma Ücreti", "ARSIV", split_channel=False)),
    ("ROW", RowSpec("Mevduat Araştırma", "MEVDUAT_ARASTIRMA", split_channel=False)),
    ("ROW", RowSpec("Referans Mektubu", "REFERANS_MEKTUBU", split_channel=False)),
    ("ROW", RowSpec("Vize ve Özel Okullar İçin Mektup", "VIZE_MEKTUBU", split_channel=False)),
    ("ROW", RowSpec("Hesap Özeti Verilmesi", "HESAP_OZETI", split_channel=False)),
    ("ROW", RowSpec("Hesap Araştırma Talebi", "HESAP_ARASTIRMA", split_channel=False)),
    ("ROW", RowSpec("Borcu Yoktur Yazısı", "BORCU_YOKTUR", split_channel=False)),
    ("ROW", RowSpec("Hesap Özeti Posta Yoluyla", "HESAP_OZETI_POSTA", split_channel=False)),
    ("ROW", RowSpec("Bakiye Sorma - Yurtdışı ATM", "BAKIYE_ATM_YURTDISI", split_channel=False)),
]



# ---------------------------------------------------------------------------
# NORMALİZASYON
# ---------------------------------------------------------------------------

@lru_cache(maxsize=65536)
def _norm(value) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\u0307", "").replace("\xa0", " ")
    text = " ".join(text.split())
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.translate(str.maketrans({
        "ı": "i", "İ": "i", "ğ": "g", "Ğ": "g",
        "ü": "u", "Ü": "u", "ş": "s", "Ş": "s",
        "ö": "o", "Ö": "o", "ç": "c", "Ç": "c",
    }))
    return text.lower().strip()


def _clean(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _parse_number(raw: str) -> Optional[float]:
    s = re.sub(r"[^0-9.,]", "", raw or "")
    if not s:
        return None

    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        parts = s.split(".")
        if len(parts) > 1 and all(len(part) == 3 for part in parts[1:]):
            s = "".join(parts)

    try:
        return float(s)
    except ValueError:
        return None


def _parse_band(text: str) -> Optional[Band]:
    """MASRAF adından işlem TUTARI aralığını çıkarır; ücret kolonlarına bakmaz."""
    t = _norm(text)

    # 8.300 TL ve altı
    m = re.search(
        r"(\d[\d.,]*)\s*(?:tl|try)?\s*(?:ve\s*)?(?:alti|altinda)",
        t,
    )
    if m:
        high = _parse_number(m.group(1))
        if high is not None:
            return Band(0.0, high)

    # 399.000,01 TL ve üzeri / üstü
    m = re.search(
        r"(\d[\d.,]*)\s*(?:tl|try)?\s*(?:ve\s*)?(?:uzeri|ustu)",
        t,
    )
    if m:
        low = _parse_number(m.group(1))
        if low is not None:
            return Band(low, None)

    # 8.300,01 TL - 399.000 TL
    for m in re.finditer(
        r"(\d[\d.,]*)\s*(?:tl|try)?\s*[-–—]\s*"
        r"(\d[\d.,]*)\s*(?:tl|try)?",
        t,
    ):
        local = t[max(0, m.start() - 6): min(len(t), m.end() + 6)]
        if ":" in local:  # saat aralığı (16:00-17:15) olmasın
            continue

        low = _parse_number(m.group(1))
        high = _parse_number(m.group(2))

        if low is None or high is None or high < low:
            continue

        return Band(low, high)

    return None



@lru_cache(maxsize=32768)
def _all_band_keys(text: str) -> Set[str]:
    """
    MASRAF + AÇIKLAMA içinde geçen bütün standart tutar bantlarını bulur.

    Böylece:
      - "149,99 TL'ye kadar ... 150 TL ve üzeri ..."
      - "0-150 TL ... 150,01 TL ve üzeri"
      - "100 TL altı ... 100 TL ve üzeri"
    gibi tek satır içinde iki farklı kademe yayınlayan bankalar da
    iki karşılaştırma satırına bağlanabilir.
    """
    t = _norm(text)
    bands: List[Band] = []

    # "8.300 TL ve altı", "100 TL altında"
    for m in re.finditer(
        r"(\d[\d.,]*)\s*(?:tl|try)?\s*(?:ve\s*)?(?:alti|altinda)",
        t,
    ):
        high = _parse_number(m.group(1))
        if high is not None:
            bands.append(Band(0.0, high))

    # "149,99 TL'ye kadar"
    for m in re.finditer(
        r"(\d[\d.,]*)\s*(?:tl|try)?(?:'?[yea]*)?\s*kadar",
        t,
    ):
        high = _parse_number(m.group(1))
        if high is not None:
            bands.append(Band(0.0, high))

    # "399.000,01 TL ve üzeri / üstü"
    for m in re.finditer(
        r"(\d[\d.,]*)\s*(?:tl|try)?\s*(?:ve\s*)?(?:uzeri|ustu)",
        t,
    ):
        low = _parse_number(m.group(1))
        if low is not None:
            bands.append(Band(low, None))

    # Normal aralıklar: 0-8.300 / 8.300,01-399.000
    for m in re.finditer(
        r"(\d[\d.,]*)\s*(?:tl|try)?\s*[-–—]\s*"
        r"(\d[\d.,]*)\s*(?:tl|try)?(?:\s*arasi)?",
        t,
    ):
        local = t[max(0, m.start() - 6): min(len(t), m.end() + 6)]

        if ":" in local:
            continue

        low = _parse_number(m.group(1))
        high = _parse_number(m.group(2))

        if low is None or high is None or high < low:
            continue

        bands.append(Band(low, high))

    keys: Set[str] = set()

    for band in bands:
        key = _band_key(band)

        if key and not key.startswith("RAW:"):
            keys.add(key)

    return keys


def _close(a: Optional[float], b: Optional[float], tolerance: float = 1.01) -> bool:
    if a is None or b is None:
        return a is b
    return abs(a - b) <= tolerance


def _band_key(band: Optional[Band]) -> Optional[str]:
    if band is None:
        return None

    low, high = band.low, band.high

    # EFT / Havale / FAST mevzuat dilimleri.
    if low is not None and low <= 1.0 and _close(high, 8300.0):
        return "TRANSFER_1"
    if _close(low, 8300.01) and _close(high, 399000.0):
        return "TRANSFER_2"
    if low is not None and abs(low - 399000.01) <= 1.01 and high is None:
        return "TRANSFER_3"

    # Fatura / Aidat / Okul örnek şablon dilimleri.
    if low is not None and low <= 1.0 and _close(high, 149.99):
        return "FATURA_1"
    if low is not None and abs(low - 150.0) <= 1.01 and high is None:
        return "FATURA_2"

    # SGK örnek şablon dilimleri.
    if low is not None and low <= 1.0 and _close(high, 99.99):
        return "SGK_1"
    if low is not None and abs(low - 100.0) <= 1.01 and high is None:
        return "SGK_2"

    return f"RAW:{low}:{high}"


# ---------------------------------------------------------------------------
# HİZMET / KANAL SINIFLANDIRMA
# ---------------------------------------------------------------------------

@lru_cache(maxsize=16384)
def _service_tags(row: FeeRow) -> Set[str]:
    """Farklı bankaların adlandırmalarını ortak hizmet kimliklerine çevirir."""
    cat = _norm(row.kategori)
    mas = _norm(row.masraf)
    short = f"{cat} | {mas}"
    full = _norm(row.text)

    tags: Set[str] = set()

    # Supplemental satırları kendi kanonik hizmetini açıkça taşır.
    marker = re.search(r"service\s*=\s*([a-z0-9_]+)", full, flags=re.I)
    if marker:
        tags.add(marker.group(1).upper())

    # ---------------- PARA TRANSFERİ ----------------
    international = any(x in full for x in (
        "swift", "uluslararasi fon transfer", "yurt disi fast",
        "global fast", "western union", "fast uluslararasi",
    ))
    package = any(x in short for x in ("paket", "kota"))
    card_cash = any(x in short for x in ("nakit avans", "faiz orani"))

    if (any(x in full for x in ("eft", "elektronik fon transfer"))
            and not international and not package and "altin eft" not in full):
        if not card_cash or any(x in short for x in (
            "eft ucreti", "eft gonder", "elektronik fon transfer",
        )):
            tags.add("EFT")

    if (any(x in full for x in ("fast", "fonlarin anlik ve surekli transferi"))
            and not international and not package):
        tags.add("FAST")

    if (any(x in full for x in ("havale", "hesaptan hesaba"))
            and not international and not package):
        if not card_cash or any(x in short for x in (
            "havale ucreti", "havale gonder", "hesaptan hesaba",
        )):
            tags.add("HAVALE")

    # Referans karşılaştırma tablosundaki ek transfer aileleri.
    if any(x in full for x in ("swift", "uluslararasi fon transfer")):
        if any(x in full for x in ("gelen swift", "gelen doviz", "gelen uluslararasi", "gelen fon transfer")):
            tags.add("SWIFT_GELEN")
        if any(x in full for x in ("giden swift", "giden doviz", "doviz havale", "hesaptan giden", "gonderim")):
            tags.add("SWIFT_GIDEN")

    if any(x in full for x in (
        "yurt disi fast", "yurtdisi fast", "global fast",
        "fast uluslararasi", "fast uluslararasi turkiye disina",
    )):
        tags.add("YURT_DISI_FAST")

    if any(x in full for x in (
        "visa ile yurt disi para transferi", "visa direct",
        "visa yp direct", "visa ile yurtdisi para transferi",
    )):
        tags.add("VISA_YP_DIRECT")

    if "duzenli" in full and any(x in full for x in ("eft", "elektronik fon transfer")):
        tags.add("DUZENLI_EFT")
    if "duzenli" in full and "havale" in full:
        tags.add("DUZENLI_HAVALE")

    # ---------------- ATM / NAKİT ----------------
    if any(x in full for x in ("gunluk limit uzeri para cekme", "limit uzeri para cekme", "gunluk limit ustu para cekme", "limit ustu para cekme")):
        tags.add("LIMIT_UZERI_PARA_CEKME")

    if (
        "para cekme" in full
        and any(x in full for x in ("ortak atm", "diger banka atm", "baska kurulus", "baska banka atm"))
        and "limit uzeri" not in full
        and "nakit avans" not in full
    ):
        tags.add("ORTAK_ATM_PARA_CEKME")

    if (
        "para yatirma" in full
        and any(x in full for x in ("ortak atm", "diger banka atm", "baska kurulus", "baska banka atm", "atm"))
        and "nakit avans" not in mas
    ):
        tags.add("PARA_YATIRMA")

    if (
        any(x in mas for x in (
            "altin transfer",
            "ats ile altin gonderimi",
            "kiymetli maden transferi ucreti - altin",
            "kiymetli maden transfer - altin",
        ))
        and not any(x in mas for x in (
            "fiziki", "teslim", "kulce altin cekme",
            "western union", "eft", "havale", "fast",
        ))
    ):
        tags.add("ALTIN_TRANSFER")

    # ---------------- KASA / RAPOR ----------------
    if any(x in full for x in ("kiralik kasa", "kasa kiralama", "kasa ucreti")):
        tags.add("KASA")

    if any(x in full for x in (
        "kiymetli maden teslim", "fiziki altin teslim", "altin teslim",
        "fiziki kiymetli maden", "kulce altin cekme", "kulce altin teslim",
        "fiziki altin cekme",
    )):
        tags.add("KIYMETLI_MADEN_TESLIM")

    # Garanti gibi bazı bankalar "KKB Çek / Risk Raporu"nu tek satırda
    # yayımlar. Böyle bir satır hem çek risk hem genel risk raporunu temsil eder.
    combined_risk = any(x in full for x in (
        "kkb cek / risk raporu", "kkb cek/risk raporu", "cek / risk raporu",
    ))
    if combined_risk:
        tags.update({"CEK_RISK", "KREDI_RISK"})
    else:
        if any(x in full for x in (
            "cek risk raporu", "cek bilgileri raporu", "kkb cek", "findeks cek raporu",
            "cek sorgu raporu",
        )):
            tags.add("CEK_RISK")
        if any(x in full for x in (
            "risk raporu", "kkb risk", "kredi risk", "findeks risk", "risk merkezi raporu",
        )):
            tags.add("KREDI_RISK")

    # ---------------- TAHSİLAT / ÖDEME ----------------
    if (("fatura" in short or "fatura / kurum" in short or "fatura/kurum" in short
         or "fatura tahsil" in short or "kurum tahsil" in short)
            and "e-fatura" not in short):
        tags.add("FATURA")

    if "sgk" in short or "sosyal guvenlik" in full:
        tags.add("SGK")
    if "hgs" in full:
        tags.add("HGS")
    if "sans oyun" in full or any(x in full for x in (
        "bilyoner", "nesine", "tuttur", "oley", "misli", "sisal sans", "tjk",
    )):
        tags.add("SANS_OYUNU")

    # Aidat ödeme hizmetini kredi kartı "Aidatsız Kart" / kart aidatı kavramından ayır.
    has_aidat_word = bool(re.search(r"\baidat\b", short)) or (
        bool(re.search(r"\baidat\b", full))
        and any(x in full for x in ("fatura", "tahsilat", "odeme", "site", "apartman"))
    )
    if has_aidat_word and not any(x in full for x in (
        "aidatsiz kart", "kart aidati", "yillik kart ucreti", "uyelik ucreti - kart",
    )):
        tags.add("AIDAT")

    # Özel okul ÖDEME hizmetini "Vize ve Özel Okullar İçin Düzenlenen Mektup"
    # gibi belge hizmetlerinden kesin olarak ayır.
    if (any(x in full for x in (
        "ozel okul", "okul odeme", "okul taksiti", "egitim odeme", "egitim kurumu odeme",
    )) and not any(x in full for x in (
        "mektup", "vize", "konsolosluk", "referans yazisi", "referans mektubu",
    ))):
        tags.add("OZEL_OKUL")

    telefon_candidate = (
        any(x in short for x in (
            "telefon odeme", "telefon fatur", "cep telefonu fatur",
            "telefon operatorleri odemelerine aracilik",
            "gsm odeme", "telekom odeme", "turkcell", "vodafone",
        ))
        or (
            any(x in full for x in ("turkcell", "vodafone", "superonline", "tellcom", "turk telekom"))
            and any(x in short for x in ("fatura", "kurum odeme"))
        )
    )
    if (
        telefon_candidate
        and "tl/paket yukleme" not in mas
        and "paket yukleme" not in mas
        and "otomatik fatura odeme" not in mas
        and "alisveris faiz" not in mas
    ):
        tags.add("TELEFON")

    if ("vergi" in short and any(x in short for x in (
        "vergi tahsil", "vergi odeme", "fatura/vergi/sgk", "fatura / vergi / sgk", "mtv",
    )) and not any(x in short for x in ("vergi numarasi", "vergi yazisi", "kredi"))):
        tags.add("VERGI")

    # ---------------- BELGE / HESAP ----------------
    if (any(x in full for x in (
        "arsiv arastirma", "gecmis donem ekstre arsiv", "gecmis donem bankacilik islemleri bildirimi",
        "arsivden belge", "dokuman arastirma", "dokuman talebi",
        "1 yildan eski islemlere ait gecmise yonelik dekont", "gecmise yonelik dekont",
        "gecmis donem ekstre", "dekont masrafi - gecmise yonelik",
    )) or ("arsiv" in full and "arastirma" in full)):
        tags.add("ARSIV")

    if (any(x in full for x in ("mevduat arastirma", "mevduat hesap arastirma"))
            or ("arsiv" in cat and "arastirma" in cat and mas.startswith("merkezden"))):
        tags.add("MEVDUAT_ARASTIRMA")

    if any(x in full for x in (
        "referans mektubu", "referans yazisi", "banka referans", "itibar /niyet/referans",
        "itibar/niyet/referans",
    )):
        tags.add("REFERANS_MEKTUBU")

    if any(x in full for x in (
        "vize icin", "konsolosluk icin mektup", "konsolosluk icin",
        "ozel okullar icin duzenlenen mektup", "ozel okul icin duzenlenen mektup",
    )):
        tags.add("VIZE_MEKTUBU")

    if ("hesap ozeti" in short or "ekstre masraf" in short or "ekstre ucret" in short
            or "ekstre veril" in short or "ekstre - " in short):
        tags.add("HESAP_OZETI")

    if any(x in full for x in (
        "posta ile aylik hesap ozeti", "basili ekstre", "basili hesap ozeti",
        "ekstre gonderim", "hesap ozeti gonderimi", "hesap ozeti posta", "ekstre posta",
        "gecmis donem kredi karti hesap ozeti gonderimi",
    )):
        tags.update({"HESAP_OZETI", "HESAP_OZETI_POSTA"})

    if any(x in full for x in ("hesap arastirma", "hesap tespit", "hesap bulma")):
        tags.add("HESAP_ARASTIRMA")
    if any(x in full for x in ("borcu yoktur", "borc yoktur", "borcsuzluk")):
        tags.add("BORCU_YOKTUR")

    # ---------------- ATM BAKİYE ----------------
    if (any(x in full for x in ("bakiye sorma", "bakiye sorgu", "bakiye goruntule"))
            and (any(x in full for x in ("atm", "bankamatik"))
                 or "baska kurulus araciligiyla yapilan islemler" in full)):
        if any(x in full for x in ("yurt disi", "yurtdisi")):
            tags.add("BAKIYE_ATM_YURTDISI")
        else:
            tags.add("BAKIYE_ATM_YURTICI")

    # ---------------- ÇEK ----------------
    if any(x in full for x in ("cek defteri", "cek karnesi", "cek yaprak", "cek kitabi")):
        tags.add("CEK_DEFTERI")

    if (any(x in full for x in (
        "cek duzenleme", "cek duzenlenmesi", "bloke cek duzenleme", "dovizi natik cek duzenleme",
    )) and "ozel nitelikli" not in full):
        tags.add("CEK_DUZENLEME")

    if any(x in full for x in (
        "ozel nitelikli cek", "ozel cek duzenleme", "dovizli cek duzenleme",
        "dovizi natik cek duzenleme", "dth'dan cek duzenlenmesi", "dth dan cek duzenlenmesi",
        "seyahat ceki duzenleme",
    )):
        tags.add("CEK_OZEL")

    if any(x in full for x in (
        "cek iade", "cek iadesi", "cek muamelesiz iade", "cekin islemsiz iades",
        "ceklerin islemsiz iades",
    )):
        tags.add("CEK_IADE")

    if (any(x in full for x in ("cek tahsil", "tahsile alinan cek", "cek takas", "cek odeme"))
            or ("tahsile alinan" in full and "cek" in full)):
        tags.add("CEK_TAHSIL")

    # Karşılıksız çek BELGELENDİRME satırını düzeltme hakkından ayır.
    # Sırf "karşılıksız çek" geçmesi belgelendirme için yeterli değildir.
    if ("karsiliksiz cek" in full and any(x in full for x in (
        "belgelendirme", "elden odeme",
    ))):
        tags.add("CEK_KARSILIKSIZ")

    if any(x in full for x in ("cek duzeltme", "duzeltme hakki", "duzeltme ucreti")):
        tags.add("CEK_DUZELTME_HAKKI")

    # ---------------- SENET ----------------
    if "senet" in full and "iade" in full:
        tags.add("SENET_IADE")
    if "senet" in full and "protesto" in full and "kaldir" not in full:
        tags.add("SENET_PROTESTO")
    if "senet" in full and "protesto" in full and "kaldir" in full:
        tags.add("SENET_PROTESTO_KALDIRMA")
    if "senet" in short and any(x in short for x in ("tahsil", "tahsile alma")):
        tags.add("SENET_TAHSIL")

    return tags


@lru_cache(maxsize=16384)
def _channels(row: FeeRow) -> Set[str]:
    """Bir satırın geçerli olduğu kanal kümesini döndürür."""
    mas = _norm(row.masraf)
    full = _norm(row.text)

    explicit = re.search(r"channel\s*=\s*([a-z0-9_]+)", full, flags=re.I)
    if explicit:
        channel = explicit.group(1).upper()
        if channel == "GENEL":
            return {"GENEL"}
        return {channel}

    if "tum kanal" in full:
        return {"MOBIL", "SUBE", "ATM"}

    channels: Set[str] = set()
    mas2 = mas.replace("internet subesi", "internet").replace("internet sube", "internet")

    if any(x in mas2 for x in (
        "mobil", "internet", "dijital", "cepteteb", "iscep", "asistan", "sgk.gov.tr", "web",
    )):
        channels.add("MOBIL")
    if any(x in mas2 for x in (
        "sube", "subeden", "musteri iletisim merkezi", "cozum merkezi", "gise", "kasadan",
    )):
        channels.add("SUBE")
    if any(x in mas2 for x in ("atm", "btm", "kiosk", "bankamatik")):
        channels.add("ATM")

    if not channels:
        full2 = full.replace("internet subesi", "internet").replace("internet sube", "internet")
        if any(x in full2 for x in ("mobil", "internet", "dijital", "iscep", "sgk.gov.tr")):
            channels.add("MOBIL")
        if any(x in full2 for x in (
            "sube", "musteri iletisim merkezi", "cozum merkezi", "gise", "kasadan",
        )):
            channels.add("SUBE")
        if any(x in full2 for x in ("atm", "bankamatik")):
            channels.add("ATM")

    return channels or {"GENEL"}


def _detail_match(row: FeeRow, detail: Optional[str]) -> bool:
    if not detail:
        return True

    text = _norm(row.text)

    tests = {
        "BUYUK": lambda: any(x in text for x in ("buyuk", "large")),
        "ORTA": lambda: any(x in text for x in ("orta", "medium")),
        "KUCUK": lambda: any(x in text for x in ("kucuk", "small")),
        "OZEL": lambda: any(
            x in text
            for x in (
                "ozel",
                "super",
                "extra buyuk",
                "xl",
            )
        ),
        "AYNI": lambda: any(
            x in text
            for x in (
                "ayni banka",
                "ayni sube",
                "bankamiza ait",
                "bankamiz cek",
                "bankamiz ceki",
                "ykb ceki",
                "gb subesinden",
                "kendi bankasi",
            )
        ),
        "DIGER": lambda: any(
            x in text
            for x in (
                "diger banka",
                "diger sube",
                "farkli sube",
                "muhabir banka",
                "muhabirden",
                "baska banka",
                "yurtici banka ceki",
            )
        ),
        "DOVIZ": lambda: any(
            x in text
            for x in (
                "doviz cek",
                "dovizli cek",
                "dovizi natik",
                "yp cek",
                "yabanci banka",
                "yurt disi",
            )
        ),
    }

    fn = tests.get(detail)
    return fn() if fn else True



# ---------------------------------------------------------------------------
# EŞLEŞME PUANI
# ---------------------------------------------------------------------------

def _candidate_score(row: FeeRow, spec: RowSpec, wanted_channel: str) -> int:
    tags = _service_tags(row)

    if spec.service not in tags:
        return -10_000

    score = 100
    mas = _norm(row.masraf)
    cat = _norm(row.kategori)
    full = _norm(row.text)

    # Kaynak önceliği:
    # - Bireysel ana Ürün/Hizmet Ücretleri satırı temel referanstır.
    # - Ek resmî kaynak yalnız ana kaynakta eksik kalan kalemi tamamlar.
    # - Ticari paket/ücret, bireysel karşılaştırmaya taşınmaz.
    is_supplemental = cat.startswith("ek kaynak")
    if not is_supplemental:
        score += 38
    if "ek kaynak - akbank ticari" in cat:
        score -= 55

    if STATUS_NUMERIC.lower() in row.aciklama.lower():
        score += 18
    if STATUS_AVAILABLE.lower() in row.aciklama.lower():
        score -= 55
    if STATUS_EMPTY.lower() in row.aciklama.lower():
        score -= 45
    if STATUS_NOT_APPLICABLE.lower() in row.aciklama.lower():
        score -= 80

    # ---------------- BAND ----------------
    if spec.band_key:
        masraf_keys = _all_band_keys(row.masraf)
        all_keys = _all_band_keys(
            f"{row.masraf} | {row.aciklama}"
        )

        if spec.band_key in all_keys:
            score += 40

            # Band doğrudan MASRAF adındaysa en güvenilir eşleşme.
            if spec.band_key in masraf_keys:
                score += 22

        else:
            # Bazı bankalar Fatura / SGK gibi hizmetleri tutar bandı
            # belirtmeden "tüm tutarlara %X" olarak yayınlıyor.
            # Bu durumda iki ortak banda da aynı tarife uygulanabilir.
            if spec.service in {
                "FATURA",
                "SGK",
                "AIDAT",
                "OZEL_OKUL",
            }:
                score -= 12
            else:
                return -10_000

    # ---------------- DETAY ----------------
    if not _detail_match(row, spec.detail):
        generic_detail_fallback = False

        # Bazı bankalar aynı/diğer banka ayrımı yapmadan tek bir genel
        # çek veya senet tahsil tarifesi yayımlıyor. Böyle bir durumda
        # boş bırakmak yerine aynı genel tarifeyi ilgili ortak satırda
        # düşük öncelikli fallback olarak kullan.
        if spec.service == "SENET_TAHSIL":
            generic_detail_fallback = (
                "senet tahsil" in full
                and not any(
                    x in full
                    for x in (
                        "ayni sube",
                        "diger sube",
                        "muhabir",
                        "baska banka",
                    )
                )
            )

        elif spec.service == "CEK_TAHSIL":
            generic_detail_fallback = (
                (
                    "cek tahsil" in full
                    or ("tahsile alinan" in full and "cek" in full)
                )
                and not any(
                    x in full
                    for x in (
                        "ayni banka",
                        "bankamiz ceki",
                        "ykb ceki",
                        "garanti bankasi",
                        "diger banka",
                        "baska banka",
                        "yabanci banka",
                        "dovizli cek",
                        "dovizi natik",
                    )
                )
            )

        if generic_detail_fallback:
            score -= 18
        else:
            return -10_000

    # ---------------- KANAL ----------------
    if spec.split_channel:
        channels = _channels(row)

        if wanted_channel in channels:
            score += 60
        elif "GENEL" in channels:
            score += 15
        else:
            return -10_000

    # Hizmet MASRAF adında açıkça geçiyorsa açıklama fallback'ından üstündür.
    if spec.service == "EFT" and "eft" in mas:
        score += 30
    elif spec.service == "FAST" and "fast" in mas:
        score += 35
    elif spec.service == "HAVALE" and "havale" in mas:
        score += 30
    elif spec.service in _service_tags(
        FeeRow(
            banka=row.banka,
            kategori=row.kategori,
            masraf=row.masraf,
        )
    ):
        score += 20

    if spec.service in {"SWIFT_GELEN", "SWIFT_GIDEN"}:
        if "swift" in mas:
            score += 35
        if spec.service == "SWIFT_GELEN" and "gelen" in full:
            score += 45
        if spec.service == "SWIFT_GIDEN" and any(x in full for x in ("giden", "gonder")):
            score += 45

    if spec.service == "YURT_DISI_FAST" and any(x in full for x in ("yurt disi fast", "global fast", "fast uluslararasi")):
        score += 65
    if spec.service == "VISA_YP_DIRECT" and "visa" in full:
        score += 65
    if spec.service == "LIMIT_UZERI_PARA_CEKME" and any(x in full for x in ("limit uzeri", "limit ustu")):
        score += 60
    if spec.service == "ORTAK_ATM_PARA_CEKME" and "para cekme" in full:
        score += 45
    if spec.service == "PARA_YATIRMA" and "para yatirma" in full:
        score += 45
    if spec.service == "DUZENLI_EFT" and "duzenli" in full and "eft" in full:
        score += 55
    if spec.service == "DUZENLI_HAVALE" and "duzenli" in full and "havale" in full:
        score += 55
    if spec.service == "ALTIN_TRANSFER" and "altin" in full:
        score += 45

    # Arşiv kategorisi birçok farklı özel raporu içeriyor.
    # "Borcu Yoktur", Risk Raporu, Vize Mektubu gibi alt hizmetlerin
    # Arşiv Araştırma satırına yanlış düşmesini engelle.
    if spec.service == "ARSIV":
        specialized = {
            "BORCU_YOKTUR",
            "KREDI_RISK",
            "CEK_RISK",
            "VIZE_MEKTUBU",
            "REFERANS_MEKTUBU",
            "MEVDUAT_ARASTIRMA",
        }

        if tags & specialized:
            score -= 90

        if any(
            x in mas
            for x in (
                "gecmis donem bankacilik islemleri bildirimi",
                "sozlesme, dekont",
                "dokuman talebi",
                "1 yildan eski",
                "gecmise yonelik",
                "arsiv arastirma",
            )
        ):
            score += 65

    if spec.service == "HESAP_OZETI_POSTA":
        if "kktc" in full:
            score -= 20

        if any(
            x in full
            for x in (
                "posta ile",
                "basili ekstre",
                "ekstre gonderim",
                "hesap ozeti gonderimi",
            )
        ):
            score += 35

    # Standart işlem ücretini özel varyantlardan öne al.
    if any(
        x in mas
        for x in (
            "gonderimi",
            "gonderilmesi",
            "gonderme",
        )
    ):
        score += 18

    if any(
        x in mas
        for x in (
            "gec eft",
            "gec havale",
            "gec fast",
        )
    ):
        score -= 35

    if (
        spec.service in {"EFT", "HAVALE", "FAST", "SWIFT_GELEN", "SWIFT_GIDEN", "YURT_DISI_FAST", "VISA_YP_DIRECT", "DUZENLI_EFT", "DUZENLI_HAVALE"}
        and any(
            x in mas
            for x in (
                "duzenli",
                "talimat",
                "supurme",
            )
        )
    ):
        score -= 12

    if any(
        x in mas
        for x in (
            "cebe",
            "kartsiz",
            "kartli para yatirma",
        )
    ):
        score -= 24

    if any(
        x in cat
        for x in (
            "kampanyali",
            "urun ve hizmet paket",
        )
    ):
        score -= 45

    # Template'teki Fatura satırları kart ödemesi odaklı.
    if spec.service == "FATURA":
        if (
            "kredi kart" in full
            or "kartindan" in full
        ):
            score += 25

        if "nakit" in mas:
            score -= 20

    if (
        spec.service == "SGK"
        and "kredi kart" in full
    ):
        score += 20

    if spec.service == "CEK_RISK":
        if any(x in mas for x in ("cek risk", "cek bilgileri", "kkb cek", "cek sorgu")):
            score += 70

    if spec.service == "KREDI_RISK":
        if "gercek kisiler" in full:
            score += 120
        if "ek kaynak - akbank ticari" in cat or "urun ve hizmet paketleri" in cat:
            score -= 110

    if spec.service == "PARA_YATIRMA":
        if "banka kartiyla" in mas and "cari hesaba para yatirma" in mas:
            score += 85
        if "t.c. subelerinden" in mas:
            score += 35
        if "standart atm" in mas:
            score += 28
        if "k.k.t.c." in mas:
            score -= 55
        if "kredi kartiyla" in mas or "maxipara" in mas:
            score -= 45
        if "ek kaynak - is bankasi banka karti sozlesmesi" in cat:
            score += 125

    if spec.service == "TELEFON":
        if any(x in mas for x in ("tl/paket", "paket yukleme", "otomatik fatura odeme faizi")):
            return -10_000

    if spec.service == "KASA":
        if "depozito" in full:
            score -= 80

        if "aylik" in full:
            score -= 35

        if "yillik" in full:
            score += 45

    return score



def _best_match(
    rows: Sequence[FeeRow],
    bank: str,
    spec: RowSpec,
    wanted_channel: str,
) -> Optional[FeeRow]:
    candidates = []

    for row in rows:
        if row.banka != bank:
            continue

        score = _candidate_score(row, spec, wanted_channel)
        if score <= -10_000:
            continue

        candidates.append((score, row))

    if not candidates:
        return None

    candidates.sort(
        key=lambda item: (
            item[0],
            -len(_norm(item[1].masraf)),
        ),
        reverse=True,
    )

    return candidates[0][1]


# ---------------------------------------------------------------------------
# EXCEL OKUMA / GÖSTERİM
# ---------------------------------------------------------------------------

def _find_source_sheet(wb):
    for ws in wb.worksheets:
        if ws.title == COMPARISON_SHEET:
            continue

        for row_idx in range(1, min(ws.max_row, 20) + 1):
            vals = [
                _norm(ws.cell(row=row_idx, column=col).value)
                for col in range(1, min(ws.max_column, 30) + 1)
            ]
            if "banka" in vals and "masraf" in vals:
                return ws, row_idx

    raise RuntimeError("BANKA ve MASRAF kolonlarını içeren ana veri sayfası bulunamadı.")


def _header_map(ws, header_row: int) -> Dict[str, int]:
    result: Dict[str, int] = {}

    for col in range(1, ws.max_column + 1):
        key = _norm(ws.cell(row=header_row, column=col).value)
        if key in HEADER_ALIASES:
            result[HEADER_ALIASES[key]] = col

    required = {"BANKA", "KATEGORİ", "MASRAF"}
    missing = required - set(result)
    if missing:
        raise RuntimeError("Eksik kolon: " + ", ".join(sorted(missing)))

    return result


def _read_rows(ws, header_row: int, cols: Mapping[str, int]) -> List[FeeRow]:
    def get(row_idx: int, name: str) -> str:
        col = cols.get(name)
        return _clean(ws.cell(row=row_idx, column=col).value) if col else ""

    result: List[FeeRow] = []

    for row_idx in range(header_row + 1, ws.max_row + 1):
        bank = get(row_idx, "BANKA")
        masraf = get(row_idx, "MASRAF")
        if not bank or not masraf:
            continue

        if bank not in BANKS:
            continue

        result.append(FeeRow(
            banka=bank,
            kategori=get(row_idx, "KATEGORİ"),
            masraf=masraf,
            asgari_tutar=get(row_idx, "ASGARİ TUTAR"),
            asgari_oran=get(row_idx, "ASGARİ ORAN"),
            azami_tutar=get(row_idx, "AZAMİ TUTAR"),
            azami_oran=get(row_idx, "AZAMİ ORAN"),
            aciklama=get(row_idx, "AÇIKLAMA"),
            guncelleme_tarihi=get(row_idx, "GÜNCELLEME TARİHİ"),
        ))

    return result


def _percent(value: str) -> str:
    value = _clean(value)
    if not value or value in ("-", "0", "0.0", "0.00"):
        return ""

    if "%" in value:
        return value

    try:
        number = float(value.replace(",", "."))
        # API bazı bankalarda 0,7'yi "0.7", bazılarında 0.007 gibi verebilir.
        if 0 < abs(number) < 0.1:
            number *= 100
        return (f"{number:g}%").replace(".", ",")
    except ValueError:
        return value


def _display_amount(value: str) -> str:
    value = _clean(value)
    if not value or value == "-":
        return ""

    # Birimi standartlaştır. Sayısal nokta ondalıksa Türkçe virgüle çevir.
    value = re.sub(r"\s*TL\b", " TRY", value, flags=re.I)
    value = re.sub(r"\s*TRY\b", " TRY", value, flags=re.I)
    value = re.sub(r"\s*(USD|EUR)\b", r" \1", value, flags=re.I)

    m = re.fullmatch(r"\s*([0-9.,]+)\s*(TRY|USD|EUR)?\s*", value, flags=re.I)
    if not m:
        return value

    raw, currency = m.group(1), (m.group(2) or "").upper()

    # 1234.56 -> 1.234,56 ; 1.234,56 zaten Türkçe formattadır.
    number = _parse_number(raw)
    if number is None:
        return value

    if abs(number - round(number)) < 1e-9:
        formatted = f"{int(round(number)):,}".replace(",", ".")
    else:
        formatted = f"{number:,.2f}"
        formatted = formatted.replace(",", "X").replace(".", ",").replace("X", ".")
        formatted = formatted.rstrip("0").rstrip(",")

    return f"{formatted} {currency}".strip()



def _fee_value_compact(row: Optional[FeeRow]) -> str:
    """Tutar/oranı BSMV notu olmadan tek satırda özetler."""
    if row is None:
        return ""

    min_amount = _display_amount(row.asgari_tutar)
    max_amount = _display_amount(row.azami_tutar)
    min_rate = _percent(row.asgari_oran)
    max_rate = _percent(row.azami_oran)

    if min_amount and max_amount:
        amount = (
            min_amount
            if _norm(min_amount) == _norm(max_amount)
            else f"{min_amount} - {max_amount}"
        )
    else:
        amount = max_amount or min_amount

    if min_rate and max_rate:
        rate = (
            min_rate
            if _norm(min_rate) == _norm(max_rate)
            else f"{min_rate} - {max_rate}"
        )
    else:
        rate = max_rate or min_rate

    if amount and rate:
        return f"{amount} / {rate}"

    return amount or rate


def _extract_deposit_from_description(row: Optional[FeeRow]) -> str:
    """
    Akbank gibi depozitoyu ayrı satır yerine yıllık kira açıklamasında
    yayımlayan bankalar için açıklamadan depozito tutarını çıkarır.
    """
    if row is None:
        return ""

    raw = _clean(row.aciklama)

    patterns = (
        r"depozito\s+bedeli\s*[:\-]?\s*([0-9][0-9.\s]*(?:,[0-9]+)?)\s*(TL|TRY)",
        r"depozito\s+ucreti\s*[:\-]?\s*([0-9][0-9.\s]*(?:,[0-9]+)?)\s*(TL|TRY)",
    )

    norm_raw = _norm(raw)

    # _norm Türkçe karakterleri sadeleştirir fakat rakam biçimini korur.
    for pattern in patterns:
        m = re.search(pattern, norm_raw, flags=re.I)
        if not m:
            continue

        amount = _display_amount(f"{m.group(1).strip()} {m.group(2)}")
        if amount:
            return amount

    return ""


def _deposit_match_score(row: FeeRow, spec: RowSpec) -> int:
    full = _norm(row.text)
    short = _norm(
        f"{row.kategori} | {row.masraf}"
    )

    # Açıklamasında "depozito bedeli ..." geçen yıllık kira satırını
    # ayrı depozito kaydı sanma. Ayrı satır fallback'i yalnız kategori
    # veya MASRAF adında depozito açıkça yazıyorsa çalışır.
    if "depozito" not in short:
        return -10_000

    if "KASA" not in _service_tags(row):
        return -10_000

    score = 100

    if spec.detail == "BUYUK":
        if "buyuk" in full:
            score += 50
        elif "standart kasa depozito" in full:
            score += 15
        elif any(x in full for x in ("orta", "kucuk", "ozel")):
            return -10_000

    elif spec.detail == "ORTA":
        if "orta" in full:
            score += 50
        elif "standart kasa depozito" in full:
            score += 15
        elif any(x in full for x in ("buyuk", "kucuk", "ozel")):
            return -10_000

    elif spec.detail == "KUCUK":
        if "kucuk" in full:
            score += 50
        elif "standart kasa depozito" in full:
            score += 15
        elif any(x in full for x in ("buyuk", "orta", "ozel")):
            return -10_000

    elif spec.detail == "OZEL":
        if "ozel" in full:
            score += 60
        else:
            return -10_000

    # Kasa24 tipleri boy adı olmayan ayrı sınıflar; yanlış boyla
    # eşleştirmemek için standart büyük/orta/küçük satırlarda geriye at.
    if "kasa24" in full and spec.detail in {"BUYUK", "ORTA", "KUCUK"}:
        score -= 40

    return score


def _best_deposit_match(
    rows: Sequence[FeeRow],
    bank: str,
    spec: RowSpec,
    annual_row: Optional[FeeRow],
) -> Tuple[str, str]:
    """(depozito_değeri, BSMV_durumu) döndürür."""
    candidates = []
    for row in rows:
        if row.banka != bank:
            continue
        score = _deposit_match_score(row, spec)
        if score <= -10_000:
            continue
        candidates.append((score, row))

    if candidates:
        candidates.sort(key=lambda item: (item[0], -len(_norm(item[1].masraf))), reverse=True)
        dep_row = candidates[0][1]
        value = _fee_value_compact(dep_row)
        return value, _bsmv_label(dep_row)

    # Akbank gibi depozitoyu yıllık satır açıklamasında veren kaynaklar.
    value = _extract_deposit_from_description(annual_row)
    if not value or annual_row is None:
        return "", ""

    desc = _norm(annual_row.aciklama)
    # Depozito cümlesinin yakınında BSMV bilgisi varsa onu kullan; yoksa tahmin etme.
    m = re.search(r"depozito.{0,160}?(bsmv\s+(?:dahil|haric))", desc, flags=re.I)
    if not m:
        m = re.search(r"(bsmv\s+(?:dahil|haric)).{0,160}?depozito", desc, flags=re.I)
    tax = ""
    if m:
        tax = "BSMV dahil" if "dahil" in _norm(m.group(1)) else "BSMV hariç"
    return value, tax


def _bsmv_label(row: Optional[FeeRow]) -> str:
    if row is None:
        return ""
    desc = _norm(row.aciklama)
    if "bsmv dahil" in desc:
        return "BSMV dahil"
    if "bsmv haric" in desc or "bsmv'den haric" in desc:
        return "BSMV hariç"
    if "bsiv dahil" in desc:
        return "BSİV dahil"
    if "bsiv haric" in desc:
        return "BSİV hariç"
    return ""


def _status_meta(row: Optional[FeeRow]) -> Dict[str, str]:
    if row is None:
        return {}
    desc = row.aciklama or ""
    meta: Dict[str, str] = {}
    for key in ("SERVICE", "CHANNEL", "BAND", "DISPLAY_TEXT"):
        m = re.search(rf"(?:^|[;|]\s*){key}\s*=\s*([^;|]+)", desc, flags=re.I)
        if m:
            meta[key] = _clean(m.group(1)).replace("\\n", "\n")
    return meta


def _status_kind(row: Optional[FeeRow]) -> str:
    if row is None:
        return ""
    desc = row.aciklama or ""
    if STATUS_NOT_APPLICABLE in desc:
        return "NOT_APPLICABLE"
    if STATUS_EMPTY in desc:
        return "PUBLISHED_EMPTY"
    if STATUS_AVAILABLE in desc:
        return "AVAILABLE"
    if STATUS_NUMERIC in desc:
        return "OFFICIAL_FEE"
    return ""


def _has_numeric_fee(row: Optional[FeeRow]) -> bool:
    if row is None:
        return False
    return any(_clean(getattr(row, attr, "")) not in ("", "-", "0", "0.0", "0.00")
               for attr in ("asgari_tutar", "asgari_oran", "azami_tutar", "azami_oran"))


def _fee_text(row: Optional[FeeRow], spec: Optional[RowSpec] = None) -> str:
    if row is None:
        return "N/A"

    status = _status_kind(row)
    meta = _status_meta(row)
    display_text = meta.get("DISPLAY_TEXT", "").strip()
    if display_text and status in {"NOT_APPLICABLE", "PUBLISHED_EMPTY", "AVAILABLE"}:
        return display_text

    if status == "NOT_APPLICABLE" and not _has_numeric_fee(row):
        return "Uygulanmıyor"
    if status == "PUBLISHED_EMPTY" and not _has_numeric_fee(row):
        return "Ücret ilan edilmemiş"
    if status == "AVAILABLE" and not _has_numeric_fee(row):
        return "Hizmet var\nAyrı ücret ilan edilmemiş"

    min_amount = _display_amount(row.asgari_tutar)
    max_amount = _display_amount(row.azami_tutar)
    min_rate = _percent(row.asgari_oran)
    max_rate = _percent(row.azami_oran)

    amount = ""
    if min_amount and max_amount:
        amount = min_amount if _norm(min_amount) == _norm(max_amount) else f"min {min_amount}\nmax {max_amount}"
    else:
        amount = max_amount or min_amount

    rate = ""
    if min_rate and max_rate:
        rate = min_rate if _norm(min_rate) == _norm(max_rate) else f"min {min_rate}\nmax {max_rate}"
    else:
        rate = max_rate or min_rate

    if spec and spec.band_key in {"FATURA_1", "SGK_1"}:
        keys = _all_band_keys(f"{row.masraf} | {row.aciklama}")
        upper_key = "FATURA_2" if spec.band_key == "FATURA_1" else "SGK_2"
        if spec.band_key in keys and upper_key in keys and amount:
            result = amount
        elif amount and rate:
            result = f"{amount}\n{rate}"
        else:
            result = amount or rate
    elif spec and spec.band_key in {"FATURA_2", "SGK_2"}:
        keys = _all_band_keys(f"{row.masraf} | {row.aciklama}")
        lower_key = "FATURA_1" if spec.band_key == "FATURA_2" else "SGK_1"
        if spec.band_key in keys and lower_key in keys and rate:
            result = rate
        else:
            result = rate or amount
    else:
        parts = [x for x in (amount, rate) if x]
        result = "\n".join(parts)

    tax = _bsmv_label(row)
    if tax:
        result = f"{result}\n{tax}" if result else tax
    return result or "Ücret bilgisi açıklamada"



def _service_status_row(
    rows: Sequence[FeeRow],
    bank: str,
    spec: RowSpec,
    wanted_channel: str,
    *,
    kinds: Optional[Set[str]] = None,
) -> Optional[FeeRow]:
    candidates = []
    allowed = kinds or {"AVAILABLE", "PUBLISHED_EMPTY", "NOT_APPLICABLE"}

    for row in rows:
        if row.banka != bank or spec.service not in _service_tags(row):
            continue

        status = _status_kind(row)
        if status not in allowed:
            continue

        meta = _status_meta(row)
        row_band = meta.get("BAND", "").upper()

        # Status belirli bir banda özel ise yalnız o bandı çözsün.
        if row_band and spec.band_key and row_band != spec.band_key:
            continue
        if row_band and not spec.band_key:
            continue

        channels = _channels(row)
        if wanted_channel in channels:
            score = 120
        elif "GENEL" in channels:
            score = 75
        else:
            continue

        if row_band and spec.band_key and row_band == spec.band_key:
            score += 80
        if status == "NOT_APPLICABLE":
            score += 30

        candidates.append((score, row))

    return max(candidates, key=lambda x: x[0])[1] if candidates else None



def _generic_institution_fee(
    rows: Sequence[FeeRow], bank: str, wanted_channel: str,
) -> Optional[FeeRow]:
    """Aidat/okul/telefon için yalnız gerçek Fatura/Kurum tahsilat tarifesini seçer."""
    candidates = []

    for row in rows:
        if row.banka != bank or not _has_numeric_fee(row):
            continue

        full = _norm(row.text)
        mas = _norm(row.masraf)
        cat = _norm(row.kategori)

        if not any(x in full for x in (
            "fatura/kurum", "fatura / kurum", "fatura ve anlasmali kurum",
            "fatura odemeleri", "kurum tahsilat", "kurum odeme",
        )):
            continue

        # Faiz / kredi ürünü / SGK / şans / vergi / telefon yükleme gibi
        # komşu fakat farklı ücretleri genel kurum tarifesine sokma.
        if any(x in mas for x in (
            "faiz", "otomatik fatura odeme", "talimatli fatura odeme islem faizi",
            "alisveris faiz", "sgk", "sans oyun", "vergi", "tl/paket yukleme",
            "paket yukleme", "nakit avans", "konsolosluk", "vize randevu",
        )):
            continue

        channels = _channels(row)
        score = 100

        if wanted_channel in channels:
            score += 70
        elif "GENEL" in channels:
            score += 28
        else:
            continue

        # Bankaların ana, karşılaştırılabilir kurum tarifeleri.
        if bank == "YAPIKREDI" and "fatura ve anlasmali kurum odemeleri" in mas:
            score += 220
        if bank == "AKBANK" and "fatura / kurum tahsil" in mas:
            score += 190
        if bank == "GARANTİ" and "fatura/kurum odemesi" in mas:
            score += 190
        if bank == "İŞBANKASI" and mas == "fatura odemeleri":
            score += 190

        # Kart faiz başlıklarını ve ticari paketleri geriye at.
        if "kredi karti faiz" in cat:
            score -= 200
        if "paket" in mas or "ticari" in cat:
            score -= 90

        candidates.append((score, row))

    if not candidates:
        return None

    candidates.sort(
        key=lambda item: (item[0], -len(_norm(item[1].masraf))),
        reverse=True,
    )
    return candidates[0][1]



def _fast_branch_publication_status(
    rows: Sequence[FeeRow], bank: str, spec: RowSpec, wanted_channel: str,
) -> Optional[str]:
    if spec.service != "FAST" or wanted_channel != "SUBE":
        return None
    mobile = _best_match(rows, bank, spec, "MOBIL")
    if mobile is None:
        return None
    # Bankanın kaynak satırlarında bu band için şube FAST kaydı yoksa bunu eşleşme
    # hatası gibi N/A göstermek yerine kaynakta ayrı tarife bulunmadığını belirt.
    for row in rows:
        if row.banka != bank or "FAST" not in _service_tags(row):
            continue
        if spec.band_key and spec.band_key not in _all_band_keys(f"{row.masraf} | {row.aciklama}"):
            continue
        if "SUBE" in _channels(row):
            return None
    return "Şube FAST tarifesi\nyayımlanmıyor"



def _compact_hint(row: FeeRow) -> str:
    """Ücret hücresinde kısa bant/varyant etiketi üretir."""
    raw = _clean(row.masraf)
    parens = re.findall(r"\(([^()]*(?:TL|TRY|gr|USD|EUR)[^()]*)\)", raw, flags=re.I)
    if parens:
        hint = parens[-1].strip()
        if len(hint) <= 48:
            return hint
    parts = [p.strip() for p in re.split(r"\s+-\s+", raw) if p.strip()]
    if parts:
        tail = parts[-1]
        if len(tail) <= 48:
            return tail
    return ""


def _aggregate_service_fee(
    rows: Sequence[FeeRow], bank: str, spec: RowSpec, wanted_channel: str,
) -> Optional[Tuple[str, FeeRow]]:
    """SWIFT / uluslararası / altın gibi çok satırlı tarifeleri tek hücrede özetler."""
    aggregate_services = {
        "SWIFT_GELEN", "SWIFT_GIDEN", "YURT_DISI_FAST",
        "VISA_YP_DIRECT", "ALTIN_TRANSFER", "CEK_IADE",
    }
    if spec.service not in aggregate_services:
        return None

    lookup_channel = wanted_channel if spec.split_channel else "GENEL"
    candidates = []
    for row in rows:
        if row.banka != bank:
            continue
        score = _candidate_score(row, spec, lookup_channel)
        if score <= -10_000 or not _has_numeric_fee(row):
            continue

        mas = _norm(row.masraf)
        cat = _norm(row.kategori)

        if spec.service in {"SWIFT_GELEN", "SWIFT_GIDEN", "YURT_DISI_FAST"}:
            if any(x in mas for x in ("paket", "kobi")) or "ek kaynak - akbank ticari" in cat:
                continue

        # SWIFT gelen/giden yönünü yalnız MASRAF adındaki açık ifadelerle ayır.
        # Açıklama/kategori içinde "gelen/giden" geçmesi tek başına yeterli değil.
        if spec.service == "SWIFT_GELEN":
            if not any(
                x in mas
                for x in (
                    "gelen swift",
                    "gelen doviz",
                    "yurtdisindan",
                    "yurt disindan",
                    "diger bankalardan gelen",
                    "gelen uluslararasi fon transfer",
                )
            ):
                continue
            if any(x in mas for x in ("giden swift", "gonderim", "gonderilen", "baska bankaya doviz")):
                continue

        if spec.service == "SWIFT_GIDEN":
            if not any(
                x in mas
                for x in (
                    "giden swift",
                    "hesaptan giden swift",
                    "doviz havale gonder",
                    "baska bankaya doviz transfer",
                    "diger bankaya giden",
                    "giden doviz",
                    "uluslararasi fon transferi giden",
                )
            ):
                continue
            if any(x in mas for x in ("gelen swift", "gelen doviz", "yurtdisindan", "yurt disindan")):
                continue

        if spec.service == "ALTIN_TRANSFER":
            if not any(x in mas for x in (
                "altin transfer", "ats ile altin gonderimi",
                "kiymetli maden transferi ucreti - altin",
            )):
                continue
            if any(x in mas for x in ("western union", "eft", "havale", "fast", "fiziki", "teslim")):
                continue

        candidates.append((score, row))

    if not candidates:
        return None

    candidates.sort(key=lambda item: (item[0], -len(_norm(item[1].masraf))), reverse=True)
    blocks = []
    seen = set()
    first_row = candidates[0][1]
    for _, row in candidates:
        fee = _fee_value_compact(row)
        if not fee:
            continue
        hint = _compact_hint(row)
        block = f"{hint}: {fee}" if hint else fee
        tax = _bsmv_label(row)
        if tax:
            block += f" ({tax})"
        key = _norm(block)
        if key in seen:
            continue
        seen.add(key)
        blocks.append(block)
        if len(blocks) >= 5:
            break

    if not blocks:
        return None
    return "\n".join(blocks), first_row


def _kasa24_special_fee(
    rows: Sequence[FeeRow], bank: str, spec: RowSpec
) -> Optional[Tuple[str, FeeRow]]:
    """Yapı Kredi Özel/Süper satırında Kasa24 A-E tarifesini dürüstçe özetler."""
    if bank != "YAPIKREDI" or spec.service != "KASA" or spec.detail != "OZEL":
        return None

    blocks: List[str] = []
    first: Optional[FeeRow] = None

    for kasa_type in ("A", "B", "C", "D", "E"):
        token = _norm(f"{kasa_type} Tipi Kasa24")
        annual = None
        deposit = None
        for row in rows:
            if row.banka != bank:
                continue
            full = _norm(row.text)
            if "kasa24" not in full or token not in full:
                continue
            if "depozito" in full:
                deposit = row
            elif "yillik" in full or "kasa24 ucreti" in full:
                annual = row

        if annual is None:
            continue
        first = first or annual
        annual_fee = _fee_value_compact(annual)
        dep_fee = _fee_value_compact(deposit) if deposit else ""
        line = f"Kasa24 {kasa_type}: Yıllık {annual_fee}"
        if dep_fee:
            line += f" | Depozito {dep_fee}"
        blocks.append(line)

    if not blocks or first is None:
        return None
    return "\n".join(blocks), first


def _source_url(row: Optional[FeeRow]) -> str:
    if row is None:
        return ""
    m = re.search(r"Resmî ek kaynak:\s*(https?://[^\s|]+)", row.aciklama or "", flags=re.I)
    if m:
        return m.group(1).rstrip(".,;")
    return PRIMARY_SOURCE_URLS.get(row.banka, "")


def _cell_comment(row: Optional[FeeRow], resolution: str, value: str) -> Optional[Comment]:
    if row is None:
        return None
    parts = [
        f"Çözüm türü: {resolution}",
        f"Banka: {row.banka}",
        f"Kaynak masraf: {row.masraf}",
    ]
    if row.kategori:
        parts.append(f"Kategori: {row.kategori}")
    if row.guncelleme_tarihi:
        parts.append(f"Kaynak güncelleme: {row.guncelleme_tarihi}")
    url = _source_url(row)
    if url:
        parts.append(f"Kaynak: {url}")
    return Comment("\n".join(parts), "Otomatik Eşleştirme")


def _source_gap_text(spec: RowSpec, bank: str, wanted_channel: str) -> str:
    # Bu ifade "hizmet yok" anlamına gelmez. Yalnız tarife kaynaklarında
    # karşılaştırılabilir ücret bulunamadığını açıkça söyler.
    return "Ayrı karşılaştırılabilir\ntarife doğrulanamadı"


def _resolve_cell(
    rows: Sequence[FeeRow], bank: str, spec: RowSpec, wanted_channel: str,
) -> Tuple[str, Optional[FeeRow], str]:
    """Hücre metni, dayanak satır ve çözüm türünü döndürür."""

    lookup_channel = wanted_channel if spec.split_channel else "GENEL"

    # Resmî olarak uygulanmadığı/kanalda yayımlanmadığı doğrulanan durum,
    # aynı isimli EFT tarife satırının FAST'e yanlış taşınmasını engellemek için
    # sayısal eşleşmeden ÖNCE değerlendirilir.
    pre_status = _service_status_row(
        rows, bank, spec, lookup_channel, kinds={"NOT_APPLICABLE"}
    )
    if pre_status is not None:
        return _fee_text(pre_status, spec), pre_status, "NOT_APPLICABLE"

    # Yapı Kredi Kasa24, klasik Büyük/Orta/Küçük kasa ailesinden farklı
    # yayımlandığı için Özel/Süper satırında A-E birlikte gösterilir.
    kasa24 = _kasa24_special_fee(rows, bank, spec)
    if kasa24 is not None:
        value, source = kasa24
        return value, source, "NUMERIC"

    aggregated = _aggregate_service_fee(rows, bank, spec, wanted_channel)
    if aggregated is not None:
        value, agg_row = aggregated
        return value, agg_row, "NUMERIC"

    row = _best_match(rows, bank, spec, lookup_channel)

    if row is None and not spec.split_channel:
        found = [_best_match(rows, bank, spec, possible) for possible in ("GENEL", "MOBIL", "SUBE")]
        row = next((x for x in found if x is not None), None)

    # Status satırı sayısal ücretin önüne geçmesin; NOT_APPLICABLE yukarıda
    # zaten özel olarak ele alındı.
    if row is not None and _has_numeric_fee(row):
        value = _fee_text(row, spec)
        if spec.service == "KASA":
            annual_tax = _bsmv_label(row)
            annual_fee = _fee_value_compact(row)
            annual = f"Yıllık: {annual_fee}" if annual_fee else "Yıllık: ücret bilgisi açıklamada"
            if annual_tax:
                annual += f" ({annual_tax})"
            dep_value, dep_tax = _best_deposit_match(rows, bank, spec, row)
            value = annual
            if dep_value:
                value += f"\nDepozito: {dep_value}"
                if dep_tax:
                    value += f" ({dep_tax})"
        return value, row, "NUMERIC"

    # Aidat/okul: hizmetin resmî varlık kanıtı varsa, bankanın genel
    # Fatura/Kurum tarifesi açıkça "genel tarife" etiketiyle kullanılabilir.
    if spec.service in {"AIDAT", "OZEL_OKUL", "TELEFON"}:
        status_row = _service_status_row(rows, bank, spec, lookup_channel)
        if status_row is not None:
            generic = _generic_institution_fee(rows, bank, lookup_channel)
            if generic is not None:
                fee = _fee_text(generic, None)
                return f"Genel kurum tarifesi:\n{fee}", generic, "GENERIC_TARIFF"
            return _fee_text(status_row, spec), status_row, "STATUS"

    if row is not None and _status_kind(row):
        return _fee_text(row, spec), row, "STATUS"

    status_row = _service_status_row(rows, bank, spec, lookup_channel)
    if status_row is not None:
        kind = "NOT_APPLICABLE" if _status_kind(status_row) == "NOT_APPLICABLE" else "STATUS"
        return _fee_text(status_row, spec), status_row, kind

    fast_status = _fast_branch_publication_status(rows, bank, spec, wanted_channel)
    if fast_status:
        return fast_status, None, "PUBLICATION_STATUS"

    # Son çare olarak veri uydurmak yerine açıkça kaynak boşluğu belirtilir.
    # Böylece N/A'nın "0 TL" veya "hizmet yok" sanılması engellenir.
    return _source_gap_text(spec, bank, wanted_channel), None, "SOURCE_GAP"



def _preserve_notes(old_ws) -> Dict[Tuple[str, str], str]:
    """Notları satır numarasına değil ORTAK satır adına göre korur."""
    notes: Dict[Tuple[str, str], str] = {}
    if old_ws is None:
        return notes

    section = ""
    for row in range(3, old_ws.max_row + 1):
        label = _clean(old_ws.cell(row=row, column=1).value)
        note = _clean(old_ws.cell(row=row, column=13).value)
        if not label:
            continue

        # Section satırları bankalarda veri taşımadığı için kalın başlık gibi davranır.
        if label in {payload for kind, payload in LAYOUT if kind == "SECTION"}:
            section = label
            continue

        if note:
            notes[(section, label)] = note

    return notes


def _sheet_row_key(section: str, label: str) -> Tuple[str, str]:
    return (section, label)


def _write_comparison(ws, rows: Sequence[FeeRow], notes: Mapping[Tuple[str, str], str]) -> int:
    thin = Side(style="thin", color="B7B7B7")
    medium = Side(style="medium", color="7F7F7F")

    ws["M1"] = "NOTLAR"
    ws.merge_cells("M1:M2")
    ws["M1"].font = Font(bold=True, color="1F1F1F")
    ws["M1"].fill = PatternFill("solid", fgColor="D9EAD3")
    ws["M1"].alignment = Alignment(horizontal="center", vertical="center")

    ws["A2"] = "ORTAK MASRAF ADI"
    ws["A2"].font = Font(bold=True, color="666666")
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")

    col = 2
    for bank in BANKS:
        color = BANK_COLORS[bank]
        ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 1)
        cell = ws.cell(row=1, column=col)
        cell.value = DISPLAY_BANKS[bank]
        cell.fill = PatternFill("solid", fgColor=color)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        for offset, sub in enumerate(("Mobil", "Şube")):
            c = ws.cell(row=2, column=col + offset)
            c.value = sub
            c.fill = PatternFill("solid", fgColor=color)
            c.font = Font(bold=True, italic=True, color="FFFFFF")
            c.alignment = Alignment(horizontal="center", vertical="center")
        col += 2

    current_row = 3
    section = ""

    for kind, payload in LAYOUT:
        if kind == "SECTION":
            section = payload
            ws.cell(row=current_row, column=1).value = payload
            for col_idx in range(1, 10):
                c = ws.cell(row=current_row, column=col_idx)
                c.fill = PatternFill("solid", fgColor="E7E6E6")
                c.font = Font(bold=True, color="595959")
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border = Border(top=medium, bottom=thin)
            ws.row_dimensions[current_row].height = 24
            current_row += 1
            continue

        spec: RowSpec = payload
        ws.cell(row=current_row, column=1).value = spec.label
        ws.cell(row=current_row, column=1).font = Font(color="595959")
        ws.cell(row=current_row, column=1).alignment = Alignment(
            horizontal="left", vertical="center", wrap_text=True
        )

        for bank_index, bank in enumerate(BANKS):
            start_col = 2 + bank_index * 2

            if spec.split_channel:
                for offset, wanted_channel in enumerate(("MOBIL", "SUBE")):
                    value, row, resolution = _resolve_cell(rows, bank, spec, wanted_channel)
                    c = ws.cell(row=current_row, column=start_col + offset)
                    c.value = value
                    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    c.font = Font(
                        bold=resolution not in {"SOURCE_GAP"},
                        color=BANK_COLORS[bank] if resolution != "SOURCE_GAP" else "A6A6A6",
                        size=9,
                    )
                    comment = _cell_comment(row, resolution, value)
                    if comment:
                        c.comment = comment
            else:
                # Kanal ayrımı anlamsız olan kasa, çek, senet, rapor vb. satırlarda
                # Mobil/Şube hücrelerini tek mantıksal banka hücresine birleştir.
                ws.merge_cells(
                    start_row=current_row, start_column=start_col,
                    end_row=current_row, end_column=start_col + 1,
                )
                value, row, resolution = _resolve_cell(rows, bank, spec, "GENEL")
                c = ws.cell(row=current_row, column=start_col)
                c.value = value
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                c.font = Font(
                    bold=resolution not in {"SOURCE_GAP"},
                    color=BANK_COLORS[bank] if resolution != "SOURCE_GAP" else "A6A6A6",
                    size=9,
                )
                comment = _cell_comment(row, resolution, value)
                if comment:
                    c.comment = comment

        note = notes.get(_sheet_row_key(section, spec.label))
        if note:
            ws.cell(row=current_row, column=13).value = note

        if spec.service == "KASA" and spec.detail == "OZEL":
            ws.row_dimensions[current_row].height = 105
        elif spec.service == "KASA":
            ws.row_dimensions[current_row].height = 72
        else:
            ws.row_dimensions[current_row].height = 52
        current_row += 1

    for row_cells in ws.iter_rows(min_row=1, max_row=current_row - 1, min_col=1, max_col=13):
        for cell in row_cells:
            if cell.column <= 9 and cell.row > 2:
                cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.column_dimensions["A"].width = 44
    for col_letter in "BCDEFGHI":
        ws.column_dimensions[col_letter].width = 18
    for col_letter in "JKL":
        ws.column_dimensions[col_letter].width = 3
    ws.column_dimensions["M"].width = 44

    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 22
    ws.freeze_panes = "B3"
    ws.sheet_view.zoomScale = 80
    ws.sheet_view.showGridLines = False
    ws.print_title_rows = "1:2"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # Teknik eşleşme özeti artık görünür Excel alanına yazılmaz.
    # Ayrıntılı kalite kontrolü yalnız GitHub Actions logunda basılır.
    return current_row - 1



def _print_transfer_audit(rows: Sequence[FeeRow]) -> None:
    """Kritik transfer eşleşmelerini GitHub logunda görünür yapar."""
    print("[comparison] ===== ORTAK MASRAF EŞLEŞME KONTROLÜ =====")

    audit_specs = [
        spec
        for kind, spec in LAYOUT
        if kind == "ROW" and spec.service in {
            "EFT", "HAVALE", "FAST", "SWIFT_GELEN", "SWIFT_GIDEN",
            "YURT_DISI_FAST", "VISA_YP_DIRECT", "DUZENLI_EFT",
            "DUZENLI_HAVALE", "ALTIN_TRANSFER",
        }
    ]

    for spec in audit_specs:
        for bank in BANKS:
            channels = ("MOBIL", "SUBE") if spec.split_channel else ("GENEL",)
            for channel in channels:
                value, row, resolution = _resolve_cell(rows, bank, spec, channel)
                if row is None:
                    raw = value.replace("\n", " / ")
                else:
                    raw = row.masraf
                    if len(raw) > 120:
                        raw = raw[:117] + "..."
                print(
                    f"[comparison][match] {spec.service} | {spec.label} | "
                    f"{bank} | {channel} <- {raw} [{resolution}]"
                )

    print("[comparison] ==========================================")



def _is_comparison_sheet_name(title: str) -> bool:
    key = _norm(title).replace(" ", "")
    return (
        key.startswith("karsilastirma")
        or key.startswith("comparison")
    )


def _remove_old_comparison_sheets(wb) -> List[str]:
    """
    Önceki 10-bankalı / 5x5 / test karşılaştırma sayfalarını temizler.
    Böylece kullanıcı yanlışlıkla eski sheet'e bakmaz.
    """
    removed: List[str] = []

    for ws in list(wb.worksheets):
        if _is_comparison_sheet_name(ws.title):
            removed.append(ws.title)
            wb.remove(ws)

    return removed


def _assert_preview_layout(ws) -> None:
    """
    Preview'den sapmayı fatal hata yapar.
    Eski 10 bankalı formatın sessizce geri gelmesini engeller.
    """
    expected_headers = {
        "B1": "Garanti BBVA",
        "D1": "İş Bankası",
        "F1": "Akbank",
        "H1": "Yapı ve Kredi Bankası",
        "A2": "ORTAK MASRAF ADI",
        "M1": "NOTLAR",
    }

    for cell_ref, expected in expected_headers.items():
        actual = _clean(ws[cell_ref].value)
        if actual != expected:
            raise RuntimeError(
                f"Preview layout doğrulaması başarısız: "
                f"{cell_ref} beklenen={expected!r}, gelen={actual!r}"
            )

    # Preview'de banka alanı B:I, J:L boş ve M notlardır.
    for col in ("J", "K", "L"):
        if any(
            _clean(ws[f"{col}{row}"].value)
            for row in range(1, min(ws.max_row, 12) + 1)
        ):
            raise RuntimeError(
                f"Preview layout doğrulaması başarısız: {col} kolonu boş olmalı."
            )

    # Eski formatta kullanılan banka isimleri görünür başlıklarda olmamalı.
    forbidden = {
        "QNB",
        "DenizBank",
        "Halkbank",
        "VakıfBank",
        "TEB",
        "Ziraat",
    }
    visible_headers = {
        _clean(ws.cell(row=row, column=col).value)
        for row in (1, 2)
        for col in range(1, 14)
    }

    bad = sorted(forbidden & visible_headers)
    if bad:
        raise RuntimeError(
            "Preview layout yerine eski çok-bankalı format oluştu: "
            + ", ".join(bad)
        )


def update_comparison_sheet(excel_path: str = "komisyonlar_guncel.xlsx") -> Dict[str, int]:
    path = Path(excel_path)
    if not path.exists():
        raise FileNotFoundError(f"Ana Excel bulunamadı: {path}")

    print(f"[comparison] SÜRÜM: {COMPARISON_VERSION}")
    print(f"[comparison] MODÜL: {Path(__file__).resolve()}")
    print(f"[comparison] LAYOUT: {PREVIEW_LAYOUT_SIGNATURE}")
    print(f"[comparison] Ana Excel okunuyor: {path}")

    wb = load_workbook(path)
    source_ws, header_row = _find_source_sheet(wb)
    cols = _header_map(source_ws, header_row)
    rows = _read_rows(source_ws, header_row, cols)

    if not rows:
        raise RuntimeError("Karşılaştırma için uygun banka verisi bulunamadı.")

    old_ws = wb[COMPARISON_SHEET] if COMPARISON_SHEET in wb.sheetnames else None
    notes = _preserve_notes(old_ws)

    removed_sheets = _remove_old_comparison_sheets(wb)

    ws = wb.create_sheet(COMPARISON_SHEET)
    comparison_rows = _write_comparison(ws, rows, notes)
    # Dosya açıldığında kullanıcı doğrudan karşılaştırma sekmesini görsün.
    wb.active = wb.sheetnames.index(COMPARISON_SHEET)
    ws.sheet_state = "visible"

    # Oluşan sheet preview ile aynı yapıdan saparsa Excel'i kaydetme.
    _assert_preview_layout(ws)

    _print_transfer_audit(rows)

    # Ana Excel'i yarım/bozuk kaydetmemek için atomik karşılaştırma güncellemesi.
    tmp_path = path.with_name(path.stem + ".comparison.tmp" + path.suffix)
    wb.save(tmp_path)
    tmp_path.replace(path)

    print(
        f"[comparison] {COMPARISON_SHEET} referans-sıralı kanonik formatta güncellendi. "
        f"Kaynak={source_ws.title}, 4 banka kaynak satırı={len(rows)}, "
        f"karşılaştırma satırı={comparison_rows}, korunan not={len(notes)}"
    )
    print(
        f"[comparison] Silinen eski karşılaştırma sayfaları: "
        f"{removed_sheets or 'yok'}"
    )
    print(
        "[comparison] Görünüm doğrulandı: "
        "A=ortak masraf, B:I=4 banka Mobil/Şube, J:L=boş, M=NOTLAR."
    )

    resolution_counts = {
        "NUMERIC": 0,
        "GENERIC_TARIFF": 0,
        "STATUS": 0,
        "PUBLICATION_STATUS": 0,
        "NOT_APPLICABLE": 0,
        "SOURCE_GAP": 0,
        "N/A": 0,
    }
    possible_cells = 0

    for kind, payload in LAYOUT:
        if kind != "ROW":
            continue
        spec = payload
        for bank in BANKS:
            channels = ("MOBIL", "SUBE") if spec.split_channel else ("GENEL",)
            for channel in channels:
                possible_cells += 1
                _, _, resolution = _resolve_cell(rows, bank, spec, channel)
                resolution_counts[resolution] = resolution_counts.get(resolution, 0) + 1

    source_gaps = resolution_counts.get("SOURCE_GAP", 0)
    true_na = resolution_counts.get("N/A", 0)
    verified_cells = possible_cells - source_gaps - true_na
    numeric_like = resolution_counts.get("NUMERIC", 0) + resolution_counts.get("GENERIC_TARIFF", 0)
    status_like = (
        resolution_counts.get("STATUS", 0)
        + resolution_counts.get("PUBLICATION_STATUS", 0)
        + resolution_counts.get("NOT_APPLICABLE", 0)
    )

    print(
        f"[comparison] KALİTE: doğrulanmış={verified_cells}/{possible_cells} "
        f"(%{(verified_cells / possible_cells * 100):.1f}) | "
        f"sayısal={resolution_counts.get('NUMERIC', 0)} | "
        f"genel_tarife={resolution_counts.get('GENERIC_TARIFF', 0)} | "
        f"resmî_durum={status_like} | "
        f"kaynak_boşluğu={source_gaps} | N/A={true_na}"
    )


    return {
        "source_rows": len(rows),
        "comparison_rows": comparison_rows,
        "notes_preserved": len(notes),
        "matched_cells": verified_cells,
        "numeric_cells": numeric_like,
        "status_cells": status_like,
        "source_gap_cells": source_gaps,
        "missing_cells": true_na,
        "possible_cells": possible_cells,
    }


if __name__ == "__main__":
    print(update_comparison_sheet())
