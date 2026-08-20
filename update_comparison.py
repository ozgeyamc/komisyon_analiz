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
from pathlib import Path
from typing import Dict, Iterable, List, Mapping, Optional, Sequence, Set, Tuple

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


COMPARISON_VERSION = "2026-08-20-v5-preview-locked-canonical"
COMPARISON_SHEET = "KARŞILAŞTIRMA"
PREVIEW_LAYOUT_SIGNATURE = "4BANKS|A:I|J:L_EMPTY|M_NOTES|CANONICAL_BANDS"

# Eski denemelerde oluşmuş karşılaştırma sayfaları kullanıcıyı yanıltmasın.
# v5 çalıştığında KARŞILAŞTIRMA ile başlayan eski sayfaların tamamı silinir
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


@dataclass
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
    ("SECTION", "EFT Gönderimi"),
    ("ROW", RowSpec("0 TRY - 8.300 TRY", "EFT", "TRANSFER_1")),
    ("ROW", RowSpec("8.300,01 TRY - 399.000 TRY", "EFT", "TRANSFER_2")),
    ("ROW", RowSpec("399.000,01 TRY -", "EFT", "TRANSFER_3")),

    ("SECTION", "Havale Gönderimi"),
    ("ROW", RowSpec("0 TRY - 8.300 TRY", "HAVALE", "TRANSFER_1")),
    ("ROW", RowSpec("8.300,01 TRY - 399.000 TRY", "HAVALE", "TRANSFER_2")),
    ("ROW", RowSpec("399.000,01 TRY -", "HAVALE", "TRANSFER_3")),

    ("SECTION", "FAST"),
    ("ROW", RowSpec("0 TRY - 8.300 TRY", "FAST", "TRANSFER_1")),
    ("ROW", RowSpec("8.300,01 TRY - 399.000 TRY", "FAST", "TRANSFER_2")),
    ("ROW", RowSpec("399.000,01 TRY -", "FAST", "TRANSFER_3")),

    ("SECTION", "Kiralık Kasa"),
    ("ROW", RowSpec("Büyük Kasa", "KASA", detail="BUYUK", split_channel=False)),
    ("ROW", RowSpec("Orta Kasa", "KASA", detail="ORTA", split_channel=False)),
    ("ROW", RowSpec("Küçük Kasa", "KASA", detail="KUCUK", split_channel=False)),
    ("ROW", RowSpec("Özel / Süper Kasa", "KASA", detail="OZEL", split_channel=False)),

    ("ROW", RowSpec("Kıymetli Maden Teslimleri -", "KIYMETLI_MADEN_TESLIM")),
    ("ROW", RowSpec(
        "Üçüncü Kişi ve Kuruluşlardan Temin Edilecek Rapor Ücretleri - Kredi Risk Raporu",
        "KREDI_RISK",
        split_channel=False,
    )),

    ("SECTION", "Fatura Ödemeleri - Kart"),
    ("ROW", RowSpec("0 TRY - 149,99 TRY", "FATURA", "FATURA_1")),
    ("ROW", RowSpec("150 TRY -", "FATURA", "FATURA_2")),

    ("SECTION", "SGK Prim Ödemeleri"),
    ("ROW", RowSpec("0 TRY - 99,99 TRY", "SGK", "SGK_1")),
    ("ROW", RowSpec("100 TRY -", "SGK", "SGK_2")),

    ("ROW", RowSpec("HGS Etiket / Kart Bedeli", "HGS", split_channel=False)),
    ("ROW", RowSpec("Şans Oyunu Ödemeleri Aracılık", "SANS_OYUNU")),

    ("SECTION", "Aidat Ödemeleri Aracılık"),
    ("ROW", RowSpec("0 TRY - 149,99 TRY", "AIDAT", "FATURA_1")),
    ("ROW", RowSpec("150 TRY -", "AIDAT", "FATURA_2")),

    ("SECTION", "Özel Okul Ödeme"),
    ("ROW", RowSpec("0 TRY - 149,99 TRY", "OZEL_OKUL", "FATURA_1")),
    ("ROW", RowSpec("150 TRY -", "OZEL_OKUL", "FATURA_2")),

    ("ROW", RowSpec("Telefon Ödemeleri Aracılık", "TELEFON")),
    ("ROW", RowSpec("Vergi Tahsilat Aracılık", "VERGI")),
    ("ROW", RowSpec("Arşiv Araştırma Ücreti", "ARSIV", split_channel=False)),
    ("ROW", RowSpec("Mevduat Araştırma", "MEVDUAT_ARASTIRMA", split_channel=False)),
    ("ROW", RowSpec("Referans Mektubu -", "REFERANS_MEKTUBU", split_channel=False)),
    ("ROW", RowSpec("Vize ve Özel Okullar İçin Düzenlenen Mektuplar -", "VIZE_MEKTUBU", split_channel=False)),
    ("ROW", RowSpec("Hesap Özeti Verilmesi -", "HESAP_OZETI", split_channel=False)),
    ("ROW", RowSpec("Hesap Araştırma Talebi -", "HESAP_ARASTIRMA", split_channel=False)),
    ("ROW", RowSpec("Borcu Yoktur Yazısı", "BORCU_YOKTUR", split_channel=False)),
    ("ROW", RowSpec("Hesap Özeti posta yoluyla", "HESAP_OZETI_POSTA", split_channel=False)),
    ("ROW", RowSpec("Bakiye Sorma - Yurtiçi - ATM", "BAKIYE_ATM_YURTICI", split_channel=False)),
    ("ROW", RowSpec("Bakiye Sorma - Yurtdışı - ATM", "BAKIYE_ATM_YURTDISI", split_channel=False)),

    ("SECTION", "Çek Defteri ve Çek Düzenleme Ücreti"),
    ("ROW", RowSpec("Çek Defteri (Yaprak Başı) -", "CEK_DEFTERI", split_channel=False)),
    ("ROW", RowSpec("Çek Düzenleme -", "CEK_DUZENLEME", split_channel=False)),
    ("ROW", RowSpec("Özel Nitelikli Çek Düzenleme -", "CEK_OZEL", split_channel=False)),
    ("ROW", RowSpec("Çek İade Ücreti", "CEK_IADE", split_channel=False)),

    ("SECTION", "Çek Tahsilat Ücreti"),
    ("ROW", RowSpec("Aynı Banka Çeki -", "CEK_TAHSIL", detail="AYNI", split_channel=False)),
    ("ROW", RowSpec("Diğer Banka Çeki -", "CEK_TAHSIL", detail="DIGER", split_channel=False)),
    ("ROW", RowSpec("Döviz Çekleri Tahsilatı (Diğer Banka) -", "CEK_TAHSIL", detail="DOVIZ", split_channel=False)),

    ("SECTION", "Çek Belgelendirme ve Düzeltme Ücreti"),
    ("ROW", RowSpec("Karşılıksız Çek Belgelendirme -", "CEK_KARSILIKSIZ", split_channel=False)),
    ("ROW", RowSpec("Çek Düzeltme Hakkı -", "CEK_DUZELTME_HAKKI", split_channel=False)),

    ("ROW", RowSpec("Senet İade Ücreti", "SENET_IADE", split_channel=False)),

    ("SECTION", "Senet Protesto İşlemleri Ücreti"),
    ("ROW", RowSpec("Senet Protesto -", "SENET_PROTESTO", split_channel=False)),
    ("ROW", RowSpec("Senet Protesto Kaldırma -", "SENET_PROTESTO_KALDIRMA", split_channel=False)),

    ("SECTION", "Senet Tahsile Alma Ücreti"),
    ("ROW", RowSpec("Aynı Banka Senet Tahsili -", "SENET_TAHSIL", detail="AYNI", split_channel=False)),
    ("ROW", RowSpec("Muhabir Banka Senet Tahsili -", "SENET_TAHSIL", detail="DIGER", split_channel=False)),
]


# ---------------------------------------------------------------------------
# NORMALİZASYON
# ---------------------------------------------------------------------------

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

def _service_tags(row: FeeRow) -> Set[str]:
    cat = _norm(row.kategori)
    mas = _norm(row.masraf)
    text = f"{cat} | {mas}"
    tags: Set[str] = set()

    international = any(x in text for x in (
        "swift",
        "uluslararasi fon transfer",
        "yurt disi fast",
        "global fast",
        "western union",
        "visa ile yurt disi",
    ))

    package = any(x in text for x in ("paket", "kota"))
    card_cash = any(x in text for x in (
        "kredi kartindan",
        "nakit avans",
        "faiz orani",
    ))

    if (
        "eft" in text
        and not international
        and not package
        and not card_cash
        and "altin eft" not in text
    ):
        tags.add("EFT")

    if (
        "fast" in text
        and not package
        and not any(x in text for x in (
            "yurt disi fast",
            "global fast",
            "fast uluslararasi",
        ))
    ):
        tags.add("FAST")

    if (
        "havale" in text
        and not international
        and not package
        and not card_cash
    ):
        tags.add("HAVALE")

    if "kiralik kasa" in text:
        tags.add("KASA")

    if any(x in text for x in (
        "kiymetli maden teslim",
        "fiziki altin teslim",
        "altin teslim",
    )):
        tags.add("KIYMETLI_MADEN_TESLIM")

    if "risk raporu" in text or "kkb risk" in text:
        tags.add("KREDI_RISK")

    if "fatura" in text and "e-fatura" not in text:
        tags.add("FATURA")

    if "sgk" in text:
        tags.add("SGK")

    if "hgs" in text:
        tags.add("HGS")

    if "sans oyun" in text:
        tags.add("SANS_OYUNU")

    if "aidat" in text:
        tags.add("AIDAT")

    if "ozel okul" in text or "okul odeme" in text:
        tags.add("OZEL_OKUL")

    if "telefon" in text and "bankacilik" not in text:
        tags.add("TELEFON")

    if "vergi" in text and "fon transfer" not in text:
        tags.add("VERGI")

    if "arsiv arastirma" in text:
        tags.add("ARSIV")

    if "mevduat arastirma" in text:
        tags.add("MEVDUAT_ARASTIRMA")

    if "referans mektubu" in text:
        tags.add("REFERANS_MEKTUBU")

    if any(x in text for x in (
        "vize icin",
        "konsolosluk icin mektup",
        "ozel okul icin duzenlenen mektup",
    )):
        tags.add("VIZE_MEKTUBU")

    if "hesap ozeti" in text:
        tags.add("HESAP_OZETI")
        if "posta" in text:
            tags.add("HESAP_OZETI_POSTA")

    if "hesap arastirma" in text:
        tags.add("HESAP_ARASTIRMA")

    if "borcu yoktur" in text:
        tags.add("BORCU_YOKTUR")

    if "bakiye" in text and any(x in text for x in ("atm", "bankamatik")):
        if any(x in text for x in ("yurt disi", "yurtdisi")):
            tags.add("BAKIYE_ATM_YURTDISI")
        else:
            tags.add("BAKIYE_ATM_YURTICI")

    if "cek defteri" in text:
        tags.add("CEK_DEFTERI")

    if "cek duzenleme" in text and "ozel" not in text:
        tags.add("CEK_DUZENLEME")

    if "cek duzenleme" in text and "ozel" in text:
        tags.add("CEK_OZEL")

    if "cek iade" in text or "cek iadesi" in text:
        tags.add("CEK_IADE")

    if "cek tahsil" in text:
        tags.add("CEK_TAHSIL")

    if "karsiliksiz cek" in text and any(x in text for x in ("belgelendirme", "islem")):
        tags.add("CEK_KARSILIKSIZ")

    if "cek duzeltme" in text:
        tags.add("CEK_DUZELTME_HAKKI")

    if "senet" in text and "iade" in text:
        tags.add("SENET_IADE")

    if "senet" in text and "protesto" in text and "kaldir" not in text:
        tags.add("SENET_PROTESTO")

    if "senet" in text and "protesto" in text and "kaldir" in text:
        tags.add("SENET_PROTESTO_KALDIRMA")

    if "senet" in text and any(x in text for x in ("tahsil", "tahsile alma")):
        tags.add("SENET_TAHSIL")

    return tags


def _channel(row: FeeRow) -> str:
    """Kanalı MASRAF adına öncelik vererek çıkarır."""
    mas = _norm(row.masraf)
    cat = _norm(row.kategori)

    if any(x in mas for x in ("atm", "btm", "kiosk", "bankamatik")):
        return "ATM"

    # "İnternet Şubesi" fiziksel şube değildir.
    mas2 = (
        mas.replace("internet subesi", "internet")
        .replace("internet sube", "internet")
    )

    if any(x in mas2 for x in (
        "mobil", "internet", "dijital", "cepteteb", "asistan"
    )):
        return "MOBIL"

    if any(x in mas2 for x in (
        "sube", "subeden", "musteri iletisim merkezi",
        "cozum merkezi", "gise", "kasadan"
    )):
        return "SUBE"

    cat2 = (
        cat.replace("internet subesi", "internet")
        .replace("internet sube", "internet")
    )

    if any(x in cat2 for x in ("mobil", "internet", "dijital")):
        return "MOBIL"

    if any(x in cat2 for x in ("sube", "cozum merkezi")):
        return "SUBE"

    return "GENEL"


def _detail_match(row: FeeRow, detail: Optional[str]) -> bool:
    if not detail:
        return True

    text = _norm(row.text)

    tests = {
        "BUYUK": lambda: "buyuk" in text,
        "ORTA": lambda: "orta" in text,
        "KUCUK": lambda: "kucuk" in text,
        "OZEL": lambda: any(x in text for x in ("ozel", "super")),
        "AYNI": lambda: any(x in text for x in ("ayni banka", "bankamiza ait", "bankamiz cek")),
        "DIGER": lambda: any(x in text for x in ("diger banka", "muhabir banka", "baska banka")),
        "DOVIZ": lambda: any(x in text for x in ("doviz cek", "yp cek")),
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

    if spec.band_key:
        key = _band_key(_parse_band(row.masraf))
        if key != spec.band_key:
            return -10_000
        score += 40

    if not _detail_match(row, spec.detail):
        return -10_000

    if spec.split_channel:
        ch = _channel(row)
        if ch == wanted_channel:
            score += 60
        elif ch == "GENEL":
            score += 10
        else:
            return -10_000

    # Standart işlem ücretini; geç/düzenli/paket gibi özel varyantların önüne al.
    if any(x in mas for x in ("gonderimi", "gonderilmesi", "gonderme")):
        score += 18

    if spec.service == "EFT" and "eft" in mas:
        score += 15

    if spec.service == "FAST" and "fast" in mas:
        score += 25

    if spec.service == "HAVALE" and "havale" in mas:
        score += 15

    if "gec eft" in mas or " - gec - " in mas:
        score -= 35

    if any(x in mas for x in ("duzenli", "talimat", "supurme")):
        score -= 14

    if any(x in mas for x in ("cebe", "kartsiz", "kartli para yatirma")):
        score -= 24

    if any(x in cat for x in ("kredi kart", "kampanyali", "urun ve hizmet paket")):
        score -= 45

    if spec.service == "KASA":
        # Karşılaştırmada kira bedeli göster; depozito ve aylık ücret kira satırını
        # ele geçirmesin. İş Bankası gibi hem aylık hem yıllık yayınlayan bankalarda
        # yıllık satır önceliklidir.
        if "depozito" in mas or "depozito" in cat:
            score -= 80
        if "aylik" in mas or "aylik" in cat:
            score -= 35
        if "yillik" in mas or "yillik" in cat:
            score += 45
        if "yillik kira" in mas or "yillik kasa" in mas:
            score += 20

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


def _fee_text(row: Optional[FeeRow]) -> str:
    if row is None:
        return "N/A"

    min_amount = _display_amount(row.asgari_tutar)
    max_amount = _display_amount(row.azami_tutar)
    min_rate = _percent(row.asgari_oran)
    max_rate = _percent(row.azami_oran)

    parts: List[str] = []

    if min_amount and min_amount != "-" and max_amount and max_amount != "-":
        if _norm(min_amount) == _norm(max_amount):
            parts.append(min_amount)
        else:
            parts.append(
                f"min {min_amount}\n"
                f"max {max_amount}"
            )
    elif max_amount and max_amount != "-":
        parts.append(max_amount)
    elif min_amount and min_amount != "-":
        parts.append(min_amount)

    if min_rate and max_rate:
        if _norm(min_rate) == _norm(max_rate):
            parts.append(min_rate)
        else:
            parts.append(f"min {min_rate}\nmax {max_rate}")
    elif max_rate:
        parts.append(max_rate)
    elif min_rate:
        parts.append(min_rate)

    aciklama = _norm(row.aciklama)
    if "bsmv dahil" in aciklama:
        parts.append("BSMV dahil")
    elif "bsmv haric" in aciklama:
        parts.append("BSMV hariç")

    return "\n".join(parts) if parts else "Ücret bilgisi açıklamada"


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

    # Şablon: A = ortak masraf adı, B:I = 4 banka x Mobil/Şube, M = NOTLAR.
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
    matched = 0
    missing = 0

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

        col = 2
        for bank in BANKS:
            for wanted_channel in ("MOBIL", "SUBE"):
                lookup_channel = wanted_channel if spec.split_channel else "GENEL"
                row = _best_match(rows, bank, spec, lookup_channel)

                # Kanal kırılımı olmayan ortak hizmetlerde aynı ücret iki kolonda görünür.
                if row is None and not spec.split_channel:
                    # Bir kanal belirtilmişse bile hizmetin en iyi kaydını al.
                    candidates = []
                    for possible in ("MOBIL", "SUBE", "GENEL"):
                        found = _best_match(rows, bank, spec, possible)
                        if found is not None:
                            candidates.append(found)
                    row = candidates[0] if candidates else None

                value = _fee_text(row)
                c = ws.cell(row=current_row, column=col)
                c.value = value
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                c.font = Font(
                    bold=value != "N/A",
                    color=BANK_COLORS[bank] if value != "N/A" else "A6A6A6",
                    size=9,
                )

                if row is not None:
                    matched += 1
                else:
                    missing += 1

                col += 1

        note = notes.get(_sheet_row_key(section, spec.label))
        if note:
            ws.cell(row=current_row, column=13).value = note

        ws.row_dimensions[current_row].height = 48
        current_row += 1

    # Stil / ölçüler.
    for row in ws.iter_rows(min_row=1, max_row=current_row - 1, min_col=1, max_col=13):
        for cell in row:
            if cell.column <= 9 and cell.row > 2:
                cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.column_dimensions["A"].width = 44
    for col_letter in "BCDEFGHI":
        ws.column_dimensions[col_letter].width = 18
    ws.column_dimensions["J"].width = 3
    ws.column_dimensions["K"].width = 3
    ws.column_dimensions["L"].width = 3
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

    # Alt bilgi.
    ws.cell(row=current_row + 1, column=1).value = (
        f"Otomatik eşleştirme | {COMPARISON_VERSION} | "
        f"{datetime.now().strftime('%d.%m.%Y %H:%M')} | "
        f"eşleşen hücre={matched}, N/A={missing}"
    )
    ws.cell(row=current_row + 1, column=1).font = Font(
        italic=True, color="808080", size=8
    )

    return current_row - 1


def _print_transfer_audit(rows: Sequence[FeeRow]) -> None:
    """Özellikle EFT/Havale/FAST eşleştirmesini GitHub logunda görünür yapar."""
    print("[comparison] ===== ORTAK MASRAF EŞLEŞME KONTROLÜ =====")

    audit_specs = [
        spec
        for kind, spec in LAYOUT
        if kind == "ROW" and spec.service in {"EFT", "HAVALE", "FAST"}
    ]

    for spec in audit_specs:
        for bank in BANKS:
            for channel in ("MOBIL", "SUBE"):
                row = _best_match(rows, bank, spec, channel)
                if row is None:
                    raw = "N/A"
                else:
                    raw = row.masraf
                    if len(raw) > 120:
                        raw = raw[:117] + "..."

                print(
                    f"[comparison][match] {spec.service} | {spec.label} | "
                    f"{bank} | {channel} <- {raw}"
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

    # Oluşan sheet preview ile aynı yapıdan saparsa Excel'i kaydetme.
    _assert_preview_layout(ws)

    _print_transfer_audit(rows)

    wb.save(path)

    print(
        f"[comparison] {COMPARISON_SHEET} PREVIEW-LOCK formatında güncellendi. "
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

    return {
        "source_rows": len(rows),
        "comparison_rows": comparison_rows,
        "notes_preserved": len(notes),
    }


if __name__ == "__main__":
    print(update_comparison_sheet())
