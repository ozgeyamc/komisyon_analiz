"""
komisyonlar_guncel.xlsx içindeki ana komisyon verisinden otomatik
"KARŞILAŞTIRMA" sayfası üretir.

Amaç:
- Ana scraper Excel'i güncellendiğinde bu sayfa da aynı koşuda yenilensin.
- Kullanıcının paylaştığı "banka komisyonları karşılaştırma.xlsx" görünümüne
  benzer, bankaları yan yana ve Mobil / Şube kırılımında gösteren bir format.
- Kaynak veriyi kopyalamaz; her çalışmada ana veri tablosundan yeniden üretir.
- KARŞILAŞTIRMA sayfasındaki NOTLAR sütununda elle yazılmış notları korur.

Yeni dependency gerektirmez; proje zaten Excel yazmak için openpyxl kullanıyor.
"""

import re
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


COMPARISON_VERSION = "2026-08-20-v1-auto-comparison"
COMPARISON_SHEET = "KARŞILAŞTIRMA"

BANKS = [
    "GARANTİ",
    "YAPIKREDI",
    "İŞBANKASI",
    "AKBANK",
    "QNB",
    "DENİZBANK",
    "HALKBANK",
    "VAKIFBANK",
    "TEB",
    "ZİRAAT",
]

DISPLAY_BANKS = {
    "GARANTİ": "Garanti BBVA",
    "YAPIKREDI": "Yapı Kredi",
    "İŞBANKASI": "İş Bankası",
    "AKBANK": "Akbank",
    "QNB": "QNB",
    "DENİZBANK": "DenizBank",
    "HALKBANK": "Halkbank",
    "VAKIFBANK": "VakıfBank",
    "TEB": "TEB",
    "ZİRAAT": "Ziraat",
}

# Her bankanın header rengi ve aynı bankaya ait veri font rengi.
BANK_COLORS = {
    "GARANTİ": "70AD47",
    "YAPIKREDI": "17365D",
    "İŞBANKASI": "2E75B6",
    "AKBANK": "FF1F1F",
    "QNB": "7030A0",
    "DENİZBANK": "00A6A6",
    "HALKBANK": "0070C0",
    "VAKIFBANK": "F4B183",
    "TEB": "00A651",
    "ZİRAAT": "C00000",
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
    "site guncelleme tarihi": "SİTE GÜNCELLEME TARİHİ",
    "son kontrol": "SON KONTROL",
}


@dataclass
class FeeRow:
    banka: str
    kategori: str
    masraf: str
    asgari_tutar: str
    asgari_oran: str
    azami_tutar: str
    azami_oran: str
    aciklama: str
    site_tarihi: str

    @property
    def text(self) -> str:
        return " | ".join(
            part
            for part in [
                self.kategori,
                self.masraf,
                self.aciklama,
            ]
            if part
        )


@dataclass(frozen=True)
class ComparisonItem:
    label: str
    include_any: Tuple[str, ...]
    exclude_any: Tuple[str, ...] = ()
    category_any: Tuple[str, ...] = ()
    mode: str = "summary"  # summary | tier
    tier_index: int = 0
    general_channel: bool = False


# Satır şablonu.
# ("SECTION", başlık) yeni bir bölüm açar.
# ("ITEM", ComparisonItem(...)) veri satırıdır.
LAYOUT = [
    ("SECTION", "EFT Gönderimi"),
    ("ITEM", ComparisonItem("1. Kademe", ("eft",), ("fast", "swift", "altin eft", "altın eft"), ("para aktarma",), "tier", 0)),
    ("ITEM", ComparisonItem("2. Kademe", ("eft",), ("fast", "swift", "altin eft", "altın eft"), ("para aktarma",), "tier", 1)),
    ("ITEM", ComparisonItem("3. Kademe", ("eft",), ("fast", "swift", "altin eft", "altın eft"), ("para aktarma",), "tier", 2)),

    ("SECTION", "Havale Gönderimi"),
    ("ITEM", ComparisonItem("1. Kademe", ("havale",), ("swift", "kredi kart"), ("para aktarma",), "tier", 0)),
    ("ITEM", ComparisonItem("2. Kademe", ("havale",), ("swift", "kredi kart"), ("para aktarma",), "tier", 1)),
    ("ITEM", ComparisonItem("3. Kademe", ("havale",), ("swift", "kredi kart"), ("para aktarma",), "tier", 2)),

    ("SECTION", "FAST"),
    ("ITEM", ComparisonItem("1. Kademe", ("fast",), ("uluslararasi", "uluslararası"), ("para aktarma",), "tier", 0)),
    ("ITEM", ComparisonItem("2. Kademe", ("fast",), ("uluslararasi", "uluslararası"), ("para aktarma",), "tier", 1)),
    ("ITEM", ComparisonItem("3. Kademe", ("fast",), ("uluslararasi", "uluslararası"), ("para aktarma",), "tier", 2)),

    ("SECTION", "SWIFT / Uluslararası Fon Transferi"),
    ("ITEM", ComparisonItem("Giden", ("swift", "uluslararasi fon transfer", "uluslararası fon transfer"), ("sorgulama", "mesajlasma", "mesajlaşma"), (), "summary", 0)),
    ("ITEM", ComparisonItem("Gelen", ("swift", "uluslararasi fon transfer", "uluslararası fon transfer"), ("giden", "gonderil", "gönderil"), (), "summary", 0)),

    ("SECTION", "Kiralık Kasa"),
    ("ITEM", ComparisonItem("Büyük Kasa", ("kiralik kasa", "kiralık kasa"), (), (), "summary", 0, True)),
    ("ITEM", ComparisonItem("Orta Kasa", ("kiralik kasa", "kiralık kasa"), (), (), "summary", 0, True)),
    ("ITEM", ComparisonItem("Küçük Kasa", ("kiralik kasa", "kiralık kasa"), (), (), "summary", 0, True)),
    ("ITEM", ComparisonItem("Özel / Süper Kasa", ("kiralik kasa", "kiralık kasa"), (), (), "summary", 0, True)),
    ("ITEM", ComparisonItem("Depozito", ("depozito",), (), (), "summary", 0, True)),

    ("SECTION", "ATM / Ortak ATM"),
    ("ITEM", ComparisonItem("Para Çekme", ("para cekme", "para çekme"), ("kredi", "nakit avans"), ("atm",), "summary", 0, True)),
    ("ITEM", ComparisonItem("Ortak ATM Para Çekme", ("ortak", "diger banka atm", "diğer banka atm"), (), ("atm",), "summary", 0, True)),
    ("ITEM", ComparisonItem("Bakiye Sorma", ("bakiye",), (), ("atm",), "summary", 0, True)),

    ("SECTION", "Tahsilat / Ödeme"),
    ("ITEM", ComparisonItem("Fatura Ödemeleri", ("fatura",), ("e-fatura", "e arşiv", "e-arsiv"), (), "summary", 0, True)),
    ("ITEM", ComparisonItem("SGK Prim Ödemeleri", ("sgk",), (), (), "summary", 0, True)),
    ("ITEM", ComparisonItem("HGS", ("hgs",), (), (), "summary", 0, True)),
    ("ITEM", ComparisonItem("Şans Oyunu Ödemeleri", ("sans oyunu", "şans oyunu"), (), (), "summary", 0, True)),
    ("ITEM", ComparisonItem("Aidat Ödemeleri", ("aidat",), (), (), "summary", 0, True)),
    ("ITEM", ComparisonItem("Telefon Ödemeleri", ("telefon",), (), (), "summary", 0, True)),
    ("ITEM", ComparisonItem("Vergi Tahsilatı", ("vergi",), (), (), "summary", 0, True)),

    ("SECTION", "Rapor / Belge"),
    ("ITEM", ComparisonItem("Kredi Risk Raporu", ("risk raporu", "kkb risk"), (), (), "summary", 0, True)),
    ("ITEM", ComparisonItem("Arşiv Araştırma", ("arsiv arastirma", "arşiv araştırma"), (), (), "summary", 0, True)),
    ("ITEM", ComparisonItem("Mevduat Araştırma", ("mevduat arastirma", "mevduat araştırma"), (), (), "summary", 0, True)),
    ("ITEM", ComparisonItem("Referans Mektubu", ("referans mektubu",), (), (), "summary", 0, True)),
    ("ITEM", ComparisonItem("Borcu Yoktur Yazısı", ("borcu yoktur",), (), (), "summary", 0, True)),

    ("SECTION", "Kıymetli Maden"),
    ("ITEM", ComparisonItem("Kıymetli Maden Teslimi", ("kiymetli maden teslim", "kıymetli maden teslim"), (), (), "summary", 0, True)),

    ("SECTION", "Çek / Senet"),
    ("ITEM", ComparisonItem("Çek Defteri / Yaprak Başı", ("cek defteri", "çek defteri"), (), (), "summary", 0, True)),
    ("ITEM", ComparisonItem("Çek İade", ("cek iade", "çek iade"), (), (), "summary", 0, True)),
    ("ITEM", ComparisonItem("Çek Tahsilat", ("cek tahsil", "çek tahsil"), (), (), "summary", 0, True)),
    ("ITEM", ComparisonItem("Senet İade", ("senet iade",), (), (), "summary", 0, True)),
    ("ITEM", ComparisonItem("Senet Protesto", ("senet protesto",), (), (), "summary", 0, True)),
    ("ITEM", ComparisonItem("Senet Tahsil", ("senet tahsil",), (), (), "summary", 0, True)),
]


def _norm(value) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\u0307", "")
    text = text.replace("\xa0", " ")
    text = " ".join(text.split())
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    trans = str.maketrans({
        "ı": "i",
        "İ": "i",
        "ğ": "g",
        "Ğ": "g",
        "ü": "u",
        "Ü": "u",
        "ş": "s",
        "Ş": "s",
        "ö": "o",
        "Ö": "o",
        "ç": "c",
        "Ç": "c",
    })
    return text.translate(trans).lower().strip()


def _cell_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float):
        # Excel'de 0.035 gibi oranlar ham kaynaktan gelebilir.
        if 0 < abs(value) < 1:
            return f"%{value * 100:.2f}".replace(".", ",")
        if value.is_integer():
            return str(int(value))
    return str(value).strip()


def _find_source_sheet(wb):
    """
    BANKA + MASRAF header'larını içeren sheet'i otomatik bulur.
    KARŞILAŞTIRMA sheet'i kaynak olarak seçilmez.
    """
    for ws in wb.worksheets:
        if ws.title == COMPARISON_SHEET:
            continue

        max_scan_row = min(ws.max_row, 20)
        max_scan_col = min(ws.max_column, 30)

        for row_idx in range(1, max_scan_row + 1):
            vals = [
                _norm(ws.cell(row=row_idx, column=col_idx).value)
                for col_idx in range(1, max_scan_col + 1)
            ]

            if "banka" in vals and "masraf" in vals:
                return ws, row_idx

    raise RuntimeError(
        "Ana veri sayfası bulunamadı. BANKA ve MASRAF başlıklarını içeren "
        "bir sheet bekleniyordu."
    )


def _build_header_map(ws, header_row: int) -> Dict[str, int]:
    result = {}

    for col_idx in range(1, ws.max_column + 1):
        raw = _norm(ws.cell(row=header_row, column=col_idx).value)

        if raw in HEADER_ALIASES:
            result[HEADER_ALIASES[raw]] = col_idx

    required = {"BANKA", "KATEGORİ", "MASRAF"}

    missing = required - set(result)

    if missing:
        raise RuntimeError(
            "Ana Excel'de zorunlu kolonlar eksik: "
            + ", ".join(sorted(missing))
        )

    return result


def _load_rows(ws, header_row: int, colmap: Dict[str, int]) -> List[FeeRow]:
    rows: List[FeeRow] = []

    def get(row_idx: int, name: str) -> str:
        col = colmap.get(name)
        if not col:
            return ""
        return _cell_text(ws.cell(row=row_idx, column=col).value)

    for row_idx in range(header_row + 1, ws.max_row + 1):
        banka = get(row_idx, "BANKA")
        masraf = get(row_idx, "MASRAF")

        if not banka or not masraf:
            continue

        rows.append(
            FeeRow(
                banka=banka,
                kategori=get(row_idx, "KATEGORİ"),
                masraf=masraf,
                asgari_tutar=get(row_idx, "ASGARİ TUTAR"),
                asgari_oran=get(row_idx, "ASGARİ ORAN"),
                azami_tutar=get(row_idx, "AZAMİ TUTAR"),
                azami_oran=get(row_idx, "AZAMİ ORAN"),
                aciklama=get(row_idx, "AÇIKLAMA"),
                site_tarihi=get(row_idx, "SİTE GÜNCELLEME TARİHİ"),
            )
        )

    return rows


def _bank_match(row: FeeRow, bank: str) -> bool:
    bank_norm = _norm(row.banka)

    aliases = {
        "GARANTİ": ("garanti",),
        "YAPIKREDI": ("yapikredi", "yapi kredi", "yapi ve kredi"),
        "İŞBANKASI": ("isbank", "is bankasi", "turkiye is bankasi"),
        "AKBANK": ("akbank",),
        "QNB": ("qnb",),
        "DENİZBANK": ("denizbank", "deniz bank"),
        "HALKBANK": ("halkbank", "halk bank"),
        "VAKIFBANK": ("vakifbank", "vakif bank"),
        "TEB": ("teb", "turk ekonomi bankasi"),
        "ZİRAAT": ("ziraat", "ziraat bankasi"),
    }

    return any(alias in bank_norm for alias in aliases[bank])


def _channel(row: FeeRow) -> str:
    text = _norm(row.text)

    mobile_tokens = (
        "mobil",
        "internet",
        "dijital",
        "cepteteb",
        "cep sube",
        "cep şube",
    )

    branch_tokens = (
        "sube",
        "şube",
        "gise",
        "gişe",
        "musteri iletisim merkezi",
        "müşteri iletişim merkezi",
        "cagri merkezi",
        "çağrı merkezi",
    )

    if "tum kanallar" in text or "tüm kanallar" in row.text.lower():
        return "ALL"

    has_mobile = any(_norm(token) in text for token in mobile_tokens)
    has_branch = any(_norm(token) in text for token in branch_tokens)

    if has_mobile and not has_branch:
        return "MOBIL"

    if has_branch and not has_mobile:
        return "SUBE"

    if has_mobile and has_branch:
        return "ALL"

    return "GENEL"


def _matches(row: FeeRow, item: ComparisonItem) -> bool:
    text = _norm(row.text)
    category = _norm(row.kategori)

    if item.include_any:
        if not any(_norm(term) in text for term in item.include_any):
            return False

    if item.exclude_any:
        if any(_norm(term) in text for term in item.exclude_any):
            return False

    if item.category_any:
        if not any(_norm(term) in category for term in item.category_any):
            return False

    # Kasa bedenleri için label'a göre ek filtre.
    label = _norm(item.label)

    if "buyuk kasa" in label and "buyuk" not in text:
        return False

    if "orta kasa" in label and "orta" not in text:
        return False

    if "kucuk kasa" in label and "kucuk" not in text:
        return False

    if "ozel / super kasa" in label:
        if not any(term in text for term in ("ozel", "super")):
            return False

    if label == "depozito" and "depozito" not in text:
        return False

    return True


_AMOUNT_RE = re.compile(
    r"(?<!\d)(\d{1,3}(?:[.\s]\d{3})*(?:,\d+)?|\d+(?:,\d+)?)"
)


def _first_amount(text: str) -> float:
    """
    MASRAF adındaki ilk parasal/tutar benzeri sayıyı sıralama için kullanır.
    Bulamazsa çok büyük değer döndürür.
    """
    for raw in _AMOUNT_RE.findall(text):
        candidate = raw.replace(" ", "")

        if "." in candidate and "," in candidate:
            candidate = candidate.replace(".", "").replace(",", ".")
        elif "," in candidate:
            candidate = candidate.replace(",", ".")
        else:
            # 8.300 gibi Türkçe binlik ayracı.
            parts = candidate.split(".")
            if len(parts) > 1 and all(len(p) == 3 for p in parts[1:]):
                candidate = "".join(parts)

        try:
            value = float(candidate)
            # Tarih/yıl gibi yüksek ve anlamsız ilk sayıların önüne geçmek için.
            return value
        except ValueError:
            continue

    return float("inf")


def _fee_value(row: FeeRow) -> str:
    parts = []

    # Tutar özeti.
    min_amt = row.asgari_tutar
    max_amt = row.azami_tutar

    if min_amt and max_amt and _norm(min_amt) != _norm(max_amt):
        parts.append(f"{min_amt} - {max_amt}")
    elif max_amt:
        parts.append(max_amt)
    elif min_amt:
        parts.append(min_amt)

    # Oran özeti.
    min_rate = row.asgari_oran
    max_rate = row.azami_oran

    def rate_text(value: str) -> str:
        value = value.strip()
        if not value:
            return ""
        if "%" in value:
            return value
        return f"%{value}"

    if min_rate and max_rate and _norm(min_rate) != _norm(max_rate):
        parts.append(f"{rate_text(min_rate)} - {rate_text(max_rate)}")
    elif max_rate:
        parts.append(rate_text(max_rate))
    elif min_rate:
        parts.append(rate_text(min_rate))

    if not parts:
        parts.append("Ücret bilgisi açıklamada")

    note = _norm(row.aciklama)

    if row.aciklama and any(
        token in note
        for token in (
            "bsmv",
            "dahil",
            "haric",
            "hariç",
            "vergi",
            "ucretsiz",
            "ücretsiz",
        )
    ):
        short_note = row.aciklama.strip()
        if len(short_note) > 90:
            short_note = short_note[:87] + "..."
        parts.append(short_note)

    return "\n".join(parts)


def _row_range_label(row: FeeRow) -> str:
    """
    MASRAF adındaki son kısa kırılımı cell içinde göstermeye çalışır.
    """
    text = row.masraf.strip()

    if " - " in text:
        tail = text.split(" - ")[-1].strip()
        if tail:
            return tail

    return text


def _dedupe_rows(rows: Iterable[FeeRow]) -> List[FeeRow]:
    result = []
    seen = set()

    for row in rows:
        key = (
            _norm(row.masraf),
            _norm(row.asgari_tutar),
            _norm(row.asgari_oran),
            _norm(row.azami_tutar),
            _norm(row.azami_oran),
            _norm(row.aciklama),
        )

        if key in seen:
            continue

        seen.add(key)
        result.append(row)

    return result


def _select_rows(
    all_rows: Sequence[FeeRow],
    bank: str,
    item: ComparisonItem,
    channel: str,
) -> List[FeeRow]:
    rows = [
        row
        for row in all_rows
        if _bank_match(row, bank)
        and _matches(row, item)
    ]

    if not rows:
        return []

    if item.general_channel:
        return _dedupe_rows(rows)

    exact = [
        row
        for row in rows
        if _channel(row) == channel
    ]

    all_channel = [
        row
        for row in rows
        if _channel(row) == "ALL"
    ]

    generic = [
        row
        for row in rows
        if _channel(row) == "GENEL"
    ]

    if exact:
        chosen = exact + all_channel
    elif all_channel:
        chosen = all_channel
    else:
        chosen = generic

    return _dedupe_rows(chosen)


def _render_item(
    rows: Sequence[FeeRow],
    item: ComparisonItem,
) -> str:
    if not rows:
        return "N/A"

    sorted_rows = sorted(
        rows,
        key=lambda row: (
            _first_amount(row.masraf),
            _norm(row.masraf),
        ),
    )

    if item.mode == "tier":
        # Aynı bandın farklı tekrarlarını filtreleyip benzersiz sıralı tarifeler.
        unique = []
        seen = set()

        for row in sorted_rows:
            key = (
                _norm(_row_range_label(row)),
                _norm(_fee_value(row)),
            )

            if key in seen:
                continue

            seen.add(key)
            unique.append(row)

        if item.tier_index >= len(unique):
            return "N/A"

        row = unique[item.tier_index]
        range_label = _row_range_label(row)
        fee = _fee_value(row)

        return f"{range_label}\n{fee}".strip()

    # Summary hücresinde en fazla 4 farklı ücret göster.
    lines = []

    for row in sorted_rows:
        name = _row_range_label(row)
        fee = _fee_value(row)

        # Tek satırlık/generic masraf isimlerinde label'ı tekrar etmeyelim.
        if len(sorted_rows) == 1:
            block = fee
        else:
            block = f"{name}: {fee}"

        if block not in lines:
            lines.append(block)

        if len(lines) >= 4:
            break

    return "\n".join(lines) if lines else "N/A"


def _preserve_notes(ws) -> Dict[int, str]:
    """
    Aynı layout satırları korunduğu sürece NOTLAR sütunundaki manuel notları
    satır numarasına göre taşır.
    """
    notes = {}

    if ws is None:
        return notes

    note_col = None

    for col in range(1, ws.max_column + 1):
        top = _norm(ws.cell(row=1, column=col).value)
        second = _norm(ws.cell(row=2, column=col).value)

        if "notlar" in {top, second}:
            note_col = col
            break

    if note_col is None:
        return notes

    for row in range(1, ws.max_row + 1):
        value = ws.cell(row=row, column=note_col).value
        if value not in (None, ""):
            notes[row] = str(value)

    return notes


def _style_sheet(ws, notes: Dict[int, str], all_rows: Sequence[FeeRow]) -> int:
    thin = Side(style="thin", color="7F7F7F")
    medium = Side(style="medium", color="5B5B5B")

    # 1 label + 10 bank * 2 channel + 1 notes
    note_col = 2 + len(BANKS) * 2
    last_bank_col = note_col - 1

    # Row 1-2 header.
    ws.merge_cells(start_row=1, start_column=note_col, end_row=2, end_column=note_col)
    notes_cell = ws.cell(row=1, column=note_col)
    notes_cell.value = "NOTLAR"
    notes_cell.font = Font(bold=True, color="1F1F1F")
    notes_cell.alignment = Alignment(horizontal="center", vertical="center")
    notes_cell.fill = PatternFill("solid", fgColor="D9EAD3")

    ws.cell(row=2, column=1).value = "ÜCRET / HİZMET"
    ws.cell(row=2, column=1).font = Font(bold=True, color="666666")
    ws.cell(row=2, column=1).alignment = Alignment(horizontal="center")

    col = 2

    for bank in BANKS:
        color = BANK_COLORS[bank]

        ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 1)

        c = ws.cell(row=1, column=col)
        c.value = DISPLAY_BANKS[bank]
        c.fill = PatternFill("solid", fgColor=color)
        c.font = Font(bold=True, color="FFFFFF", size=11)
        c.alignment = Alignment(horizontal="center", vertical="center")

        for offset, label in enumerate(("Mobil", "Şube")):
            sub = ws.cell(row=2, column=col + offset)
            sub.value = label
            sub.fill = PatternFill("solid", fgColor=color)
            sub.font = Font(bold=True, italic=True, color="FFFFFF")
            sub.alignment = Alignment(horizontal="center", vertical="center")

        col += 2

    current_row = 3

    for kind, payload in LAYOUT:
        if kind == "SECTION":
            ws.cell(row=current_row, column=1).value = payload

            for cidx in range(1, last_bank_col + 1):
                cell = ws.cell(row=current_row, column=cidx)
                cell.fill = PatternFill("solid", fgColor="E7E6E6")
                cell.font = Font(bold=True, color="595959")
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = Border(top=medium, bottom=thin)

            ws.row_dimensions[current_row].height = 24

        else:
            item: ComparisonItem = payload
            ws.cell(row=current_row, column=1).value = item.label
            ws.cell(row=current_row, column=1).font = Font(color="666666")
            ws.cell(row=current_row, column=1).alignment = Alignment(
                horizontal="left",
                vertical="center",
                wrap_text=True,
            )

            col = 2

            for bank in BANKS:
                bank_color = BANK_COLORS[bank]

                for channel in ("MOBIL", "SUBE"):
                    selected = _select_rows(
                        all_rows,
                        bank,
                        item,
                        channel,
                    )

                    value = _render_item(
                        selected,
                        item,
                    )

                    cell = ws.cell(row=current_row, column=col)
                    cell.value = value
                    cell.font = Font(
                        bold=value != "N/A",
                        color=bank_color if value != "N/A" else "A6A6A6",
                        size=9,
                    )
                    cell.alignment = Alignment(
                        horizontal="center",
                        vertical="center",
                        wrap_text=True,
                    )

                    col += 1

            ws.row_dimensions[current_row].height = 56

        # Notes restore.
        if current_row in notes:
            ws.cell(row=current_row, column=note_col).value = notes[current_row]

        current_row += 1

    # Timestamp / kaynak bilgisi.
    ws.cell(row=current_row + 1, column=1).value = (
        f"Otomatik üretildi: {datetime.now().strftime('%d.%m.%Y %H:%M')} | "
        f"Kaynak ücret satırı: {len(all_rows)} | Sürüm: {COMPARISON_VERSION}"
    )
    ws.cell(row=current_row + 1, column=1).font = Font(
        italic=True,
        color="808080",
        size=8,
    )

    # Borders + notes style.
    for row in ws.iter_rows(
        min_row=1,
        max_row=current_row - 1,
        min_col=1,
        max_col=note_col,
    ):
        for cell in row:
            if cell.row not in (1, 2) and cell.column <= last_bank_col:
                cell.border = Border(
                    left=thin,
                    right=thin,
                    top=thin,
                    bottom=thin,
                )

    for row_idx in range(3, current_row):
        c = ws.cell(row=row_idx, column=note_col)
        c.alignment = Alignment(
            horizontal="left",
            vertical="top",
            wrap_text=True,
        )
        c.font = Font(size=9, color="595959")

    # Column widths.
    ws.column_dimensions["A"].width = 31

    for col_idx in range(2, last_bank_col + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 18

    ws.column_dimensions[get_column_letter(note_col)].width = 42

    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 22

    ws.freeze_panes = "B3"
    ws.sheet_view.zoomScale = 75
    ws.sheet_view.showGridLines = False

    # Print / filter-like usability.
    ws.auto_filter.ref = f"A2:{get_column_letter(note_col)}{current_row - 1}"
    ws.print_title_rows = "1:2"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    return current_row - 1


def update_comparison_sheet(
    excel_path: str = "komisyonlar_guncel.xlsx",
) -> Dict[str, int]:
    path = Path(excel_path)

    if not path.exists():
        raise FileNotFoundError(
            f"Karşılaştırma üretilemedi; ana Excel bulunamadı: {path}"
        )

    print(f"[comparison] SÜRÜM: {COMPARISON_VERSION}")
    print(f"[comparison] Ana Excel okunuyor: {path}")

    wb = load_workbook(path)

    source_ws, header_row = _find_source_sheet(wb)
    colmap = _build_header_map(source_ws, header_row)
    all_rows = _load_rows(source_ws, header_row, colmap)

    if not all_rows:
        raise RuntimeError(
            "Karşılaştırma üretilemedi; ana veri sayfasında ücret satırı yok."
        )

    old_ws = wb[COMPARISON_SHEET] if COMPARISON_SHEET in wb.sheetnames else None
    notes = _preserve_notes(old_ws)

    if old_ws is not None:
        wb.remove(old_ws)

    ws = wb.create_sheet(COMPARISON_SHEET)

    generated_rows = _style_sheet(
        ws,
        notes,
        all_rows,
    )

    wb.save(path)

    print(
        f"[comparison] {COMPARISON_SHEET} sayfası güncellendi. "
        f"Kaynak={source_ws.title}, ücret_satırı={len(all_rows)}, "
        f"karşılaştırma_satırı={generated_rows}"
    )

    return {
        "source_rows": len(all_rows),
        "comparison_rows": generated_rows,
        "notes_preserved": len(notes),
    }


if __name__ == "__main__":
    result = update_comparison_sheet()
    print(result)
