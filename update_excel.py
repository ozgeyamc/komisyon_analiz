"""
Çekilen komisyon/ücret verilerini Excel dosyasına yazan/güncelleyen modül.
"""

import os
from datetime import datetime
from typing import Dict, List, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.worksheet import Worksheet

from scraper import UcretSatiri

EXCEL_DOSYA_ADI = "garanti_komisyonlar.xlsx"
SHEET_ADI = "GARANTI"

BASLIKLAR = [
    "KATEGORİ",
    "MASRAF",
    "ASGARİ TUTAR",
    "ASGARİ ORAN",
    "AZAMİ TUTAR",
    "AZAMİ ORAN",
    "AÇIKLAMA",
    "GÜNCELLEME TARİHİ",
]

COL_KATEGORI = 1
COL_MASRAF = 2
COL_ASGARI_TUTAR = 3
COL_ASGARI_ORAN = 4
COL_AZAMI_TUTAR = 5
COL_AZAMI_ORAN = 6
COL_ACIKLAMA = 7
COL_GUNCELLEME_TARIHI = 8

BASLIK_FILL = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
BASLIK_FONT = Font(color="FFFFFF", bold=True)


def _bugun_tarih_str() -> str:
    return datetime.now().strftime("%d.%m.%Y")


def _yeni_workbook_olustur() -> Tuple[Workbook, Worksheet]:
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_ADI

    for idx, baslik in enumerate(BASLIKLAR, start=1):
        cell = ws.cell(row=1, column=idx, value=baslik)
        cell.fill = BASLIK_FILL
        cell.font = BASLIK_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")

    genislikler = [30, 35, 14, 14, 14, 14, 45, 20]
    for idx, genislik in enumerate(genislikler, start=1):
        ws.column_dimensions[ws.cell(row=1, column=idx).column_letter].width = genislik

    return wb, ws


def _workbook_yukle_veya_olustur(dosya_yolu: str) -> Tuple[Workbook, Worksheet]:
    if os.path.exists(dosya_yolu):
        wb = load_workbook(dosya_yolu)
        if SHEET_ADI in wb.sheetnames:
            ws = wb[SHEET_ADI]
        else:
            ws = wb.active
            ws.title = SHEET_ADI
        return wb, ws
    return _yeni_workbook_olustur()


def _mevcut_satirlari_oku(ws: Worksheet) -> Dict[Tuple[str, str], int]:
    eslesme: Dict[Tuple[str, str], int] = {}
    for row_idx in range(2, ws.max_row + 1):
        kategori = ws.cell(row=row_idx, column=COL_KATEGORI).value
        masraf = ws.cell(row=row_idx, column=COL_MASRAF).value
        if kategori is None and masraf is None:
            continue
        eslesme[(str(kategori or ""), str(masraf or ""))] = row_idx
    return eslesme


def _degerler_ayni_mi(ws: Worksheet, row_idx: int, satir: UcretSatiri) -> bool:
    mevcut_asgari_tutar = ws.cell(row=row_idx, column=COL_ASGARI_TUTAR).value or ""
    mevcut_asgari_oran = ws.cell(row=row_idx, column=COL_ASGARI_ORAN).value or ""
    mevcut_azami_tutar = ws.cell(row=row_idx, column=COL_AZAMI_TUTAR).value or ""
    mevcut_azami_oran = ws.cell(row=row_idx, column=COL_AZAMI_ORAN).value or ""

    return (
        str(mevcut_asgari_tutar) == str(satir.asgari_tutar)
        and str(mevcut_asgari_oran) == str(satir.asgari_oran)
        and str(mevcut_azami_tutar) == str(satir.azami_tutar)
        and str(mevcut_azami_oran) == str(satir.azami_oran)
    )


def _satiri_guncelle(ws: Worksheet, row_idx: int, satir: UcretSatiri) -> None:
    ws.cell(row=row_idx, column=COL_ASGARI_TUTAR, value=satir.asgari_tutar)
    ws.cell(row=row_idx, column=COL_ASGARI_ORAN, value=satir.asgari_oran)
    ws.cell(row=row_idx, column=COL_AZAMI_TUTAR, value=satir.azami_tutar)
    ws.cell(row=row_idx, column=COL_AZAMI_ORAN, value=satir.azami_oran)
    ws.cell(row=row_idx, column=COL_ACIKLAMA, value=satir.aciklama)
    ws.cell(row=row_idx, column=COL_GUNCELLEME_TARIHI, value=_bugun_tarih_str())


def _yeni_satir_ekle(ws: Worksheet, satir: UcretSatiri) -> None:
    yeni_row = ws.max_row + 1
    ws.cell(row=yeni_row, column=COL_KATEGORI, value=satir.kategori)
    ws.cell(row=yeni_row, column=COL_MASRAF, value=satir.masraf)
    ws.cell(row=yeni_row, column=COL_ASGARI_TUTAR, value=satir.asgari_tutar)
    ws.cell(row=yeni_row, column=COL_ASGARI_ORAN, value=satir.asgari_oran)
    ws.cell(row=yeni_row, column=COL_AZAMI_TUTAR, value=satir.azami_tutar)
    ws.cell(row=yeni_row, column=COL_AZAMI_ORAN, value=satir.azami_oran)
    ws.cell(row=yeni_row, column=COL_ACIKLAMA, value=satir.aciklama)
    ws.cell(row=yeni_row, column=COL_GUNCELLEME_TARIHI, value=_bugun_tarih_str())


def excel_guncelle(satirlar: List[UcretSatiri], dosya_yolu: str = EXCEL_DOSYA_ADI) -> Dict[str, int]:
    wb, ws = _workbook_yukle_veya_olustur(dosya_yolu)
    mevcut_eslesme = _mevcut_satirlari_oku(ws)

    ozet = {"eklendi": 0, "guncellendi": 0, "degismedi": 0}

    for satir in satirlar:
        anahtar = (satir.kategori, satir.masraf)

        if anahtar in mevcut_eslesme:
            row_idx = mevcut_eslesme[anahtar]
            if _degerler_ayni_mi(ws, row_idx, satir):
                ozet["degismedi"] += 1
                continue
            _satiri_guncelle(ws, row_idx, satir)
            ozet["guncellendi"] += 1
        else:
            _yeni_satir_ekle(ws, satir)
            mevcut_eslesme[anahtar] = ws.max_row
            ozet["eklendi"] += 1

    wb.save(dosya_yolu)
    return ozet
