"""
Çekilen komisyon/ücret verilerini Excel dosyasına yazan/güncelleyen modül.
"""

import os
from datetime import datetime
from typing import Dict, List, Tuple

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.worksheet import Worksheet

from scraper import UcretSatiri

EXCEL_DOSYA_ADI = "garanti_komisyonlar.xlsx"
SHEET_ADI = "GARANTI"

BASLIKLAR = [
    "BANKA",
    "KATEGORİ",
    "MASRAF",
    "ASGARİ TUTAR",
    "ASGARİ ORAN",
    "AZAMİ TUTAR",
    "AZAMİ ORAN",
    "AÇIKLAMA",
    "GÜNCELLEME TARİHİ",
]

COL_BANKA = 1
COL_KATEGORI = 2
COL_MASRAF = 3
COL_ASGARI_TUTAR = 4
COL_ASGARI_ORAN = 5
COL_AZAMI_TUTAR = 6
COL_AZAMI_ORAN = 7
COL_ACIKLAMA = 8
COL_GUNCELLEME_TARIHI = 9

BASLIK_FILL = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
BASLIK_FONT = Font(color="FFFFFF", bold=True)

GARANTI_FILL = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
GARANTI_FONT = Font(color="FFFFFF", bold=True)


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

    genislikler = [15, 30, 35, 14, 14, 14, 14, 60, 20]
    for idx, genislik in enumerate(genislikler, start=1):
        ws.column_dimensions[ws.cell(row=1, column=idx).column_letter].width = genislik

    return wb, ws


def excel_guncelle(satirlar: List[UcretSatiri], dosya_yolu: str = EXCEL_DOSYA_ADI) -> Dict[str, int]:
    if os.path.exists(dosya_yolu):
        os.remove(dosya_yolu)

    wb, ws = _yeni_workbook_olustur()

    for satir in satirlar:
        yeni_row = ws.max_row + 1

        # BANKA kolonu — yeşil renkli
        banka_cell = ws.cell(row=yeni_row, column=COL_BANKA, value="GARANTİ")
        banka_cell.fill = GARANTI_FILL
        banka_cell.font = GARANTI_FONT
        banka_cell.alignment = Alignment(horizontal="center", vertical="center")

        ws.cell(row=yeni_row, column=COL_KATEGORI, value=satir.kategori)
        ws.cell(row=yeni_row, column=COL_MASRAF, value=satir.masraf)
        ws.cell(row=yeni_row, column=COL_ASGARI_TUTAR, value=satir.asgari_tutar)
        ws.cell(row=yeni_row, column=COL_ASGARI_ORAN, value=satir.asgari_oran)
        ws.cell(row=yeni_row, column=COL_AZAMI_TUTAR, value=satir.azami_tutar)
        ws.cell(row=yeni_row, column=COL_AZAMI_ORAN, value=satir.azami_oran)
        ws.cell(row=yeni_row, column=COL_ACIKLAMA, value=satir.aciklama)
        ws.cell(row=yeni_row, column=COL_GUNCELLEME_TARIHI, value=_bugun_tarih_str())

    wb.save(dosya_yolu)

    return {"eklendi": len(satirlar), "guncellendi": 0, "degismedi": 0}
