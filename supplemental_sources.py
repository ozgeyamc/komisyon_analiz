"""
Resmî ikincil kaynaklardan ana scraper'ların kaçırabildiği ücret/hizmetleri tamamlar.

Amaç
-----
Ana 10 scraper dondurulmuş halde kalır. Bu modül, yalnızca ana "Ürün ve Hizmet
Ücretleri" sayfasında görünmeyen fakat bankanın başka bir resmî sayfasında / PDF'inde
yayımlanan verileri ekler.

Kaynaklar
---------
- İş Bankası: Sözleşme ve Formlar -> en güncel Bankacılık Hizmetleri Sözleşmesi PDF'i
  (özellikle Senet / Çek / Çek Defteri tarifeleri)
- Akbank: Ticari Müşterilerden Alınabilecek Ücretler ve Alt Kalemler
  (tüm parse edilebilir ücret satırları; boş yayımlanan alt kalemler de durum satırı)
- Akbank: Ödeme Merkezi + Düzenli Ödemeler
  (Aidat / Eğitim hizmetinin varlığı ve kanalları)
- Yapı Kredi: Düzenli Ödemeler
  (Aidat / okul taksiti hizmetinin varlığı)
- Garanti BBVA: Özel Okul Ödemeleri
  (özel okul hizmetinin varlığı)

Bu modül ücret uydurmaz. Resmî tabloda ücret alanı boşsa ücret kolonları boş kalır ve
AÇIKLAMA içine [PUBLISHED_EMPTY] durumu yazılır. Yalnız hizmet varlığı doğrulanmışsa
[AVAILABLE_NO_SEPARATE_FEE] yazılır. Karşılaştırma modülü bu durumları kullanıcıya
anlaşılır metinle gösterir.
"""

from __future__ import annotations

import io
import re
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict, Iterable, List, Mapping, MutableMapping, Optional, Sequence, Tuple
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup


SUPPLEMENTAL_VERSION = "2026-08-20-v6-official-channel-and-limit-audit"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "tr-TR,tr;q=0.9,en;q=0.8",
}

ISBANK_CONTRACTS_URL = "https://www.isbank.com.tr/sozlesme-ve-formlar"
ISBANK_BHS_FALLBACK_URL = "https://www.isbank.com.tr/Documents/BHS%20%282026-01.%29.pdf"
ISBANK_SCHOOL_URL = "https://www.isbank.com.tr/ozel-okul-odemeleri"
ISBANK_SITE_AIDAT_URL = "https://www.isbank.com.tr/apartman-yonetim-ve-site-tahsilat-sistemi"
ISBANK_TAX_URL = "https://www.isbank.com.tr/vergi-odeme"
ISBANK_FINDEKS_URL = "https://www.isbank.com.tr/is-ticari/findeks-hizmetleri"
ISBANK_BILL_URL = "https://www.isbank.com.tr/fatura-odemeleri"

ISBANK_FAST_URL = "https://www.isbank.com.tr/fast-anlik-para-transferi"
ISBANK_SGK_URL = "https://www.isbank.com.tr/sgk-odemeleri"
ISBANK_CARD_CURRENT_URL = "https://www.isbank.com.tr/Documents/KKR%20S%C3%B6zle%C5%9Fmesi%20isbank.com.tr%28%20KREDI-KARTI-SOZ%20%2912.08.2026.pdf"

AKBANK_COMMERCIAL_URL = "https://www.akbank.com/ticari-musterilerden-alinabilecek-ucretler-ve-alt-kalemler"
AKBANK_PAYMENT_CENTER_URL = "https://www.akbank.com/odeme-merkezi"
AKBANK_REGULAR_URL = "https://www.akbank.com/odeme-para-transferi/odemeler/duzenli-odemeler"
AKBANK_TAX_URL = "https://www.akbank.com/odeme-para-transferi/yasal-odemeler/vergi-odemeleri"

AKBANK_FAST_URL = "https://www.akbank.com/odeme-para-transferi/para-transferleri/fast-ve-kolay-adres-tanimlama"
AKBANK_GAME_URL = "https://www.akbank.com/odeme-para-transferi/odemeler/sans-oyunu-odemeler"
AKBANK_FEE_URL = "https://www.akbank.com/urun-ve-hizmet-ucretleri"

YAPIKREDI_REGULAR_URL = "https://www.yapikredi.com.tr/bireysel-bankacilik/odemeler-ve-hizmetler/duzenli-odemeler"
YAPIKREDI_BILL_URL = "https://www.yapikredi.com.tr/odemeler-ve-hizmetler/otomatik-fatura-odeme-talimati"
YAPIKREDI_MIM_URL = "https://www.yapikredi.com.tr/kendim-icin/sinirsiz-bankacilik/iletisim-ve-yardim/musteri-iletisim-merkezi/"
YAPIKREDI_FAST_URL = "https://www.yapikredi.com.tr/bireysel-bankacilik/odemeler-ve-hizmetler/fonlarin-anlik-transferi"
GARANTI_SCHOOL_URL = "https://www.garantibbva.com.tr/odemeler-ve-hizmetler/ozel-okul-odemeleri"
GARANTI_FAST_URL = "https://www.garantibbva.com.tr/odemeler-ve-hizmetler/fast-kolay-adres"
GARANTI_BILL_URL = "https://www.garantibbva.com.tr/odemeler-ve-hizmetler/fatura-odeme"
GARANTI_TAX_URL = "https://www.garantibbva.com.tr/odemeler-ve-hizmetler/vergi-odemeleri"
GARANTI_FEE_URL = "https://www.garantibbva.com.tr/urun-ve-hizmet-ucretleri"

STATUS_AVAILABLE = "[SUPPLEMENTAL][AVAILABLE_NO_SEPARATE_FEE]"
STATUS_EMPTY = "[SUPPLEMENTAL][PUBLISHED_EMPTY]"
STATUS_NUMERIC = "[SUPPLEMENTAL][OFFICIAL_FEE]"
STATUS_NOT_APPLICABLE = "[SUPPLEMENTAL][NOT_APPLICABLE]"


@dataclass
class SupplementalRow:
    kategori: str
    masraf: str
    asgari_tutar: str = ""
    asgari_oran: str = ""
    azami_tutar: str = ""
    azami_oran: str = ""
    aciklama: str = ""
    site_guncelleme_tarihi: str = ""


@dataclass
class SupplementalReport:
    ok: bool = True
    version: str = SUPPLEMENTAL_VERSION
    added_by_bank: Dict[str, int] = field(default_factory=dict)
    source_counts: Dict[str, int] = field(default_factory=dict)
    sources: Dict[str, str] = field(default_factory=dict)
    warnings: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)

    @property
    def total_added(self) -> int:
        return sum(self.added_by_bank.values())

    def fail(self, message: str) -> None:
        self.errors.append(message)
        self.ok = False


class SupplementalSourceError(RuntimeError):
    pass


def _norm(value) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\xa0", " ").replace("\u200b", " ")
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
    return " ".join(str(value).replace("\xa0", " ").split()).strip()


def _request(url: str, *, binary: bool = False, timeout: int = 45):
    response = requests.get(url, headers=HEADERS, timeout=timeout)
    response.raise_for_status()
    return response.content if binary else response.text

def _fetch_html(url: str, *, must_contain: Sequence[str] = (), timeout: int = 45) -> str:
    """Önce requests, gerekirse tek sayfalık Playwright fallback."""
    html = ""
    request_error = None
    try:
        html = _request(url, timeout=timeout)
        page_text = _norm(BeautifulSoup(html, "html.parser").get_text(" ", strip=True))
        if not must_contain or all(_norm(term) in page_text for term in must_contain):
            return html
    except Exception as exc:
        request_error = exc

    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=True)
            page = browser.new_page()
            page.goto(url, wait_until="domcontentloaded", timeout=90_000)
            try:
                page.wait_for_load_state("networkidle", timeout=20_000)
            except Exception:
                pass
            html = page.content()
            browser.close()
        page_text = _norm(BeautifulSoup(html, "html.parser").get_text(" ", strip=True))
        missing = [term for term in must_contain if _norm(term) not in page_text]
        if missing:
            raise SupplementalSourceError(
                f"{url}: beklenen içerik bulunamadı: {', '.join(missing)}"
            )
        return html
    except Exception as browser_exc:
        if request_error:
            raise SupplementalSourceError(
                f"{url}: requests={request_error}; playwright={browser_exc}"
            ) from browser_exc
        raise


def _source_note(marker: str, url: str, extra: str = "") -> str:
    note = f"{marker} Resmî ek kaynak: {url}"
    if extra:
        note += f" | {extra}"
    return note


def _row_signature(row) -> Tuple[str, ...]:
    return (
        _norm(getattr(row, "kategori", "")),
        _norm(getattr(row, "masraf", "")),
        _norm(getattr(row, "asgari_tutar", "")),
        _norm(getattr(row, "asgari_oran", "")),
        _norm(getattr(row, "azami_tutar", "")),
        _norm(getattr(row, "azami_oran", "")),
    )


def _semantic_signature(row) -> Tuple[str, ...]:
    """Kategori farklı olsa bile aynı masraf+tarifeyi duplicate kabul eder."""
    return (
        _norm(getattr(row, "masraf", "")),
        _norm(getattr(row, "asgari_tutar", "")),
        _norm(getattr(row, "asgari_oran", "")),
        _norm(getattr(row, "azami_tutar", "")),
        _norm(getattr(row, "azami_oran", "")),
    )


def _append_unique(existing: List, additions: Iterable[SupplementalRow]) -> int:
    exact = {_row_signature(row) for row in existing}
    semantic = {_semantic_signature(row) for row in existing}
    added = 0

    for row in additions:
        sig = _row_signature(row)
        sem = _semantic_signature(row)

        # Ücretli bir satır zaten ana scraper'da aynı ad/tarife ile varsa tekrar ekleme.
        has_fee = any(
            _clean(getattr(row, attr, "")) not in ("", "-", "0", "0.0", "0.00")
            for attr in ("asgari_tutar", "asgari_oran", "azami_tutar", "azami_oran")
        )

        if sig in exact or (has_fee and sem in semantic):
            continue

        existing.append(row)
        exact.add(sig)
        semantic.add(sem)
        added += 1

    return added


def _money(value: str) -> str:
    value = _clean(value)
    if not value or value in ("-", "–", "—"):
        return ""
    return value


def _rate(value: str) -> str:
    value = _clean(value)
    if not value or value in ("-", "–", "—"):
        return ""
    return value.replace("%", "").strip()


def _heading_before(table) -> str:
    heading = table.find_previous(["h2", "h3", "h4", "h5", "button"])
    return _clean(heading.get_text(" ", strip=True)) if heading else "Ticari Ücretler"


def _parse_akbank_commercial_html(html: str) -> List[SupplementalRow]:
    """Akbank ticari ücret sayfasındaki tüm standart HTML tablolarını toplar."""
    soup = BeautifulSoup(html, "html.parser")
    rows: List[SupplementalRow] = []

    for table in soup.find_all("table"):
        trs = table.find_all("tr")
        if not trs:
            continue

        # Header'ı bul.
        header_idx = None
        header_map: Dict[str, int] = {}
        for i, tr in enumerate(trs[:4]):
            cells = [_clean(c.get_text(" ", strip=True)) for c in tr.find_all(["th", "td"])]
            norms = [_norm(c) for c in cells]
            if any("asgari tutar" in n for n in norms) or any("azami tutar" in n for n in norms):
                header_idx = i
                for col, name in enumerate(norms):
                    if "masraf" in name or "urun" in name or "islem" in name:
                        header_map.setdefault("masraf", col)
                    elif "asgari tutar" in name:
                        header_map["min_amount"] = col
                    elif "asgari oran" in name:
                        header_map["min_rate"] = col
                    elif "azami tutar" in name:
                        header_map["max_amount"] = col
                    elif "azami oran" in name:
                        header_map["max_rate"] = col
                    elif "aciklama" in name:
                        header_map["desc"] = col
                    elif "guncelleme" in name or "tarih" in name:
                        header_map["date"] = col
                break

        if header_idx is None:
            continue

        if "masraf" not in header_map:
            header_map["masraf"] = 0

        heading = _heading_before(table)

        for tr in trs[header_idx + 1:]:
            cells = [_clean(c.get_text(" ", strip=True)) for c in tr.find_all(["td", "th"])]
            if not cells:
                continue

            def cell(key: str) -> str:
                idx = header_map.get(key)
                return cells[idx] if idx is not None and idx < len(cells) else ""

            masraf = cell("masraf")
            if not masraf or _norm(masraf) in {"masraf", "islem", "urun"}:
                continue

            min_amount = _money(cell("min_amount"))
            min_rate = _rate(cell("min_rate"))
            max_amount = _money(cell("max_amount"))
            max_rate = _rate(cell("max_rate"))
            desc = cell("desc")
            date = cell("date")

            has_fee = any((min_amount, min_rate, max_amount, max_rate))
            marker = STATUS_NUMERIC if has_fee else STATUS_EMPTY
            status_extra = "" if has_fee else "Resmî tabloda hizmet satırı mevcut; ücret alanları boş."

            rows.append(SupplementalRow(
                kategori=f"EK KAYNAK - Akbank Ticari - {heading}",
                masraf=masraf,
                asgari_tutar=min_amount,
                asgari_oran=min_rate,
                azami_tutar=max_amount,
                azami_oran=max_rate,
                aciklama=(desc + " | " if desc else "") + _source_note(marker, AKBANK_COMMERCIAL_URL, status_extra),
                site_guncelleme_tarihi=date,
            ))

    # HTML yapısı değişse bile kullanıcı için kritik iki boş alt kalemi kaybetme.
    page_text = _clean(soup.get_text(" ", strip=True))
    critical = [
        ("Karşılıksız Çek Belgelendirme", "CEK_KARSILIKSIZ"),
        ("Çek Düzeltme Hakkı", "CEK_DUZELTME_HAKKI"),
    ]
    for label, service in critical:
        if _norm(label) in _norm(page_text) and not any(_norm(label) in _norm(r.masraf) for r in rows):
            rows.append(SupplementalRow(
                kategori="EK KAYNAK - Akbank Ticari - Çek Belgelendirme ve Düzeltme",
                masraf=label,
                aciklama=_source_note(
                    STATUS_EMPTY,
                    AKBANK_COMMERCIAL_URL,
                    f"SERVICE={service}; resmî sayfada alt kalem yayımlanıyor ancak ücret alanları boş.",
                ),
            ))

    return rows


def _add_service_status(
    bank: str,
    service: str,
    label: str,
    channels: Sequence[str],
    url: str,
    evidence: str,
    *,
    marker: str = STATUS_AVAILABLE,
    display_text: str = "",
    band: str = "",
) -> List[SupplementalRow]:
    result = []
    for channel in channels:
        channel_label = {
            "MOBIL": "Mobil/İnternet",
            "SUBE": "Şube/Müşteri İletişim Merkezi",
            "ATM": "ATM",
            "GENEL": "Genel",
        }.get(channel, channel)
        extra = f"SERVICE={service}; CHANNEL={channel}; {evidence}"
        if band:
            extra += f"; BAND={band}"
        if display_text:
            # Pipe source-note ayıracı olduğu için DISPLAY_TEXT pipe içermez.
            extra += f"; DISPLAY_TEXT={display_text.replace('|', '/')}"
        result.append(SupplementalRow(
            kategori="EK KAYNAK - Hizmet Durumu",
            masraf=f"{label} - {channel_label}",
            aciklama=_source_note(marker, url, extra),
        ))
    return result


def _akbank_service_rows() -> List[SupplementalRow]:
    center = _fetch_html(AKBANK_PAYMENT_CENTER_URL, must_contain=("Site Aidat Ödemeleri",))
    regular = _fetch_html(AKBANK_REGULAR_URL, must_contain=("okul taksiti", "apartman aidatı"))
    n_center = _norm(BeautifulSoup(center, "html.parser").get_text(" ", strip=True))
    n_regular = _norm(BeautifulSoup(regular, "html.parser").get_text(" ", strip=True))

    if "egitim odemeleri" not in n_center:
        raise SupplementalSourceError("Akbank Ödeme Merkezi'nde 'Eğitim Ödemeleri' bulunamadı.")
    if "site aidat odemeleri" not in n_center and "uyelik aidat odemeleri" not in n_center:
        raise SupplementalSourceError("Akbank Ödeme Merkezi'nde aidat hizmeti bulunamadı.")
    if "okul taksiti" not in n_regular or "apartman aidati" not in n_regular:
        raise SupplementalSourceError("Akbank Düzenli Ödemeler sayfasındaki okul/aidat doğrulaması bulunamadı.")

    rows = []
    rows += _add_service_status(
        "AKBANK", "AIDAT", "Aidat Ödemeleri", ("MOBIL", "SUBE"), AKBANK_REGULAR_URL,
        "Düzenli Ödemeler sayfasında apartman aidatı açıkça sayılıyor; Mobil/İnternet, Şube ve Müşteri İletişim Merkezi kanalları belirtiliyor.",
    )
    rows += _add_service_status(
        "AKBANK", "OZEL_OKUL", "Özel Okul / Eğitim Ödemeleri", ("MOBIL", "SUBE"), AKBANK_REGULAR_URL,
        "Düzenli Ödemeler sayfasında okul taksiti açıkça sayılıyor; Ödeme Merkezi'nde Eğitim Ödemeleri ayrıca listeleniyor.",
    )
    return rows


def _yk_service_rows() -> List[SupplementalRow]:
    html = _fetch_html(YAPIKREDI_REGULAR_URL, must_contain=("aidat", "okul taksiti"))
    text = _norm(BeautifulSoup(html, "html.parser").get_text(" ", strip=True))
    if "aidat" not in text or "okul taksiti" not in text:
        raise SupplementalSourceError("Yapı Kredi Düzenli Ödemeler sayfasında aidat/okul taksiti bulunamadı.")

    rows = []
    rows += _add_service_status(
        "YAPIKREDI", "AIDAT", "Aidat Ödemeleri", ("MOBIL", "SUBE"), YAPIKREDI_REGULAR_URL,
        "Düzenli Ödemeler sayfasında aidat açıkça listeleniyor; Mobil/İnternet üzerinden oluşturma, ayrıca Müşteri İletişim Merkezi ve şubeden talimat kanalı belirtiliyor.",
    )
    rows += _add_service_status(
        "YAPIKREDI", "OZEL_OKUL", "Özel Okul / Okul Taksiti", ("MOBIL", "SUBE"), YAPIKREDI_REGULAR_URL,
        "Düzenli Ödemeler sayfasında okul taksiti açıkça listeleniyor; Mobil/İnternet üzerinden oluşturma, ayrıca Müşteri İletişim Merkezi ve şubeden talimat kanalı belirtiliyor.",
    )
    return rows


def _garanti_service_rows() -> List[SupplementalRow]:
    html = _fetch_html(GARANTI_SCHOOL_URL, must_contain=("özel okul",))
    text = _norm(BeautifulSoup(html, "html.parser").get_text(" ", strip=True))
    if "ozel okul" not in text:
        raise SupplementalSourceError("Garanti BBVA özel okul sayfası doğrulanamadı.")
    # Sayfa açıkça şubelerden özel okul ödemesi yapılabildiğini belirtiyor.
    return _add_service_status(
        "GARANTİ", "OZEL_OKUL", "Özel Okul Ödemeleri", ("SUBE",), GARANTI_SCHOOL_URL,
        "Özel Okul Ödemeleri sayfasında anlaşmalı özel okul tahsilatları ve şube hizmeti açıkça belirtiliyor.",
    )


def _isbank_service_rows() -> List[SupplementalRow]:
    school_html = _fetch_html(ISBANK_SCHOOL_URL, must_contain=("özel okul", "Bankamatik"))
    aidat_html = _fetch_html(ISBANK_SITE_AIDAT_URL, must_contain=("Apartman", "Site Tahsilat", "aidat"))
    school = _norm(BeautifulSoup(school_html, "html.parser").get_text(" ", strip=True))
    aidat = _norm(BeautifulSoup(aidat_html, "html.parser").get_text(" ", strip=True))

    rows = []
    if "ozel okul" in school and "bankamatik" in school:
        rows += _add_service_status(
            "İŞBANKASI", "OZEL_OKUL", "Özel Okul Ödemeleri", ("SUBE", "ATM"), ISBANK_SCHOOL_URL,
            "Özel okul ödeme sayfası şube ve Bankamatik kanallarını açıkça belirtiyor.",
        )
    if "site tahsilat" in aidat and "aidat" in aidat:
        rows += _add_service_status(
            "İŞBANKASI", "AIDAT", "Apartman / Site Aidat Ödemeleri", ("GENEL",), ISBANK_SITE_AIDAT_URL,
            "Apartman Yönetim ve Site Tahsilat Sistemi sayfası aidatların hesaptan otomatik tahsil edilebildiğini doğruluyor.",
        )
    return rows


def _akbank_phone_tax_rows() -> List[SupplementalRow]:
    center = _fetch_html(AKBANK_PAYMENT_CENTER_URL, must_contain=("telefon", "Eğitim Ödemeleri"))
    center_text = _norm(BeautifulSoup(center, "html.parser").get_text(" ", strip=True))
    if "telefon" not in center_text or "fatura" not in center_text:
        raise SupplementalSourceError("Akbank Ödeme Merkezi telefon/fatura hizmeti doğrulanamadı.")

    tax = _fetch_html(AKBANK_TAX_URL, must_contain=("Vergi", "Akbank İnternet", "Akbank Mobil"))
    tax_text = _norm(BeautifulSoup(tax, "html.parser").get_text(" ", strip=True))
    if "vergi" not in tax_text:
        raise SupplementalSourceError("Akbank Vergi Ödemeleri sayfası doğrulanamadı.")

    rows: List[SupplementalRow] = []
    rows += _add_service_status(
        "AKBANK", "TELEFON", "Telefon / Cep Telefonu Faturası Ödemeleri",
        ("MOBIL", "SUBE"), AKBANK_PAYMENT_CENTER_URL,
        "Ödeme Merkezi elektrik, su, doğalgaz, telefon, cep telefonu, internet ve TV faturalarının ödenebildiğini doğruluyor. Genel fatura/kurum tarifesi ücret sayfasından eşleştirilir.",
    )
    rows += _add_service_status(
        "AKBANK", "VERGI", "Vergi Ödemeleri",
        ("MOBIL", "SUBE"), AKBANK_TAX_URL,
        "Vergi sayfası Akbank Mobil/İnternet ile şube ve Telefon Şubesi kanallarını doğruluyor; ayrı vergi aracılık tarifesi bulunmazsa yalnız hizmet durumu gösterilir.",
        display_text="Hizmet var\\nAyrı vergi aracılık ücreti yayımlanmıyor",
    )
    return rows


def _yk_phone_tax_fast_rows() -> List[SupplementalRow]:
    bill = _fetch_html(YAPIKREDI_BILL_URL, must_contain=("cep telefonu", "şubeler"))
    bill_text = _norm(BeautifulSoup(bill, "html.parser").get_text(" ", strip=True))
    if "cep telefonu" not in bill_text:
        raise SupplementalSourceError("Yapı Kredi fatura/telefon sayfası doğrulanamadı.")

    mim = _fetch_html(YAPIKREDI_MIM_URL, must_contain=("Vergi / Devlet", "SGK Ödemeleri"))
    mim_text = _norm(BeautifulSoup(mim, "html.parser").get_text(" ", strip=True))
    if "vergi / devlet" not in mim_text:
        raise SupplementalSourceError("Yapı Kredi Müşteri İletişim Merkezi vergi hizmeti doğrulanamadı.")

    fast = _fetch_html(YAPIKREDI_FAST_URL, must_contain=("100.000", "FAST"))
    fast_text = _norm(BeautifulSoup(fast, "html.parser").get_text(" ", strip=True))
    if "100.000" not in fast_text and "100000" not in fast_text:
        raise SupplementalSourceError("Yapı Kredi FAST 100.000 TL limiti doğrulanamadı.")

    rows: List[SupplementalRow] = []
    rows += _add_service_status(
        "YAPIKREDI", "TELEFON", "Telefon / Cep Telefonu Faturası Ödemeleri",
        ("MOBIL", "SUBE"), YAPIKREDI_BILL_URL,
        "Fatura sayfası telekom/cep telefonu faturalarının Mobil, İnternet, Müşteri İletişim Merkezi ve şubelerden ödenebildiğini doğruluyor. Genel fatura/kurum tarifesi ücret sayfasından eşleştirilir.",
    )
    rows += _add_service_status(
        "YAPIKREDI", "VERGI", "Vergi / Devlet Ödemeleri",
        ("MOBIL", "SUBE"), YAPIKREDI_MIM_URL,
        "Müşteri İletişim Merkezi sayfası Vergi/Devlet ödemelerini yayımlıyor; Mobil/İnternet tarafında MTV/vergi hizmetleri ayrıca banka sitesinde yer alıyor.",
        display_text="Hizmet var\\nAyrı vergi aracılık ücreti yayımlanmıyor",
    )
    rows += _add_service_status(
        "YAPIKREDI", "FAST", "FAST - 399.000,01 TL ve üzeri",
        ("MOBIL", "SUBE"), YAPIKREDI_FAST_URL,
        "FAST işlem üst limiti 100.000 TL olarak yayımlanıyor.",
        marker=STATUS_NOT_APPLICABLE,
        display_text="Uygulanmıyor\\nFAST limiti 100.000 TRY",
        band="TRANSFER_3",
    )
    return rows


def _garanti_fast_rows() -> List[SupplementalRow]:
    html = _fetch_html(GARANTI_FAST_URL, must_contain=("100.000", "FAST"))
    page = _norm(BeautifulSoup(html, "html.parser").get_text(" ", strip=True))
    if "100.000" not in page and "100000" not in page:
        raise SupplementalSourceError("Garanti BBVA FAST 100.000 TL limiti doğrulanamadı.")
    return _add_service_status(
        "GARANTİ", "FAST", "FAST - 399.000,01 TL ve üzeri",
        ("MOBIL", "SUBE"), GARANTI_FAST_URL,
        "FAST işlem üst limiti 100.000 TL; FAST gönderimi Mobil/İnternet üzerinden yayımlanıyor.",
        marker=STATUS_NOT_APPLICABLE,
        display_text="Uygulanmıyor\\nFAST limiti 100.000 TRY",
        band="TRANSFER_3",
    )


def _isbank_tax_findeks_rows() -> List[SupplementalRow]:
    tax = _fetch_html(ISBANK_TAX_URL, must_contain=("İşCep", "İnternet Şubesi", "Çözüm Merkezi"))
    tax_text = _norm(BeautifulSoup(tax, "html.parser").get_text(" ", strip=True))
    if "vergi" not in tax_text:
        raise SupplementalSourceError("İş Bankası Vergi Ödeme sayfası doğrulanamadı.")

    findeks = _fetch_html(ISBANK_FINDEKS_URL, must_contain=("Çek Raporu", "Yılda 50 Çek Raporu"))
    findeks_text = _clean(BeautifulSoup(findeks, "html.parser").get_text(" ", strip=True))
    m = re.search(r"Yılda\s*50\s*Çek\s*Raporu.{0,100}?([0-9][0-9.]*,[0-9]{2})\s*TL", findeks_text, flags=re.I | re.S)
    package = m.group(1) if m else "3.660,00"

    rows: List[SupplementalRow] = []
    rows += _add_service_status(
        "İŞBANKASI", "VERGI", "Vergi / Harç Ödemeleri",
        ("MOBIL", "SUBE"), ISBANK_TAX_URL,
        "Vergi sayfası İşCep, İnternet Şubesi ve Çözüm Merkezi kanallarını doğruluyor.",
        display_text="Hizmet var\\nAyrı vergi aracılık ücreti yayımlanmıyor",
    )
    rows += _add_service_status(
        "İŞBANKASI", "CEK_RISK", "Findeks Çek Raporu",
        ("MOBIL", "SUBE"), ISBANK_FINDEKS_URL,
        "Findeks sayfası tek rapor tarifesi yerine yıllık Çek Raporu paketleri yayımlıyor.",
        display_text=f"Tek rapor ücreti ayrı yayımlanmıyor\\n50 rapor/yıl: {package} TRY (KDV dahil)",
    )
    return rows



def _garanti_phone_tax_rows() -> List[SupplementalRow]:
    """Garanti telefon ve vergi hizmetinin resmî kanal varlığını doğrular.

    Ücret, mümkünse ana Ürün/Hizmet Ücretleri sayfasındaki genel
    Fatura/Kurum tarifesinden karşılaştırma katmanında alınır. Burada ücret
    uydurulmaz; yalnız hizmet/kanal kanıtı eklenir.
    """
    bill = _fetch_html(GARANTI_BILL_URL, must_contain=("fatura",))
    bill_text = _norm(BeautifulSoup(bill, "html.parser").get_text(" ", strip=True))
    if not any(x in bill_text for x in ("telefon", "cep telefonu", "gsm", "turkcell", "vodafone")):
        raise SupplementalSourceError("Garanti BBVA fatura sayfasında telefon/cep telefonu hizmeti doğrulanamadı.")

    tax = _fetch_html(GARANTI_TAX_URL, must_contain=("vergi",))
    tax_text = _norm(BeautifulSoup(tax, "html.parser").get_text(" ", strip=True))
    if "vergi" not in tax_text:
        raise SupplementalSourceError("Garanti BBVA vergi ödeme sayfası doğrulanamadı.")

    fee = _fetch_html(GARANTI_FEE_URL, must_contain=("aidat",))
    fee_text = _norm(BeautifulSoup(fee, "html.parser").get_text(" ", strip=True))
    if "aidat" not in fee_text:
        raise SupplementalSourceError("Garanti BBVA ücret sayfasında aidat/kurum tahsilatı doğrulanamadı.")

    rows: List[SupplementalRow] = []
    rows += _add_service_status(
        "GARANTİ", "TELEFON", "Telefon / Cep Telefonu Faturası Ödemeleri",
        ("MOBIL", "SUBE"), GARANTI_BILL_URL,
        "Fatura ödeme sayfası telefon/cep telefonu faturalarının dijital kanallardan ödenebildiğini; banka servis sayfaları Müşteri İletişim Merkezi/ATM gibi ek kanalları doğruluyor. Ayrı telefon tarifesi yoksa genel Fatura/Kurum tarifesi kullanılır.",
    )
    rows += _add_service_status(
        "GARANTİ", "VERGI", "Vergi Ödemeleri",
        ("MOBIL",), GARANTI_TAX_URL,
        "Vergi Ödemeleri sayfası Mobil/İnternet kanalında vergi ödeme hizmetini doğruluyor.",
        display_text="Hizmet var\\nAyrı vergi aracılık ücreti yayımlanmıyor",
    )
    rows += _add_service_status(
        "GARANTİ", "AIDAT", "Aidat / Kurum Tahsilatı",
        ("MOBIL", "SUBE"), GARANTI_FEE_URL,
        "Ürün ve Hizmet Ücretleri sayfasındaki genel Fatura/Kurum açıklaması site/vakıf/aidat tahsilatını kapsıyor. Ayrı aidat tarifesi yoksa genel Fatura/Kurum tarifesi açıkça genel tarife etiketiyle kullanılır.",
    )
    return rows


def _isbank_phone_rows() -> List[SupplementalRow]:
    """İş Bankası telefon faturası ödeme hizmetinin kanal varlığını doğrular."""
    html = _fetch_html(ISBANK_BILL_URL, must_contain=("fatura",))
    text = _norm(BeautifulSoup(html, "html.parser").get_text(" ", strip=True))
    if not any(x in text for x in ("telefon", "cep telefonu", "gsm", "turkcell", "vodafone")):
        raise SupplementalSourceError("İş Bankası Fatura Ödemeleri sayfasında telefon faturası doğrulanamadı.")

    return _add_service_status(
        "İŞBANKASI", "TELEFON", "Telefon / Cep Telefonu Faturası Ödemeleri",
        ("MOBIL", "SUBE"), ISBANK_BILL_URL,
        "Fatura Ödemeleri sayfası telefon faturalarının İşCep/İnternet ve desteklenen şube/Çözüm Merkezi kanallarında ödenebildiğini doğruluyor. Ayrı telefon tarifesi varsa primary ücret satırı önceliklidir.",
    )



def _garanti_comparison_policy_rows() -> List[SupplementalRow]:
    fee_html = _fetch_html(GARANTI_FEE_URL, must_contain=("Şans Oyunları Ödemesi", "Kiralık Kasa"))
    fee_text = _norm(BeautifulSoup(fee_html, "html.parser").get_text(" ", strip=True))
    if "sans oyunlari odemesi" not in fee_text:
        raise SupplementalSourceError("Garanti şans oyunu kanal yapısı doğrulanamadı.")

    rows: List[SupplementalRow] = []
    rows += _add_service_status(
        "GARANTİ", "SANS_OYUNU", "Şans Oyunu Ödemeleri",
        ("SUBE",), GARANTI_FEE_URL,
        "Resmî ücret tablosu bu hizmet için Mobil/İnternet/ATM tarifesi yayımlıyor; ayrı şube tarifesi yok.",
        display_text="Şube için ayrı tarife\\nyayımlanmıyor",
    )
    rows += _add_service_status(
        "GARANTİ", "YURT_DISI_FAST", "Yurt Dışı FAST",
        ("SUBE",), GARANTI_FEE_URL,
        "Yurt Dışı FAST tarifesi Mobil kanal için yayımlanıyor.",
        display_text="Yalnız Mobil tarifesi\\nyayımlanıyor",
    )
    rows += _add_service_status(
        "GARANTİ", "VISA_YP_DIRECT", "Visa ile Yurt Dışı Para Transferi",
        ("SUBE",), GARANTI_FEE_URL,
        "Visa ile Yurt Dışı Para Transferi tarifesi Mobil kanal için yayımlanıyor.",
        display_text="Yalnız Mobil tarifesi\\nyayımlanıyor",
    )
    rows += _add_service_status(
        "GARANTİ", "KASA", "Özel / Süper Kiralık Kasa",
        ("GENEL",), GARANTI_FEE_URL,
        "Resmî ücret tablosunda Büyük/Orta/Küçük kasa tarifeleri var; ayrı Özel/Süper kasa tarifesi yayımlanmıyor.",
        display_text="Ayrı Özel/Süper kasa\\ntarifesi yayımlanmıyor",
    )
    return rows


def _akbank_comparison_policy_rows() -> List[SupplementalRow]:
    fast_html = _fetch_html(AKBANK_FAST_URL, must_contain=("100.000", "FAST"))
    game_html = _fetch_html(AKBANK_GAME_URL, must_contain=("Akbank Mobil", "ATM"))
    fee_html = _fetch_html(AKBANK_FEE_URL, must_contain=("Kiralık Kasa", "Şans Oyunu"))
    fast_text = _norm(BeautifulSoup(fast_html, "html.parser").get_text(" ", strip=True))
    game_text = _norm(BeautifulSoup(game_html, "html.parser").get_text(" ", strip=True))
    if "100.000" not in fast_text and "100000" not in fast_text:
        raise SupplementalSourceError("Akbank FAST 100.000 TL limiti doğrulanamadı.")
    if "atm" not in game_text or "mobil" not in game_text:
        raise SupplementalSourceError("Akbank şans oyunu kanalları doğrulanamadı.")

    rows: List[SupplementalRow] = []
    rows += _add_service_status(
        "AKBANK", "FAST", "FAST - 399.000,01 TL ve üzeri",
        ("MOBIL", "SUBE"), AKBANK_FAST_URL,
        "FAST işlem üst limiti 100.000 TL olarak yayımlanıyor.",
        marker=STATUS_NOT_APPLICABLE,
        display_text="Uygulanmıyor\\nFAST limiti 100.000 TRY",
        band="TRANSFER_3",
    )
    rows += _add_service_status(
        "AKBANK", "SANS_OYUNU", "Şans Oyunu Ödemeleri",
        ("SUBE",), AKBANK_GAME_URL,
        "Resmî hizmet sayfası Akbank Mobil, İnternet ve ATM kanallarını yayımlıyor; ayrı şube kanalı yok.",
        display_text="Şube için ayrı tarife\\nyayımlanmıyor",
    )
    rows += _add_service_status(
        "AKBANK", "KASA", "Özel / Süper Kiralık Kasa",
        ("GENEL",), AKBANK_FEE_URL,
        "Resmî ücret tablosunda Büyük/Orta/Küçük kasa tarifeleri var; ayrı Özel/Süper kasa tarifesi yayımlanmıyor.",
        display_text="Ayrı Özel/Süper kasa\\ntarifesi yayımlanmıyor",
    )
    return rows


def _isbank_fast_sgk_policy_rows() -> List[SupplementalRow]:
    fast_html = _fetch_html(ISBANK_FAST_URL, must_contain=("100.000", "İşCep", "İnternet Şubesi"))
    sgk_html = _fetch_html(ISBANK_SGK_URL, must_contain=("İşCep", "İnternet Şubesi"))
    fast_text = _norm(BeautifulSoup(fast_html, "html.parser").get_text(" ", strip=True))
    if "100.000" not in fast_text and "100000" not in fast_text:
        raise SupplementalSourceError("İş Bankası FAST 100.000 TL limiti doğrulanamadı.")

    rows: List[SupplementalRow] = []
    # İş Bankası FAST yalnız İşCep/İnternet Şubesi üzerinden yayımlanıyor.
    rows += _add_service_status(
        "İŞBANKASI", "FAST", "FAST Şube Kanalı",
        ("SUBE",), ISBANK_FAST_URL,
        "FAST hizmet sayfası kanalları İşCep ve İnternet Şubesi olarak yayımlıyor.",
        marker=STATUS_NOT_APPLICABLE,
        display_text="Şube FAST kanalı\\nyayımlanmıyor",
    )
    rows += _add_service_status(
        "İŞBANKASI", "FAST", "FAST - 399.000,01 TL ve üzeri",
        ("MOBIL", "SUBE"), ISBANK_FAST_URL,
        "FAST işlem üst limiti 100.000 TL.",
        marker=STATUS_NOT_APPLICABLE,
        display_text="Uygulanmıyor\\nFAST limiti 100.000 TRY",
        band="TRANSFER_3",
    )
    rows += _add_service_status(
        "İŞBANKASI", "SGK", "SGK Prim Ödemesi",
        ("SUBE",), ISBANK_SGK_URL,
        "SGK prim ödemesi/talimatı resmî sayfada İşCep ve İnternet Şubesi kanallarıyla yayımlanıyor.",
        display_text="Şube için ayrı SGK kart\\nödeme tarifesi yayımlanmıyor",
    )
    rows += _add_service_status(
        "İŞBANKASI", "YURT_DISI_FAST", "Yurt Dışı FAST / Global FAST",
        ("MOBIL", "SUBE"), ISBANK_FAST_URL,
        "FAST sayfası yurtiçi bankalararası TL FAST hizmetini yayımlıyor; ayrı Yurt Dışı FAST/Global FAST tarifesi yayımlanmıyor.",
        display_text="Ayrı Yurt Dışı FAST /\\nGlobal FAST tarifesi yayımlanmıyor",
    )
    return rows


def _isbank_card_contract_rows() -> List[SupplementalRow]:
    """Güncel resmî kart sözleşmesinden Ortak ATM para yatırma + Visa Direct durumunu alır."""
    import pdfplumber

    pdf_bytes = _request(ISBANK_CARD_CURRENT_URL, binary=True, timeout=60)
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)

    flat = " ".join(text.split())
    norm = _norm(flat)
    if "visa direct hizmeti henuz uygulamada olmayip" not in norm:
        raise SupplementalSourceError("İş Bankası kart sözleşmesinde Visa Direct durum cümlesi bulunamadı.")

    # Ortak ATM tablosunun Standart ATM sütunu:
    # Cari Hesaba Para Yatırma = %1,15 + 1,05 TL (vergi dahil).
    m = re.search(
        r"Cari Hesaba Para Yatırma.{0,180}?%?\s*1[,\.]15\s*\+?\s*1[,\.]05\s*TL",
        flat, flags=re.I | re.S,
    )
    if not m:
        raise SupplementalSourceError("İş Bankası Ortak ATM para yatırma ücreti PDF'den parse edilemedi.")

    rows = [
        SupplementalRow(
            kategori="EK KAYNAK - İş Bankası Kart Sözleşmesi - Ortak ATM",
            masraf="Ortak ATM - Cari Hesaba Para Yatırma (Standart ATM)",
            asgari_tutar="1,05 TL",
            asgari_oran="1,15%",
            azami_oran="1,15%",
            aciklama=_source_note(
                STATUS_NUMERIC, ISBANK_CARD_CURRENT_URL,
                "SERVICE=PARA_YATIRMA; CHANNEL=GENEL; Vergi dahildir; Standart ATM tarifesidir. TEK ATM tarifesi ayrıca %1,15 + 1,58 TL'dir.",
            ),
            site_guncelleme_tarihi="12.08.2026",
        )
    ]
    rows += _add_service_status(
        "İŞBANKASI", "VISA_YP_DIRECT", "Visa Direct",
        ("MOBIL", "SUBE"), ISBANK_CARD_CURRENT_URL,
        "Kart sözleşmesi Visa Direct hizmetinin henüz uygulamada olmadığını açıkça belirtiyor.",
        marker=STATUS_NOT_APPLICABLE,
        display_text="Henüz uygulamada değil",
    )
    return rows


def _yk_comparison_policy_rows() -> List[SupplementalRow]:
    fast_html = _fetch_html(YAPIKREDI_FAST_URL, must_contain=("100.000", "Yapı Kredi Mobil", "İnternet Şubesi"))
    text = _norm(BeautifulSoup(fast_html, "html.parser").get_text(" ", strip=True))
    if "100.000" not in text and "100000" not in text:
        raise SupplementalSourceError("Yapı Kredi FAST limit/kanal doğrulaması başarısız.")

    rows: List[SupplementalRow] = []
    rows += _add_service_status(
        "YAPIKREDI", "FAST", "FAST Şube Kanalı",
        ("SUBE",), YAPIKREDI_FAST_URL,
        "FAST hizmet sayfası Yapı Kredi Mobil ve Bireysel İnternet Şubesi kanallarını yayımlıyor.",
        marker=STATUS_NOT_APPLICABLE,
        display_text="Şube FAST kanalı\\nyayımlanmıyor",
    )
    # Düzenli transfer açıklamaları dijital tarife altında yayımlanıyor;
    # standart şube EFT/Havale ücretini düzenli talimata otomatik taşımıyoruz.
    rows += _add_service_status(
        "YAPIKREDI", "DUZENLI_EFT", "Düzenli EFT - Şube",
        ("SUBE",), YAPIKREDI_REGULAR_URL,
        "Düzenli ödeme/talimat hizmeti mevcut; ücret tablosunda düzenli EFT için şubeye özel ayrı tarife yayımlanmıyor.",
        display_text="Düzenli EFT için Şube\\ntarifesi ayrı yayımlanmıyor",
    )
    rows += _add_service_status(
        "YAPIKREDI", "DUZENLI_HAVALE", "Düzenli Havale - Şube",
        ("SUBE",), YAPIKREDI_REGULAR_URL,
        "Düzenli ödeme/talimat hizmeti mevcut; ücret tablosunda düzenli Havale için şubeye özel ayrı tarife yayımlanmıyor.",
        display_text="Düzenli Havale için Şube\\ntarifesi ayrı yayımlanmıyor",
    )
    return rows


def _discover_isbank_bhs() -> Tuple[str, str]:
    """En güncel normal Bankacılık Hizmetleri Sözleşmesi linkini bulur."""
    html = _fetch_html(ISBANK_CONTRACTS_URL, must_contain=("Bankacılık Hizmetleri Sözleşmesi",))
    soup = BeautifulSoup(html, "html.parser")

    candidates: List[Tuple[int, str, str]] = []
    for a in soup.find_all("a", href=True):
        text = _clean(a.get_text(" ", strip=True))
        key = _norm(text)
        if "bankacilik hizmetleri sozlesmesi" not in key:
            continue
        if "mesafeli" in key or "kredi kart" in key or "ingilizce" in key:
            continue
        href = urljoin(ISBANK_CONTRACTS_URL, a.get("href"))
        if not href:
            continue

        m = re.search(r"(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})", text)
        if m:
            d, mo, y = map(int, m.groups())
            score = y * 10000 + mo * 100 + d
            date = f"{d:02d}.{mo:02d}.{y:04d}"
        else:
            y = re.search(r"20\d{2}", text)
            score = int(y.group(0)) * 10000 if y else 0
            date = ""

        candidates.append((score, href, date))

    if not candidates:
        return ISBANK_BHS_FALLBACK_URL, "13.02.2026"

    candidates.sort(reverse=True)
    return candidates[0][1], candidates[0][2]


def _num_token(raw: str) -> str:
    return _clean(raw).replace(" ", "")


def _first_amount(pattern: str, text: str) -> str:
    m = re.search(pattern + r"[^\d%]{0,80}([0-9][0-9.,]*)\s*TL", text, flags=re.I | re.S)
    return _num_token(m.group(1)) + " TL" if m else ""


def _first_rate_minmax(pattern: str, text: str) -> Tuple[str, str, str, str]:
    m = re.search(
        pattern
        + r"[^%\d]{0,80}%\s*([0-9.,]+)\s*[-–—]\s*En\s*(?:Az|az)\s*([0-9][0-9.,]*)\s*TL"
          r"\s*[-–—]\s*En\s*(?:Çok|Cok|çok)\s*([0-9][0-9.,]*)\s*TL",
        text,
        flags=re.I | re.S,
    )
    if not m:
        return "", "", "", ""
    rate, min_amt, max_amt = m.groups()
    return _num_token(min_amt) + " TL", _num_token(rate), _num_token(max_amt) + " TL", _num_token(rate)


def _extract_isbank_bhs_rows(pdf_bytes: bytes, url: str, date: str) -> List[SupplementalRow]:
    try:
        import pdfplumber
    except ImportError as exc:
        raise SupplementalSourceError("pdfplumber bulunamadı; İş Bankası BHS PDF'i okunamıyor.") from exc

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        # Bireysel çek/senet bölümü güncel BHS'de ilk birkaç sayfada.
        page_texts = []
        for page in pdf.pages[:6]:
            page_texts.append(page.extract_text() or "")
    text = "\n".join(page_texts)
    flat = " ".join(text.replace("\xa0", " ").split())

    if "ÇEK BELGELENDİRME" not in flat.upper() or "SENET" not in flat.upper():
        raise SupplementalSourceError("İş Bankası BHS içinde çek/senet bölümü bulunamadı.")

    note = lambda extra="": _source_note(STATUS_NUMERIC, url, "BSMV hariç. " + extra)
    rows: List[SupplementalRow] = []

    def add_amount(category: str, label: str, pattern: str, extra: str = ""):
        value = _first_amount(pattern, flat)
        if value:
            rows.append(SupplementalRow(
                kategori=f"EK KAYNAK - İş Bankası BHS - {category}",
                masraf=label,
                asgari_tutar=value,
                azami_tutar=value,
                aciklama=note(extra),
                site_guncelleme_tarihi=date,
            ))

    def add_rate(category: str, label: str, pattern: str, extra: str = ""):
        min_amt, min_rate, max_amt, max_rate = _first_rate_minmax(pattern, flat)
        if any((min_amt, min_rate, max_amt, max_rate)):
            rows.append(SupplementalRow(
                kategori=f"EK KAYNAK - İş Bankası BHS - {category}",
                masraf=label,
                asgari_tutar=min_amt,
                asgari_oran=min_rate,
                azami_tutar=max_amt,
                azami_oran=max_rate,
                aciklama=note(extra),
                site_guncelleme_tarihi=date,
            ))

    # Senet
    add_amount("Senet İşlemleri", "Senet Bilgilendirme Ücreti", r"Senet Bilgilendirme Ücreti\*?\s+Senet Bilgilendirme Ücreti")
    add_amount("Senet İşlemleri", "Aynı Şube Senet Tahsili", r"Aynı Şube Senet Tahsili")
    # Tabloda aynı ücret hücresi iki tahsilat satırını kapsayabiliyor; aynı şube değeri yayımlanmışsa farklı şubeye de uygula.
    same_senet = next((r for r in rows if r.masraf == "Aynı Şube Senet Tahsili"), None)
    if same_senet and "Farklı Şube Senet Tahsili" in flat:
        rows.append(SupplementalRow(
            kategori="EK KAYNAK - İş Bankası BHS - Senet İşlemleri",
            masraf="Farklı Şube Senet Tahsili",
            asgari_tutar=same_senet.asgari_tutar,
            azami_tutar=same_senet.azami_tutar,
            aciklama=note("BHS tablosunda aynı tarife hücresi Aynı/Farklı Şube Senet Tahsili satırlarını kapsıyor."),
            site_guncelleme_tarihi=date,
        ))
    add_amount("Senet İşlemleri", "Senet İade", r"Senet İade Ücreti\s+Senet İade")
    protest = re.search(r"Senet Protestosu\s+([0-9][0-9.,]*)\s*TL\s+Senet Protesto Kaldırma", flat, flags=re.I)
    if protest:
        value = _num_token(protest.group(1)) + " TL"
        for label in ("Senet Protestosu", "Senet Protesto Kaldırma"):
            rows.append(SupplementalRow(
                kategori="EK KAYNAK - İş Bankası BHS - Senet İşlemleri",
                masraf=label,
                asgari_tutar=value,
                azami_tutar=value,
                aciklama=note(),
                site_guncelleme_tarihi=date,
            ))

    # Çek tahsil / ödeme
    for label, pattern in (
        ("Bankamız TP/YP - Tahsile Alınan", r"Bankamız TP/YP\s*[–-]\s*Tahsile Alınan"),
        ("Diğer Banka TP - Tahsile Alınan", r"Diğer Banka TP\*?\s*[–-]\s*Tahsile Alınan"),
        ("Diğer Banka YP - Tahsile Alınan", r"Diğer Banka YP\*?\s*[–-]\s*Tahsile Alınan"),
        ("Bankamız TP/YP - Teminata Alınan", r"Bankamız TP/YP\s*[–-]\s*Teminata Alınan"),
        ("Diğer Banka TP - Teminata Alınan", r"Diğer Banka TP\s*[–-]\s*Teminata Alınan"),
    ):
        add_amount("Çek Tahsilatı", label, pattern)

    for label, pattern in (
        ("Dövizli Çek - Tahsile Alınan", r"Dövizli Çek\s*-\s*Tahsile Alınan"),
        ("Aynı Şube Çek Ödeme TP/YP", r"Aynı Şube Çek Ödeme TP/YP"),
        ("Farklı Şube Çek Ödeme TP/YP", r"Farklı Şube Çek Ödeme TP/YP"),
        ("Bloke Çek Ödeme TP/YP", r"Bloke Çek Ödeme TP/YP"),
        ("Dövizli Çek Ödeme", r"Dövizli Çek Ödeme"),
        ("Bankamız Üzerine Keşideli Dövizli Çek Alışı", r"Bankamız Üzerine Keşideli Dövizli Çek Alışı"),
        ("Yabancı Banka Üzerine Keşideli Döviz Çeki Alışı", r"Yabancı Banka Üzerine Keşideli Döviz Çeki Alışı"),
        ("Diğer Banka TP/YP Takas Harici Çek Ödeme", r"Diğer Banka TP/YP Takas Harici Çek Ödeme"),
        ("Dövizli Çek Düzenleme", r"Dövizli Çek Düzenleme\s+%"),
        ("Dövizli Çek Düzenleme - Dövizi Natık Çek Karşılığı", r"Dövizli Çek Düzenleme\s*[–-]\s*Dövizi Natık Çek Karşılığı"),
        ("Bloke Çek Düzenleme TP/YP", r"Bloke Çek Düzenleme TP/YP"),
    ):
        add_rate("Çek İşlemleri", label, pattern)

    # Çek defteri - yurtiçi tarifeler.
    for label, pattern in (
        ("Karekodlu Çek Defteri - 10 Yapraklı (Yurtiçi Şubeler)", r"Karekodlu Çek Defteri Ücreti\*?\s+10 Yapraklı \(Yurtiçi Şubeler\)"),
        ("Karekodlu Çek Defteri - 25 Yapraklı (Yurtiçi Şubeler)", r"Karekodlu Çek Defteri Ücreti\*?.{0,120}?25 Yapraklı \(Yurtiçi Şubeler\)"),
        ("Karekodlu ve Logolu Çek Defteri - 10 Yapraklı (Yurtiçi Şubeler)", r"Karekodlu ve Logolu Çek Defteri(?: Ücreti)?\*?.{0,40}?10 Yapraklı \(Yurtiçi Şubeler\)"),
        ("Karekodlu ve Logolu Çek Defteri - 25 Yapraklı (Yurtiçi Şubeler)", r"Karekodlu ve Logolu Çek Defteri(?: Ücreti)?\*?.{0,120}?25 Yapraklı \(Yurtiçi Şubeler\)"),
    ):
        add_amount("Çek Defteri", label, pattern, "Değerli Kâğıt Bedeli ayrıca tahsil edilir.")

    # Yaprak başı 50-350 ve 351+ değerlerini ayrı yakala.
    for kind in ("Karekodlu Çek Defteri", "Karekodlu ve Logolu Çek Defteri"):
        kind_re = re.escape(kind)
        block_match = re.search(kind_re + r".{0,550}", flat, flags=re.I | re.S)
        if block_match:
            block = block_match.group(0)
            for range_label in ("50 – 350", "351 Yaprak ve üzeri"):
                if range_label.startswith("50"):
                    m = re.search(r"50\s*[–-]\s*350 Yapraklı \(Yurtiçi Şubeler\) Yaprak Başı\s*([0-9][0-9.,]*)\s*TL", block, flags=re.I)
                else:
                    m = re.search(r"351 Yaprak ve (?:üzeri|Üzeri) \(Yurtiçi Şubeler\) Yaprak Başı\s*([0-9][0-9.,]*)\s*TL", block, flags=re.I)
                if m:
                    value = _num_token(m.group(1)) + " TL"
                    rows.append(SupplementalRow(
                        kategori="EK KAYNAK - İş Bankası BHS - Çek Defteri",
                        masraf=f"{kind} - {range_label} (Yurtiçi Şubeler) - Yaprak Başı",
                        asgari_tutar=value,
                        azami_tutar=value,
                        aciklama=note("Değerli Kâğıt Bedeli ayrıca tahsil edilir."),
                        site_guncelleme_tarihi=date,
                    ))

    add_amount("Çek İade", "Çek Muamelesiz İade Ücreti - Çek Başına", r"Çek Muamelesiz İade Ücreti\s+Çek Başına")
    add_amount("Çek Belgelendirme ve Düzeltme", "Çek Düzeltme Ücreti - Çek Başına", r"Çek Düzeltme Ücreti\s+Çek Başına")

    # Güncel BHS bölüm başlığında belgelendirme + düzeltme birlikte geçiyor,
    # ancak ayrı bir Karşılıksız Çek Belgelendirme ücret tutarı yayımlanmıyor.
    if "ÇEK BELGELENDİRME VE DÜZELTME" in flat.upper():
        rows.append(SupplementalRow(
            kategori="EK KAYNAK - İş Bankası BHS - Çek Belgelendirme ve Düzeltme",
            masraf="Karşılıksız Çek Belgelendirme",
            aciklama=_source_note(
                STATUS_EMPTY, url,
                "SERVICE=CEK_KARSILIKSIZ; BHS bölümünde ayrı belgelendirme tarife tutarı yayımlanmıyor; DISPLAY_TEXT=Ayrı belgelendirme ücreti yayımlanmıyor",
            ),
            site_guncelleme_tarihi=date,
        ))

    # Kritik doğrulama: kullanıcı için asıl eksik olan düzeltme satırı mutlaka gelmeli.
    critical_checks = {
        "çek düzeltme": any("cek duzeltme ucreti" in _norm(r.masraf) and r.asgari_tutar for r in rows),
        "çek iade": any("cek muamelesiz iade" in _norm(r.masraf) and r.asgari_tutar for r in rows),
        "senet iade": any("senet iade" in _norm(r.masraf) and r.asgari_tutar for r in rows),
    }
    missing = [name for name, ok in critical_checks.items() if not ok]
    if missing:
        raise SupplementalSourceError(
            "İş Bankası BHS kritik satırları parse edilemedi: " + ", ".join(missing)
        )

    return rows


def _isbank_bhs_rows() -> Tuple[List[SupplementalRow], str]:
    url, date = _discover_isbank_bhs()
    try:
        pdf_bytes = _request(url, binary=True, timeout=60)
    except Exception:
        # Dinamik link bir viewer/link wrapper ise güncel doğrudan PDF fallback'i dene.
        url = ISBANK_BHS_FALLBACK_URL
        date = date or "13.02.2026"
        pdf_bytes = _request(url, binary=True, timeout=60)
    return _extract_isbank_bhs_rows(pdf_bytes, url, date), url


def enrich_all(banka_verileri: Mapping[str, Sequence]) -> Tuple[Dict[str, List], SupplementalReport]:
    """
    Primary scraper sonuçlarını kopyalar ve ek resmî kaynak satırlarını ekler.
    Her kritik ek kaynak doğrulanamazsa report.ok=False döner; main.py bu durumda
    mevcut doğru Excel'i korumalıdır.
    """
    report = SupplementalReport()
    result: Dict[str, List] = {bank: list(rows) for bank, rows in banka_verileri.items()}

    def apply(bank: str, source_name: str, url: str, producer, *, required: bool = True):
        if bank not in result:
            if required:
                report.fail(f"{source_name}: {bank} primary verisi yok.")
            return
        try:
            additions = producer()
            added = _append_unique(result[bank], additions)
            report.added_by_bank[bank] = report.added_by_bank.get(bank, 0) + added
            report.source_counts[source_name] = len(additions)
            report.sources[source_name] = url
            print(f"[supplemental] {source_name}: bulunan={len(additions)}, eklenen={added}")
        except Exception as exc:
            message = f"{source_name}: {exc}"
            if required:
                report.fail(message)
                print(f"[supplemental][FATAL] {message}")
            else:
                report.warnings.append(message)
                print(f"[supplemental][UYARI] {message}")

    print(f"[supplemental] SÜRÜM: {SUPPLEMENTAL_VERSION}")

    # İş Bankası BHS URL'i dinamik keşfedildiği için özel çağrı.
    if "İŞBANKASI" in result:
        try:
            rows, actual_url = _isbank_bhs_rows()
            added = _append_unique(result["İŞBANKASI"], rows)
            report.added_by_bank["İŞBANKASI"] = report.added_by_bank.get("İŞBANKASI", 0) + added
            report.source_counts["İŞBANKASI_BHS"] = len(rows)
            report.sources["İŞBANKASI_BHS"] = actual_url
            print(f"[supplemental] İŞBANKASI_BHS: bulunan={len(rows)}, eklenen={added}, url={actual_url}")
        except Exception as exc:
            report.fail(f"İŞBANKASI_BHS: {exc}")
            print(f"[supplemental][FATAL] İŞBANKASI_BHS: {exc}")
    else:
        report.fail("İŞBANKASI_BHS: İŞBANKASI primary verisi yok.")

    def akbank_commercial_rows():
        html = _fetch_html(
            AKBANK_COMMERCIAL_URL,
            must_contain=("Çek Belgelendirme ve Düzeltme", "Karşılıksız Çek Belgelendirme", "Çek Düzeltme Hakkı"),
        )
        parsed = _parse_akbank_commercial_html(html)
        if not parsed:
            raise SupplementalSourceError("Akbank ticari ücret sayfasından satır parse edilemedi.")
        required_labels = ("karsiliksiz cek belgelendirme", "cek duzeltme hakki")
        page_rows = " | ".join(_norm(r.masraf) for r in parsed)
        missing = [x for x in required_labels if x not in page_rows]
        if missing:
            raise SupplementalSourceError("Akbank kritik çek alt kalemleri parse edilemedi: " + ", ".join(missing))
        return parsed

    apply(
        "AKBANK", "AKBANK_TICARI", AKBANK_COMMERCIAL_URL,
        akbank_commercial_rows,
    )
    apply("AKBANK", "AKBANK_HIZMETLER", AKBANK_REGULAR_URL, _akbank_service_rows)
    apply("YAPIKREDI", "YAPIKREDI_HIZMETLER", YAPIKREDI_REGULAR_URL, _yk_service_rows)

    # Karşılaştırmada N/A yerine doğrulanmış hizmet durumu / genel tarife kullanabilmek
    # için gerekli resmî ikincil kaynaklar. Bunlardan biri doğrulanamazsa final Excel
    # güncellenmez; son doğru dosya korunur.
    apply("AKBANK", "AKBANK_TELEFON_VERGI", AKBANK_PAYMENT_CENTER_URL, _akbank_phone_tax_rows)
    apply("YAPIKREDI", "YAPIKREDI_TELEFON_VERGI_FAST", YAPIKREDI_BILL_URL, _yk_phone_tax_fast_rows)
    apply("GARANTİ", "GARANTI_FAST", GARANTI_FAST_URL, _garanti_fast_rows)
    apply("İŞBANKASI", "ISBANK_VERGI_FINDEKS", ISBANK_TAX_URL, _isbank_tax_findeks_rows)
    apply("GARANTİ", "GARANTI_TELEFON_VERGI", GARANTI_BILL_URL, _garanti_phone_tax_rows)
    apply("İŞBANKASI", "ISBANK_TELEFON", ISBANK_BILL_URL, _isbank_phone_rows)

    # Özel okul / aidat hizmet kanıtları da karşılaştırma mantığının parçasıdır.
    apply("GARANTİ", "GARANTI_OZEL_OKUL", GARANTI_SCHOOL_URL, _garanti_service_rows)
    apply("İŞBANKASI", "ISBANK_HIZMETLER", ISBANK_SCHOOL_URL, _isbank_service_rows)

    # Karşılaştırma doğruluğu için kanal/limit/publikasyon statüleri.
    apply("GARANTİ", "GARANTI_COMPARISON_POLICY", GARANTI_FEE_URL, _garanti_comparison_policy_rows)
    apply("AKBANK", "AKBANK_COMPARISON_POLICY", AKBANK_FAST_URL, _akbank_comparison_policy_rows)
    apply("İŞBANKASI", "ISBANK_FAST_SGK_POLICY", ISBANK_FAST_URL, _isbank_fast_sgk_policy_rows)
    apply("YAPIKREDI", "YAPIKREDI_COMPARISON_POLICY", YAPIKREDI_FAST_URL, _yk_comparison_policy_rows)

    # Kart sözleşmesi PDF'i kritik primary veri için değil yalnız Ortak ATM/Visa
    # karşılaştırmasını zenginleştirir. PDF yapısı değişirse ana Excel bloke edilmez.
    apply(
        "İŞBANKASI", "ISBANK_CARD_CONTRACT", ISBANK_CARD_CURRENT_URL,
        _isbank_card_contract_rows, required=False,
    )

    print(
        f"[supplemental] SONUÇ: ok={report.ok}, toplam_eklenen={report.total_added}, "
        f"hata={len(report.errors)}, uyarı={len(report.warnings)}"
    )
    return result, report


def print_supplemental_report(report: SupplementalReport) -> None:
    print("\n" + "=" * 60)
    print("EK RESMÎ KAYNAK RAPORU")
    print(f"Sürüm: {report.version}")
    for bank in ("GARANTİ", "İŞBANKASI", "AKBANK", "YAPIKREDI"):
        if bank in report.added_by_bank:
            print(f"{bank:12s}: +{report.added_by_bank[bank]} satır")
    for name, count in report.source_counts.items():
        print(f"  {name}: {count} kaynak satırı")
    if report.warnings:
        print("UYARILAR:")
        for item in report.warnings:
            print(f"  - {item}")
    if report.errors:
        print("HATALAR:")
        for item in report.errors:
            print(f"  - {item}")
    print("DURUM:", "OK" if report.ok else "BAŞARISIZ - Excel yazılmayacak")
    print("=" * 60)


if __name__ == "__main__":
    print("Bu modül main.py tarafından primary scraper sonuçlarını zenginleştirmek için kullanılır.")
