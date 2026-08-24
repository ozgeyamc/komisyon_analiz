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

Bu modül ücret uydurmaz. Ek kaynak yalnız hizmet/kanal varlığını doğruluyorsa ve aynı
bankanın ana resmî ücret tablosunda güvenli biçimde eşleşen bir tarife varsa, o tarifenin
tutar/oranları ek kaynak satırına da taşınır. Eşleşme yoksa ücret kolonları boş kalır.
NOT_APPLICABLE / PUBLISHED_EMPTY gibi resmen ücret bulunmadığını anlatan satırlar hiçbir
zaman yapay bir tutarla doldurulmaz.
"""

from __future__ import annotations

import io
import re
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict, Iterable, List, Mapping, MutableMapping, Optional, Sequence, Set, Tuple
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup


SUPPLEMENTAL_VERSION = "2026-08-24-v13-primary-source-priority"

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
ISBANK_FAQ_URL = "https://www.isbank.com.tr/sss"

ISBANK_FAST_URL = "https://www.isbank.com.tr/fast-anlik-para-transferi"
ISBANK_SGK_URL = "https://www.isbank.com.tr/sgk-odemeleri"
ISBANK_CARD_CURRENT_URL = "https://www.isbank.com.tr/Documents/KKR%20S%C3%B6zle%C5%9Fmesi%20Dijital%20Kanallar%28%20KREDI-KARTI-SOZ%20%2903.02.2026.pdf"

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
YAPIKREDI_FEE_URL = "https://www.yapikredi.com.tr/bireysel-bankacilik/hesaplama-araclari/bireysel-urun-ve-hizmet-ucretleri"
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


# ---------------------------------------------------------------------------
# HİZMET DURUMU SATIRLARINA PRIMARY RESMÎ TARİFE BACKFILL
# ---------------------------------------------------------------------------
# Ek kaynakların bir bölümü yalnız "bu hizmet bu kanalda var" bilgisini doğrular.
# Kullanıcı ana KOMİSYONLAR sayfasında bu satırların tutar/oranını da görmek istediği
# için, aynı bankanın ORIGINAL primary scraper sonucundaki güvenli tarifeyi bu satıra
# taşıyoruz. Burada yeni ücret hesaplanmaz / tahmin edilmez.


def _has_published_fee(row) -> bool:
    """Satırda yayımlanmış bir tutar/oran var mı? 0 / 0,00 da geçerli tarifedir."""
    return any(
        _clean(getattr(row, attr, "")) not in ("", "-", "–", "—")
        for attr in ("asgari_tutar", "asgari_oran", "azami_tutar", "azami_oran")
    )


def _status_meta_value(row: SupplementalRow, key: str) -> str:
    match = re.search(
        rf"(?:^|[;|]\s*){key}\s*=\s*([^;|]+)",
        _clean(getattr(row, "aciklama", "")),
        flags=re.I,
    )
    return _clean(match.group(1)).upper() if match else ""


def _primary_row_text(row) -> Tuple[str, str, str]:
    kategori = _norm(getattr(row, "kategori", ""))
    masraf = _norm(getattr(row, "masraf", ""))
    aciklama = _norm(getattr(row, "aciklama", ""))
    return kategori, masraf, f"{kategori} | {masraf} | {aciklama}"


def _primary_service_tags(row) -> Set[str]:
    """Supplemental kaynaklarda kullanılan kanonik hizmetleri primary satırlarda bulur."""
    kategori, masraf, full = _primary_row_text(row)
    short = f"{kategori} | {masraf}"
    tags: Set[str] = set()

    explicit = re.search(r"service\s*=\s*([a-z0-9_]+)", full, flags=re.I)
    if explicit:
        tags.add(explicit.group(1).upper())

    international = any(x in full for x in (
        "swift", "uluslararasi fon transfer", "yurt disi fast", "yurtdisi fast",
        "global fast", "western union", "fast uluslararasi",
    ))

    if any(x in full for x in ("fast", "fonlarin anlik ve surekli transferi")) and not international:
        tags.add("FAST")
    if any(x in full for x in ("yurt disi fast", "yurtdisi fast", "global fast", "fast uluslararasi")):
        tags.add("YURT_DISI_FAST")
    if any(x in full for x in (
        "visa ile yurt disi para transferi", "visa ile yurtdisi para transferi",
        "visa direct", "visa yp direct",
    )):
        tags.add("VISA_YP_DIRECT")
    if "duzenli" in full and any(x in full for x in ("eft", "elektronik fon transfer")):
        tags.add("DUZENLI_EFT")
    if "duzenli" in full and "havale" in full:
        tags.add("DUZENLI_HAVALE")

    if (
        any(x in masraf for x in (
            "altin transfer", "ats ile altin gonderimi",
            "kiymetli maden transferi ucreti - altin", "kiymetli maden transfer - altin",
        ))
        and not any(x in masraf for x in (
            "fiziki", "teslim", "kulce altin cekme", "western union", "eft", "havale", "fast",
        ))
    ):
        tags.add("ALTIN_TRANSFER")

    if any(x in full for x in ("kiralik kasa", "kasa kiralama", "kasa ucreti")):
        tags.add("KASA")

    combined_risk = any(x in full for x in (
        "kkb cek / risk raporu", "kkb cek/risk raporu", "cek / risk raporu",
    ))
    if combined_risk or any(x in full for x in (
        "cek risk raporu", "cek bilgileri raporu", "kkb cek", "findeks cek raporu",
        "cek sorgu raporu",
    )):
        tags.add("CEK_RISK")

    if (
        ("fatura" in short or "fatura / kurum" in short or "fatura/kurum" in short
         or "fatura tahsil" in short or "kurum tahsil" in short or "kurum odeme" in short)
        and "e-fatura" not in short
    ):
        tags.add("FATURA")

    if "sgk" in short or "sosyal guvenlik" in full:
        tags.add("SGK")

    if "sans oyun" in full or any(x in full for x in (
        "bilyoner", "nesine", "tuttur", "oley", "misli", "sisal sans", "tjk",
    )):
        tags.add("SANS_OYUNU")

    has_aidat = bool(re.search(r"\baidat\b", short)) or (
        bool(re.search(r"\baidat\b", full))
        and any(x in full for x in ("fatura", "tahsilat", "odeme", "site", "apartman"))
    )
    if has_aidat and not any(x in full for x in (
        "aidatsiz kart", "kart aidati", "yillik kart ucreti", "uyelik ucreti - kart",
    )):
        tags.add("AIDAT")

    if (
        any(x in full for x in (
            "ozel okul", "okul odeme", "okul taksiti", "egitim odeme", "egitim kurumu odeme",
        ))
        and not any(x in full for x in (
            "mektup", "vize", "konsolosluk", "referans yazisi", "referans mektubu",
        ))
    ):
        tags.add("OZEL_OKUL")

    telefon_candidate = (
        any(x in short for x in (
            "telefon odeme", "telefon fatur", "cep telefonu fatur",
            "telefon operatorleri odemelerine aracilik", "gsm odeme", "telekom odeme",
            "turkcell", "vodafone",
        ))
        or (
            any(x in full for x in ("turkcell", "vodafone", "superonline", "tellcom", "turk telekom"))
            and any(x in short for x in ("fatura", "kurum odeme", "tahsilat"))
        )
    )
    if (
        telefon_candidate
        and "tl/paket yukleme" not in masraf
        and "paket yukleme" not in masraf
        and "otomatik fatura odeme faizi" not in masraf
        and "alisveris faiz" not in masraf
    ):
        tags.add("TELEFON")

    if (
        "vergi" in short
        and any(x in short for x in (
            "vergi tahsil", "vergi odeme", "vergi / devlet", "fatura/vergi/sgk",
            "fatura / vergi / sgk", "mtv", "harc",
        ))
        and not any(x in short for x in ("vergi numarasi", "vergi yazisi", "kredi"))
    ):
        tags.add("VERGI")

    return tags


def _primary_channels(row) -> Set[str]:
    """Primary satırın kanal kümesini çıkarır."""
    _, masraf, full = _primary_row_text(row)
    explicit = re.search(r"channel\s*=\s*([a-z0-9_]+)", full, flags=re.I)
    if explicit:
        channel = explicit.group(1).upper()
        return {"GENEL"} if channel == "GENEL" else {channel}

    if "tum kanal" in full:
        return {"MOBIL", "SUBE", "ATM"}

    channels: Set[str] = set()
    mas2 = masraf.replace("internet subesi", "internet").replace("internet sube", "internet")

    if any(x in mas2 for x in (
        "mobil", "internet", "dijital", "cepteteb", "iscep", "web",
    )):
        channels.add("MOBIL")
    if any(x in mas2 for x in (
        "sube", "subeden", "musteri iletisim merkezi", "cozum merkezi", "telefon subesi",
        "gise", "kasadan",
    )):
        channels.add("SUBE")
    if any(x in mas2 for x in ("atm", "btm", "kiosk", "bankamatik")):
        channels.add("ATM")

    if not channels:
        full2 = full.replace("internet subesi", "internet").replace("internet sube", "internet")
        if any(x in full2 for x in ("mobil", "internet", "dijital", "iscep", "web")):
            channels.add("MOBIL")
        if any(x in full2 for x in (
            "sube", "musteri iletisim merkezi", "cozum merkezi", "telefon subesi", "gise", "kasadan",
        )):
            channels.add("SUBE")
        if any(x in full2 for x in ("atm", "bankamatik", "btm")):
            channels.add("ATM")

    return channels or {"GENEL"}


def _generic_institution_score(row, bank: str, wanted_channel: str) -> int:
    """Aidat / okul / telefon için bankanın gerçek genel Fatura/Kurum tarifesini seçer."""
    if not _has_published_fee(row):
        return -10_000

    kategori, masraf, full = _primary_row_text(row)

    # İş Bankası kanal isimleri özel dikkat ister: "İnternet Şube" fiziksel
    # şube değildir. Supplemental backfill doğrudan ilk sayfadaki üç gerçek
    # Fatura Ödemeleri kategorisinden doğru olanı seçer.
    if bank == "İŞBANKASI" and masraf == "fatura odemeleri":
        if wanted_channel == "MOBIL":
            if ("internet sube" in kategori or "iscep" in kategori) and "bankamatik" not in kategori:
                return 1000
            return -10_000
        if wanted_channel == "SUBE":
            if (
                "fatura odemeleri - sube" in kategori
                and "internet sube" not in kategori
                and "iscep" not in kategori
                and "bankamatik" not in kategori
            ):
                return 1000
            return -10_000
        if wanted_channel == "ATM":
            return 1000 if "bankamatik" in kategori else -10_000

    tags = _primary_service_tags(row)
    if "FATURA" not in tags:
        return -10_000

    if not any(x in full for x in (
        "fatura/kurum", "fatura / kurum", "fatura ve anlasmali kurum",
        "fatura odemeleri", "kurum tahsilat", "kurum odeme",
    )):
        return -10_000

    if any(x in masraf for x in (
        "faiz", "otomatik fatura odeme faizi", "talimatli fatura odeme islem faizi",
        "alisveris faiz", "sgk", "sans oyun", "vergi", "tl/paket yukleme",
        "paket yukleme", "nakit avans", "konsolosluk", "vize randevu",
    )):
        return -10_000

    channels = _primary_channels(row)
    score = 100
    if wanted_channel in channels:
        score += 80
    elif "GENEL" in channels:
        # Banka yalnız tek genel tarife yayımlıyorsa mobil/şube hizmet satırına
        # aynı resmî tarife taşınabilir.
        score += 30
    else:
        return -10_000

    # Banka bazında bilinen ana kurum tarifesini öne çıkar.
    if bank == "YAPIKREDI" and "fatura ve anlasmali kurum odemeleri" in masraf:
        score += 220
    elif bank == "AKBANK" and "fatura / kurum tahsil" in masraf:
        score += 190
    elif bank == "GARANTİ" and "fatura/kurum odemesi" in masraf:
        score += 190
    elif bank == "İŞBANKASI" and masraf == "fatura odemeleri":
        score += 190

    # Kaynak türü belirtilmeyen supplemental satırlarda mümkünse hesaptan tarifeyi
    # kredi kartı tarifesine tercih et.
    if "hesaptan" in masraf:
        score += 35
    if "kredi kart" in masraf:
        score -= 20

    # Aynı satır hem dijital hem Çözüm Merkezi gibi ifadeler taşıyorsa,
    # Şube hedefinde gerçek gişe/şube tarifesini; Mobil hedefinde gerçek
    # İnternet/İşCep tarifesini öne çıkar.
    if wanted_channel == "SUBE":
        if (
            " - sube" in kategori
            or kategori.endswith("sube")
            or any(x in masraf for x in ("giseden", "subeden", "sube/", "sube "))
        ):
            score += 65
        elif "cozum merkezi" in kategori or "musteri iletisim merkezi" in kategori:
            score += 15
    elif wanted_channel == "MOBIL":
        if any(x in kategori for x in ("internet", "iscep", "mobil")):
            score += 55
        if any(x in masraf for x in ("mobil", "internet")):
            score += 35
    elif wanted_channel == "ATM":
        if any(x in kategori for x in ("bankamatik", "atm")) or any(x in masraf for x in ("bankamatik", "atm")):
            score += 55

    if "kredi karti faiz" in kategori:
        score -= 200
    if "paket" in masraf or "ticari" in kategori:
        score -= 90

    return score


def _exact_service_score(row, service: str, wanted_channel: str, target_masraf: str) -> int:
    if not _has_published_fee(row):
        return -10_000

    tags = _primary_service_tags(row)
    if service not in tags:
        return -10_000

    _, masraf, full = _primary_row_text(row)
    channels = _primary_channels(row)

    # "Özel/Süper kasa" statüsüne Büyük/Orta/Küçük kasa ücreti taşınmasın.
    target_norm = _norm(target_masraf)
    if service == "KASA" and any(x in target_norm for x in ("ozel", "super", "xl", "extra buyuk")):
        if not any(x in full for x in ("ozel", "super", "xl", "extra buyuk")):
            return -10_000

    score = 140
    if wanted_channel in channels:
        score += 80
    elif wanted_channel == "GENEL" and "GENEL" in channels:
        score += 55
    else:
        # Genel bir satırı özel kanal tarifesi diye kullanma. Bu fallback yalnız
        # Aidat/Okul/Telefon için aşağıdaki generic Fatura/Kurum kuralında yapılır.
        return -10_000

    # Doğrudan MASRAF adındaki hizmet ifadesi açıklamadaki tesadüfi eşleşmeden güçlüdür.
    direct_tokens = {
        "TELEFON": ("telefon", "gsm", "turkcell", "vodafone"),
        "AIDAT": ("aidat", "site", "apartman"),
        "OZEL_OKUL": ("ozel okul", "okul odeme", "egitim odeme"),
        "VERGI": ("vergi", "mtv", "harc"),
        "SGK": ("sgk", "sosyal guvenlik"),
        "SANS_OYUNU": ("sans oyun", "bilyoner", "nesine", "misli", "tjk"),
        "CEK_RISK": ("cek risk", "cek raporu", "kkb cek", "findeks cek"),
        "FAST": ("fast",),
        "YURT_DISI_FAST": ("yurt disi fast", "global fast"),
        "VISA_YP_DIRECT": ("visa direct", "visa ile yurt disi", "visa ile yurtdisi"),
        "DUZENLI_EFT": ("duzenli eft",),
        "DUZENLI_HAVALE": ("duzenli havale",),
        "ALTIN_TRANSFER": ("altin transfer", "ats ile altin", "kiymetli maden transfer"),
        "KASA": ("kiralik kasa", "kasa kiralama", "kasa ucreti"),
    }
    if any(x in masraf for x in direct_tokens.get(service, ())):
        score += 30

    return score


def _fee_signature(row) -> Tuple[str, str, str, str]:
    return tuple(
        _clean(getattr(row, attr, ""))
        for attr in ("asgari_tutar", "asgari_oran", "azami_tutar", "azami_oran")
    )


def _copy_primary_fee(source, target: SupplementalRow, *, generic: bool) -> None:
    for attr in ("asgari_tutar", "asgari_oran", "azami_tutar", "azami_oran"):
        setattr(target, attr, _clean(getattr(source, attr, "")))

    source_date = _clean(
        getattr(source, "site_guncelleme_tarihi", "")
        or getattr(source, "guncelleme_tarihi", "")
        or getattr(source, "komisyon_guncelleme_tarihi", "")
    )
    if source_date and not target.site_guncelleme_tarihi:
        target.site_guncelleme_tarihi = source_date

    # Artık satırda gerçekten resmî bir numeric tarife var. Böylece
    # update_comparison.py DISPLAY_TEXT yerine numeric tutarı gösterir.
    if STATUS_AVAILABLE in target.aciklama:
        target.aciklama = target.aciklama.replace(STATUS_AVAILABLE, STATUS_NUMERIC, 1)

    source_name = _clean(getattr(source, "masraf", ""))
    source_type = "genel Fatura/Kurum tarifesi" if generic else "aynı hizmet tarifesi"
    target.aciklama += (
        " | [SUPPLEMENTAL][FEE_BACKFILLED_FROM_PRIMARY] "
        f"Tutar/oran bankanın ana resmî ücret tablosundaki '{source_name}' satırından "
        f"({source_type}) taşındı; yeni ücret hesaplanmadı."
    )


def _fill_supplemental_fees(
    bank: str,
    additions: Sequence[SupplementalRow],
    primary_rows: Sequence,
) -> Tuple[int, int]:
    """Hizmet-durumu ek satırlarına güvenli primary numeric tarife taşır."""
    filled = 0
    unresolved = 0

    for target in additions:
        if _norm(getattr(target, "kategori", "")) != _norm("EK KAYNAK - Hizmet Durumu"):
            continue
        if _has_published_fee(target):
            continue

        note = _clean(getattr(target, "aciklama", ""))
        if STATUS_NOT_APPLICABLE in note or STATUS_EMPTY in note:
            continue
        if STATUS_AVAILABLE not in note:
            continue

        service = _status_meta_value(target, "SERVICE")
        channel = _status_meta_value(target, "CHANNEL")
        band = _status_meta_value(target, "BAND")
        if not service or not channel:
            unresolved += 1
            continue

        # Aidat / özel okul / telefon satırları yalnız hizmet-kanal kanıtıdır.
        # Genel Fatura/Kurum tarifesini bu supplemental satırlara numeric ücret
        # gibi taşımıyoruz. Karşılaştırma katmanı gerekiyorsa genel tarifeyi
        # açıkça "genel tarife" etiketiyle ayrıca gösterir.
        if service in {"AIDAT", "OZEL_OKUL", "TELEFON", "FATURA_KREDI_KARTI", "FATURA_HESAPTAN"}:
            continue

        # Band statüsü varsa ve uygulanabilir bir satırsa yalnız aynı banda ait
        # primary tarife kullanılmalı. Şu an TRANSFER_3 statüleri NOT_APPLICABLE
        # olduğu için yukarıda zaten atlanıyor; bu guard gelecekteki değişiklikler için.
        target_band_text = _norm(getattr(target, "masraf", ""))

        candidates = []
        for row in primary_rows:
            exact_score = _exact_service_score(row, service, channel, getattr(target, "masraf", ""))
            if exact_score > -10_000:
                if band and band == "TRANSFER_3":
                    row_text = _primary_row_text(row)[2]
                    if not (
                        any(x in row_text for x in ("399.000,01", "399000,01", "399000.01"))
                        and any(x in row_text for x in ("uzeri", "ustu"))
                    ):
                        exact_score = -10_000
                if exact_score > -10_000:
                    candidates.append((exact_score, False, row))


        if not candidates:
            unresolved += 1
            continue

        candidates.sort(
            key=lambda item: (
                item[0],
                0 if item[1] else 1,  # eşit puanda exact hizmet generic'ten önce
                -len(_norm(getattr(item[2], "masraf", ""))),
            ),
            reverse=True,
        )

        best_score, best_generic, best = candidates[0]
        best_fee = _fee_signature(best)

        # Aynı güçlü puanda birden fazla farklı ücret varsa rastgele seçim yapma.
        conflict = any(
            score == best_score and _fee_signature(row) != best_fee
            for score, _, row in candidates[1:]
        )
        if conflict:
            unresolved += 1
            continue

        _copy_primary_fee(best, target, generic=best_generic)
        filled += 1

    return filled, unresolved


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
        "AKBANK", "AIDAT", "Aidat Ödemeleri", ("MOBIL",), AKBANK_PAYMENT_CENTER_URL,
        "Ödeme Merkezi sayfasında Site Aidat Ödemeleri ve Üyelik Aidat Ödemeleri Akbank Mobil hizmetleri olarak yayımlanıyor.",
    )
    rows += _add_service_status(
        "AKBANK", "OZEL_OKUL", "Özel Okul / Eğitim Ödemeleri", ("MOBIL",), AKBANK_PAYMENT_CENTER_URL,
        "Ödeme Merkezi sayfasında Eğitim Ödemeleri Akbank Mobil üzerinden sunulan ödeme hizmetleri arasında yayımlanıyor.",
    )
    return rows


def _yk_service_rows() -> List[SupplementalRow]:
    html = _fetch_html(YAPIKREDI_REGULAR_URL, must_contain=("aidat", "okul taksiti"))
    text = _norm(BeautifulSoup(html, "html.parser").get_text(" ", strip=True))
    if "aidat" not in text or "okul taksiti" not in text:
        raise SupplementalSourceError("Yapı Kredi Düzenli Ödemeler sayfasında aidat/okul taksiti bulunamadı.")

    rows = []
    rows += _add_service_status(
        "YAPIKREDI", "AIDAT", "Aidat Ödemeleri", ("MOBIL",), YAPIKREDI_REGULAR_URL,
        "Düzenli Ödemeler sayfasında aidat açıkça listeleniyor ve talimatın Yapı Kredi Mobil / İnternet Şubesi üzerinden verilebildiği belirtiliyor.",
    )
    rows += _add_service_status(
        "YAPIKREDI", "OZEL_OKUL", "Özel Okul / Okul Taksiti", ("MOBIL",), YAPIKREDI_REGULAR_URL,
        "Düzenli Ödemeler sayfasında okul taksiti açıkça listeleniyor ve talimatın Yapı Kredi Mobil / İnternet Şubesi üzerinden verilebildiği belirtiliyor.",
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
    if not (
        "vadesiz hesap" in bill_text
        and "fatura odem" in bill_text
        and "herhangi bir ucret alinmaz" in bill_text
    ):
        raise SupplementalSourceError(
            "Yapı Kredi Mobil/İnternet vadesiz hesaptan fatura ödemesinin ücretsiz olduğu doğrulanamadı."
        )

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
        "YAPIKREDI",
        "FATURA_HESAPTAN",
        "Hesaptan Fatura / Kurum Ödemesi",
        ("MOBIL",),
        YAPIKREDI_BILL_URL,
        "Resmî fatura sayfası Yapı Kredi Mobil ve Bireysel İnternet Şubesi'nden "
        "vadesiz hesap kullanılarak yapılan fatura ödemelerinde ücret alınmadığını belirtiyor.",
        display_text="Ücretsiz / 0 TRY\\nVadesiz hesaptan",
    )
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
    tax = _fetch_html(
        ISBANK_TAX_URL,
        must_contain=("İşCep", "İnternet Şubesi", "Çözüm Merkezi"),
    )
    tax_text = _norm(BeautifulSoup(tax, "html.parser").get_text(" ", strip=True))
    if "vergi" not in tax_text:
        raise SupplementalSourceError("İş Bankası Vergi Ödeme sayfası doğrulanamadı.")

    findeks = _fetch_html(
        ISBANK_FINDEKS_URL,
        must_contain=("Çek Raporu", "Yılda 50 Çek Raporu", "Yılda 250 Çek Raporu"),
    )
    findeks_text = _clean(BeautifulSoup(findeks, "html.parser").get_text(" ", strip=True))

    def package_amount(count: int, fallback: str) -> str:
        match = re.search(
            rf"Yılda\s*{count}\s*Çek\s*Raporu.{{0,140}}?([0-9][0-9.]*,[0-9]{{2}})\s*TL",
            findeks_text,
            flags=re.I | re.S,
        )
        return match.group(1) if match else fallback

    package_50 = package_amount(50, "3.660,00")
    package_250 = package_amount(250, "15.240,00")

    rows: List[SupplementalRow] = []

    # Vergi sayfası kanalları doğruluyor; ayrı bir numeric "Vergi Tahsilat
    # Komisyonu" ürünü uydurulmuyor. Karşılaştırma bu statüyü kullanır.
    rows += _add_service_status(
        "İŞBANKASI",
        "VERGI",
        "Vergi / Harç Ödemeleri",
        ("MOBIL", "SUBE"),
        ISBANK_TAX_URL,
        "Vergi sayfası İşCep, İnternet Şubesi ve Çözüm Merkezi kanallarını doğruluyor.",
        display_text="Hizmet var\nAyrı vergi aracılık ücreti yayımlanmıyor",
    )

    # İş Bankası Çek Raporu tek-rapor fiyatı yerine yıllık paket olarak
    # yayımlanıyor. Kredi Risk Raporu'nun 95 TL tarifesi buraya taşınmaz.
    # Karekodlu Çek Raporu da farklı bir üründür ve bu satıra karıştırılmaz.
    rows += _add_service_status(
        "İŞBANKASI",
        "CEK_RISK",
        "Findeks Çek Raporu",
        ("MOBIL", "SUBE"),
        ISBANK_FINDEKS_URL,
        "Findeks sayfası Çek Raporu için tek rapor yerine yıllık paket tarifeleri yayımlıyor.",
        display_text=(
            "Çek Raporu paketleri\n"
            f"50 rapor/yıl: {package_50} TRY\n"
            f"250 rapor/yıl: {package_250} TRY\n"
            "KDV dahil"
        ),
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



def _isbank_fatura_faq_rows() -> List[SupplementalRow]:
    """
    İş Bankası'nın resmî SSS sayfasındaki kredi kartından anında fatura ödeme
    ücretini doğrular.

    Kullanıcı denetiminde esas alınan kural:
      0-150 TL       -> 5 TL
      150 TL üzeri   -> işlem tutarının %3,5'i + BSMV

    SSS metni bu üç kritik unsuru içermiyorsa status üretmeyiz; kritik kaynak
    kontrolü başarısız olur ve final Excel korunur.
    """
    html = _fetch_html(
        ISBANK_FAQ_URL,
        must_contain=(
            "Kredi kartlarından Fatura Ödemelerinde ücret alınıyor mu",
            "0-150 TL",
            "150 TL üzeri",
        ),
    )
    text = _norm(BeautifulSoup(html, "html.parser").get_text(" ", strip=True))

    has_question = "kredi kartlarindan fatura odemelerinde ucret aliniyor mu" in text
    has_low = "0-150 tl" in text and "5 tl" in text
    has_high = (
        "150 tl uzeri" in text
        and ("3,5" in text or "3.5" in text)
        and "bsmv" in text
    )
    if not (has_question and has_low and has_high):
        raise SupplementalSourceError(
            "İş Bankası SSS kredi kartından fatura ödeme ücret kuralı doğrulanamadı."
        )

    return _add_service_status(
        "İŞBANKASI",
        "FATURA_KREDI_KARTI",
        "Kredi Kartından Fatura / Kurum Ödemesi",
        ("GENEL",),
        ISBANK_FAQ_URL,
        "Resmî SSS, kredi kartıyla anında fatura ödemelerinde 0-150 TL için 5 TL; "
        "150 TL üzeri için işlem tutarının %3,5'i kadar ücret + BSMV uygulandığını belirtiyor.",
        display_text=(
            "Genel kredi kartı tarifesi:\\n"
            "0-150 TRY: 5 TRY\\n"
            "150 TRY üzeri: %3,50 + BSMV\\n"
            "Kanal ayrımı yayımlanmıyor"
        ),
    )


def _isbank_phone_rows() -> List[SupplementalRow]:
    """İş Bankası telefon faturası ödeme hizmetinin kanal varlığını doğrular."""
    html = _fetch_html(ISBANK_BILL_URL, must_contain=("fatura",))
    text = _norm(BeautifulSoup(html, "html.parser").get_text(" ", strip=True))
    if not any(x in text for x in ("telefon", "cep telefonu", "gsm", "turkcell", "vodafone")):
        raise SupplementalSourceError("İş Bankası Fatura Ödemeleri sayfasında telefon faturası doğrulanamadı.")

    return _add_service_status(
        "İŞBANKASI", "TELEFON", "Telefon / Cep Telefonu Faturası Ödemeleri",
        ("MOBIL", "SUBE"), ISBANK_BILL_URL,
        "Fatura Ödemeleri sayfası telefon faturalarının İşCep/İnternet, Çözüm Merkezi ve şubelerden ödenebildiğini doğruluyor. Telefon yükleme ücreti telefon faturası ücreti olarak kullanılmaz; genel Fatura Ödemeleri tarifesi varsa karşılaştırma katmanında o tarife kullanılır.",
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
    """
    İş Bankası'nın 03.02.2026 tarihli resmî kart sözleşmesinden:
      - Ortak ATM cari hesaba para yatırma tarifesini,
      - Visa Direct'in sözleşmedeki güncel yayın durumunu
    alır.

    Bir alan parse edilemezse diğer doğrulanmış alan yine kullanılabilir.
    """
    import pdfplumber

    pdf_bytes = _request(ISBANK_CARD_CURRENT_URL, binary=True, timeout=60)
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)

    flat = " ".join(text.split())
    norm = _norm(flat)
    rows: List[SupplementalRow] = []

    # ORTAK ATM - Cari Hesaba Para Yatırma
    m = re.search(
        r"Cari\s+Hesaba\s+Para\s+Yat[ıi]rma.{0,300}?%?\s*1[,\.]15\s*"
        r"\+?\s*1[,\.]05\s*TL",
        flat,
        flags=re.I | re.S,
    )
    if not m:
        m = re.search(
            r"cari hesaba para yatirma.{0,300}?1[,\.]15.{0,100}?1[,\.]05\s*tl",
            norm,
            flags=re.I | re.S,
        )

    if m:
        rows.append(
            SupplementalRow(
                kategori="EK KAYNAK - İş Bankası Kart Sözleşmesi 03.02.2026",
                masraf="Ortak ATM - Cari Hesaba Para Yatırma - TEK ATM Değil",
                asgari_tutar="1,05 TL",
                asgari_oran="1,15%",
                azami_oran="1,15%",
                aciklama=_source_note(
                    STATUS_NUMERIC,
                    ISBANK_CARD_CURRENT_URL,
                    "SERVICE=PARA_YATIRMA; CHANNEL=GENEL; Vergi dahildir; "
                    "TEK ATM tarifesi ayrıca %1,15 + 1,58 TL'dir.",
                ),
                site_guncelleme_tarihi="03.02.2026",
            )
        )

    # VISA Direct - sözleşmedeki açık yayın durumu
    if "visa direct hizmeti henuz uygulamada olmayip" in norm:
        rows += _add_service_status(
            "İŞBANKASI",
            "VISA_YP_DIRECT",
            "Visa Direct",
            ("MOBIL", "SUBE"),
            ISBANK_CARD_CURRENT_URL,
            "03.02.2026 tarihli resmî kart sözleşmesinde Visa Direct hizmetinin "
            "henüz uygulamada olmadığı ve devreye alındığında internet sitesinden "
            "duyurulacağı belirtiliyor.",
            marker=STATUS_NOT_APPLICABLE,
            display_text="03.02.2026 sözleşmesinde\\nhenüz uygulamada değil",
        )

    if not rows:
        raise SupplementalSourceError(
            "İş Bankası 03.02.2026 kart sözleşmesinden beklenen karşılaştırma verileri parse edilemedi."
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



def _yk_altin_transfer_status_rows() -> List[SupplementalRow]:
    html = _fetch_html(YAPIKREDI_FEE_URL, must_contain=("Külçe Altın Çekme",))
    page = _norm(BeautifulSoup(html, "html.parser").get_text(" ", strip=True))
    if "kulce altin cekme" not in page:
        raise SupplementalSourceError("Yapı Kredi altın/fiziki teslim bölümü doğrulanamadı.")

    return _add_service_status(
        "YAPIKREDI", "ALTIN_TRANSFER", "Elektronik Altın / Altın Transferi",
        ("GENEL",), YAPIKREDI_FEE_URL,
        "Ürün ve Hizmet Ücretleri sayfasında Külçe Altın Çekme/fiziki altın ücreti yayımlanıyor; "
        "ayrı elektronik bankalararası altın transfer tarifesi yayımlanmıyor.",
        display_text="Ayrı elektronik altın transfer\\ntarifesi yayımlanmıyor",
    )


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
        ("Karekodlu ve Logolu Çek Defteri - 10 Yapraklı (Yurtiçi Şubeler)", r"Karekodlu ve Logolu Çek Defteri(?: Ücreti)?\*?.{0,120}?10 Yapraklı \(Yurtiçi Şubeler\)"),
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
        "karekodlu çek defteri 10 yaprak": any(
            "karekodlu cek defteri - 10 yaprak" in _norm(r.masraf)
            and "logolu" not in _norm(r.masraf)
            and _clean(r.asgari_tutar).startswith("750")
            for r in rows
        ),
        "logolu çek defteri 10 yaprak": any(
            "logolu cek defteri - 10 yaprak" in _norm(r.masraf)
            and _clean(r.asgari_tutar).startswith("900")
            for r in rows
        ),
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


def _drop_obsolete_isbank_bhs_rows(
    additions: Sequence[SupplementalRow], primary_rows: Sequence,
) -> Tuple[List[SupplementalRow], int]:
    """Güncel ana tabloda yayımlanan YP çek ücretini eski BHS ile gölgelemez."""
    has_current_yp_cheque = any(
        _norm(getattr(row, "masraf", "")) == "diger banka - tahsile alinan - yp"
        and _has_published_fee(row)
        for row in primary_rows
    )
    if not has_current_yp_cheque:
        return list(additions), 0

    filtered = [
        row for row in additions
        if _norm(row.masraf) != "diger banka yp - tahsile alinan"
    ]
    return filtered, len(additions) - len(filtered)


def enrich_all(banka_verileri: Mapping[str, Sequence]) -> Tuple[Dict[str, List], SupplementalReport]:
    """
    Primary scraper sonuçlarını kopyalar ve ek resmî kaynak satırlarını ekler.
    Her kritik ek kaynak doğrulanamazsa report.ok=False döner; main.py bu durumda
    mevcut doğru Excel'i korumalıdır.
    """
    report = SupplementalReport()

    # ORIGINAL primary snapshot ayrı tutulur. Supplemental satırların ücreti başka
    # supplemental satırlardan değil, yalnız ana resmî scraper sonucundan tamamlanır.
    primary_only: Dict[str, List] = {
        bank: list(rows)
        for bank, rows in banka_verileri.items()
    }
    result: Dict[str, List] = {
        bank: list(rows)
        for bank, rows in banka_verileri.items()
    }

    def apply(bank: str, source_name: str, url: str, producer, *, required: bool = True):
        print(
            f"[supplemental] başlıyor: {source_name} | "
            f"{'KRİTİK' if required else 'opsiyonel'}"
        )
        if bank not in result:
            if required:
                report.fail(f"{source_name}: {bank} primary verisi yok.")
            return
        try:
            additions = producer()

            filled, unresolved = _fill_supplemental_fees(
                bank,
                additions,
                primary_only.get(bank, ()),
            )
            if filled:
                print(
                    f"[supplemental] {source_name}: {filled} hizmet satırına "
                    "primary resmî tarifeden tutar/oran yazıldı."
                )
            if unresolved:
                print(
                    f"[supplemental] {source_name}: {unresolved} hizmet satırında "
                    "güvenli numeric tarife eşleşmesi yok; ücret uydurulmadı."
                )

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
            rows, skipped = _drop_obsolete_isbank_bhs_rows(
                rows,
                primary_only.get("İŞBANKASI", ()),
            )
            if skipped:
                print(
                    "[supplemental] İŞBANKASI_BHS: güncel primary YP çek "
                    f"tarifesi nedeniyle eski satır atlandı={skipped}"
                )
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
    apply("İŞBANKASI", "ISBANK_FATURA_KART_SSS", ISBANK_FAQ_URL, _isbank_fatura_faq_rows)

    # Özel okul / aidat hizmet kanıtları da karşılaştırma mantığının parçasıdır.
    apply("GARANTİ", "GARANTI_OZEL_OKUL", GARANTI_SCHOOL_URL, _garanti_service_rows)
    apply("İŞBANKASI", "ISBANK_HIZMETLER", ISBANK_SCHOOL_URL, _isbank_service_rows)

    # Karşılaştırma doğruluğu için kanal/limit/publikasyon statüleri.
    apply("GARANTİ", "GARANTI_COMPARISON_POLICY", GARANTI_FEE_URL, _garanti_comparison_policy_rows)
    apply("AKBANK", "AKBANK_COMPARISON_POLICY", AKBANK_FAST_URL, _akbank_comparison_policy_rows)
    apply("İŞBANKASI", "ISBANK_FAST_SGK_POLICY", ISBANK_FAST_URL, _isbank_fast_sgk_policy_rows)
    apply("YAPIKREDI", "YAPIKREDI_COMPARISON_POLICY", YAPIKREDI_FAST_URL, _yk_comparison_policy_rows)
    def yk_altin_status_from_primary():
        primary_rows = primary_only.get("YAPIKREDI", [])
        primary_text = " | ".join(
            _norm(
                f"{getattr(row, 'kategori', '')} "
                f"{getattr(row, 'masraf', '')} "
                f"{getattr(row, 'aciklama', '')}"
            )
            for row in primary_rows
        )

        # Resmî ana ücret sayfasından gelen primary veride fiziksel külçe altın
        # satırı yoksa bu status'u üretmeyiz; böylece tahmin yapılmaz.
        if "kulce altin cekme" not in primary_text:
            raise SupplementalSourceError(
                "Primary Yapı Kredi verisinde 'Külçe Altın Çekme' satırı bulunamadı."
            )

        # Primary veride gerçekten elektronik/bankalararası altın transfer
        # tarifesi varsa status eklemeye gerek yok; numeric satır kullanılacaktır.
        electronic_tokens = (
            "altin transfer",
            "kiymetli maden transferi ucreti - altin",
            "ats ile altin gonderimi",
        )
        if any(token in primary_text for token in electronic_tokens):
            return []

        return _add_service_status(
            "YAPIKREDI",
            "ALTIN_TRANSFER",
            "Elektronik Altın / Altın Transferi",
            ("GENEL",),
            YAPIKREDI_FEE_URL,
            "Primary resmî ücret verisinde Külçe Altın Çekme/fiziki altın tarifesi "
            "bulunuyor; ayrı elektronik bankalararası altın transfer tarifesi "
            "tespit edilmedi.",
            display_text=(
                "Ayrı elektronik altın transfer\n"
                "tarifesi yayımlanmıyor"
            ),
        )

    # Bu satır karşılaştırmayı zenginleştirir fakat ana finansal veri değildir.
    # Site HTML'i/başlığı değişirse bütün pipeline'ı bloke etmemeli.
    apply(
        "YAPIKREDI",
        "YAPIKREDI_ALTIN_TRANSFER_STATUS",
        YAPIKREDI_FEE_URL,
        yk_altin_status_from_primary,
        required=False,
    )

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
