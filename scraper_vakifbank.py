"""
VakıfBank "Ürün ve Hizmet Ücretleri" scraper.

Bu sürüm:
- VakıfBank sayfasının çağırdığı getProductServicePrices API yanıtını Playwright ile yakalar.
- İlk yakalanan yanıtı körlemesine kullanmaz; tüm aday API cevaplarını inceler.
- JSON yapısı değişse bile Fee/fee altında veya iç içe listelerde ücret listesini bulmaya çalışır.
- API item anahtarlarını loglar; şema değişikliğini erken fark ettirir.
- Ana kategori + varsa alt grup/kanal + ücret adını korur.
- "Uluslararası Fon Transferi" kalemlerini SWIFT filtresinde görünür yapar.
- FAST / EFT / Havale / SWIFT kontrolü yapar.
- Duplicate, eksik isim, boş kayıt ve tarih alanlarını raporlar.
- 30 saniyelik sabit beklemeyi kısaltır; API gelmezse bir kez kontrollü reload dener.
- main.py / update_excel.py ile uyumlu UcretSatiri listesi döndürür.
"""

import json
import re
import sys
import time
from dataclasses import dataclass
from typing import Dict, List, Optional, Set, Tuple


SCRAPER_VERSION = "2026-08-19-v3-vakifbank-channel-currency-fix"

VAKIFBANK_API_URL = (
    "https://inbound.apigateway.vakifbank.com.tr:8443/"
    "getProductServicePrices"
)

VAKIFBANK_PAGE_URL = (
    "https://www.vakifbank.com.tr/tr/urun-ve-hizmet-ucretleri"
)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "tr-TR,tr;q=0.9,en;q=0.8",
}


# =========================================================
# VERİ YAPISI
# =========================================================

@dataclass
class UcretSatiri:
    kategori: str
    masraf: str
    asgari_tutar: str = ""
    asgari_oran: str = ""
    azami_tutar: str = ""
    azami_oran: str = ""
    aciklama: str = ""
    site_guncelleme_tarihi: str = ""


class ScraperError(Exception):
    pass


# =========================================================
# NORMALİZASYON
# =========================================================

def _normalize(value) -> str:
    if value is None:
        return ""

    text = str(value)
    text = text.replace("\xa0", " ")
    text = text.replace("\u200b", "")
    text = text.replace("\r", " ")
    text = text.replace("\n", " ")
    text = re.sub(r"\s+", " ", text)

    return text.strip()


def _normalize_key(value) -> str:
    text = _normalize(value).lower()

    replacements = {
        "ı": "i",
        "ğ": "g",
        "ü": "u",
        "ş": "s",
        "ö": "o",
        "ç": "c",
    }

    for old, new in replacements.items():
        text = text.replace(old, new)

    text = re.sub(r"\s+", " ", text)

    return text.strip()


def _normalize_tutar(value) -> str:
    """
    API'deki gerçek 0 değerlerini boş bırakır.
    '0,50' gibi sıfır olmayan oranları korur.
    """
    if value is None:
        return ""

    if isinstance(value, bool):
        return ""

    if isinstance(value, (int, float)):
        if value == 0:
            return ""

        if isinstance(value, float) and value.is_integer():
            return str(int(value))

        return str(value)

    text = _normalize(value)

    if not text:
        return ""

    if text in {
        "0",
        "0.0",
        "0.00",
        "0,0",
        "0,00",
        "-",
    }:
        return ""

    try:
        numeric = float(
            text.replace(".", "").replace(",", ".")
            if "," in text
            else text
        )

        if numeric == 0:
            return ""

    except Exception:
        pass

    return text


def _normalize_tarih(value) -> str:
    """
    Çeşitli VakıfBank tarih formatlarını:
        dd.mm.yyyy HH:MM
    biçimine getirir.
    """
    text = _normalize(value)

    if not text:
        return ""

    # 2026-04-28T10:10:41 / Z / +03:00
    match = re.match(
        r"^(\d{4})-(\d{2})-(\d{2})"
        r"[T\s](\d{2}):(\d{2})",
        text,
    )

    if match:
        return (
            f"{match.group(3)}."
            f"{match.group(2)}."
            f"{match.group(1)} "
            f"{match.group(4)}:"
            f"{match.group(5)}"
        )

    # dd.mm.yyyyHH:MM
    match = re.match(
        r"^(\d{1,2}[./]\d{1,2}[./]\d{4})"
        r"(\d{2}:\d{2})",
        text,
    )

    if match:
        return (
            match.group(1).replace("/", ".")
            + " "
            + match.group(2)
        )

    # dd.mm.yyyy HH:MM:SS
    match = re.match(
        r"^(\d{1,2}[./]\d{1,2}[./]\d{4})"
        r"\s+(\d{2}:\d{2}):\d{2}",
        text,
    )

    if match:
        return (
            match.group(1).replace("/", ".")
            + " "
            + match.group(2)
        )

    # dd.mm.yyyy HH:MM
    match = re.match(
        r"^(\d{1,2}[./]\d{1,2}[./]\d{4})"
        r"\s+(\d{2}:\d{2})",
        text,
    )

    if match:
        return (
            match.group(1).replace("/", ".")
            + " "
            + match.group(2)
        )

    # dd.mm.yyyy
    match = re.match(
        r"^(\d{1,2}[./]\d{1,2}[./]\d{4})$",
        text,
    )

    if match:
        return match.group(1).replace("/", ".")

    return text


# =========================================================
# TRANSFER TERİMLERİ
# =========================================================

TRANSFER_TERMS = (
    "fast",
    "eft",
    "havale",
    "swift",
)


def _has_transfer_term(
    text: str,
    term: str,
) -> bool:
    normalized = _normalize_key(text)

    patterns = {
        "fast": r"(?<![a-z0-9])fast(?![a-z0-9])",
        "eft": r"(?<![a-z0-9])eft(?![a-z0-9])",
        "havale": r"(?<![a-z0-9])havale(?![a-z0-9])",
        "swift": r"(?<![a-z0-9])swift(?![a-z0-9])",
    }

    pattern = patterns.get(term)

    if pattern:
        return re.search(
            pattern,
            normalized,
        ) is not None

    return term in normalized


# =========================================================
# JSON / API ŞEMA BULMA
# =========================================================

FEE_NAME_KEYS = (
    "FeeName",
    "feeName",
    "ItemName",
    "itemName",
    "Name",
    "name",
)

FEE_VALUE_KEYS = (
    "MinimumAmount",
    "minimumAmount",
    "MinimumRate",
    "minimumRate",
    "MaximumAmount",
    "maximumAmount",
    "MaximumRate",
    "maximumRate",
    "Description",
    "description",
    "UpdateDate",
    "updateDate",
)

CATEGORY_KEYS = (
    "MainTransactionGroupName",
    "mainTransactionGroupName",
    "MainGroupName",
    "mainGroupName",
    "CategoryName",
    "categoryName",
)

SUBGROUP_KEYS = (
    "TransactionGroupName",
    "transactionGroupName",
    "SubTransactionGroupName",
    "subTransactionGroupName",
    "FeeGroupName",
    "feeGroupName",
    "GroupName",
    "groupName",
)

CHANNEL_KEYS = (
    "Channel",
    "channel",
    "ChannelName",
    "channelName",
    "TransactionChannelName",
    "transactionChannelName",
)


def _first_value(
    item: dict,
    keys,
):
    for key in keys:
        if key not in item:
            continue

        value = item.get(key)

        if value is None:
            continue

        if isinstance(
            value,
            (dict, list),
        ):
            continue

        text = _normalize(value)

        if text:
            return text

    return ""



def _with_currency(
    amount: str,
    currency: str,
) -> str:
    """
    API CurrencyCode bilgisini tutarlarda kaybetme.
    Örn:
      5 + USD -> USD 5
      10 + TL -> TL 10
    """
    amount = _normalize(amount)
    currency = _normalize(currency)

    if not amount:
        return ""

    if not currency:
        return amount

    amount_key = _normalize_key(amount)
    currency_key = _normalize_key(currency)

    if (
        amount_key.startswith(currency_key + " ")
        or amount_key.endswith(" " + currency_key)
    ):
        return amount

    return f"{currency} {amount}"


def _looks_like_fee_item(
    item,
) -> bool:
    if not isinstance(
        item,
        dict,
    ):
        return False

    has_name = any(
        key in item
        and _normalize(item.get(key))
        for key in FEE_NAME_KEYS
    )

    if not has_name:
        return False

    has_value_key = any(
        key in item
        for key in FEE_VALUE_KEYS
    )

    has_category_key = any(
        key in item
        for key in CATEGORY_KEYS
    )

    return (
        has_value_key
        or has_category_key
    )


def _collect_fee_lists(
    node,
    path: str = "$",
) -> List[
    Tuple[
        str,
        List[dict],
    ]
]:
    """
    Bilinen Data.Fee yapısına ek olarak, şema değişirse
    iç içe JSON'da ücret item'larına benzeyen listeleri bulur.
    """
    candidates: List[
        Tuple[
            str,
            List[dict],
        ]
    ] = []

    if isinstance(
        node,
        dict,
    ):
        for key, value in node.items():
            child_path = (
                f"{path}.{key}"
            )

            if isinstance(
                value,
                list,
            ):
                dict_items = [
                    item
                    for item in value
                    if isinstance(
                        item,
                        dict,
                    )
                ]

                fee_like = [
                    item
                    for item in dict_items
                    if _looks_like_fee_item(
                        item
                    )
                ]

                if fee_like:
                    candidates.append(
                        (
                            child_path,
                            fee_like,
                        )
                    )

            candidates.extend(
                _collect_fee_lists(
                    value,
                    child_path,
                )
            )

    elif isinstance(
        node,
        list,
    ):
        for index, value in enumerate(
            node
        ):
            candidates.extend(
                _collect_fee_lists(
                    value,
                    f"{path}[{index}]",
                )
            )

    return candidates


def _extract_fee_list(
    data,
) -> Tuple[
    List[dict],
    str,
]:
    """
    Önce bilinen VakıfBank şemasını dener.
    Sonra recursive şema keşfine düşer.
    """
    if isinstance(
        data,
        dict,
    ):
        for data_key in (
            "Data",
            "data",
        ):
            block = data.get(
                data_key
            )

            if isinstance(
                block,
                dict,
            ):
                for fee_key in (
                    "Fee",
                    "fee",
                    "Fees",
                    "fees",
                ):
                    fee_list = block.get(
                        fee_key
                    )

                    if isinstance(
                        fee_list,
                        list,
                    ):
                        dict_items = [
                            item
                            for item in fee_list
                            if isinstance(
                                item,
                                dict,
                            )
                        ]

                        if dict_items:
                            return (
                                dict_items,
                                f"$.{data_key}.{fee_key}",
                            )

            elif isinstance(
                block,
                list,
            ):
                dict_items = [
                    item
                    for item in block
                    if isinstance(
                        item,
                        dict,
                    )
                ]

                if any(
                    _looks_like_fee_item(
                        item
                    )
                    for item in dict_items
                ):
                    return (
                        dict_items,
                        f"$.{data_key}",
                    )

        for fee_key in (
            "Fee",
            "fee",
            "Fees",
            "fees",
        ):
            fee_list = data.get(
                fee_key
            )

            if isinstance(
                fee_list,
                list,
            ):
                dict_items = [
                    item
                    for item in fee_list
                    if isinstance(
                        item,
                        dict,
                    )
                ]

                if dict_items:
                    return (
                        dict_items,
                        f"$.{fee_key}",
                    )

    if isinstance(
        data,
        list,
    ):
        dict_items = [
            item
            for item in data
            if isinstance(
                item,
                dict,
            )
        ]

        if any(
            _looks_like_fee_item(
                item
            )
            for item in dict_items
        ):
            return (
                dict_items,
                "$",
            )

    candidates = _collect_fee_lists(
        data
    )

    if not candidates:
        return [], ""

    # En çok ücret item'ı içeren aday.
    candidates.sort(
        key=lambda item: len(
            item[1]
        ),
        reverse=True,
    )

    return candidates[0][1], candidates[0][0]


# =========================================================
# MASRAF HİYERARŞİSİ
# =========================================================

def _append_unique(
    parts: List[str],
    value: str,
) -> None:
    value = _normalize(value)

    if not value:
        return

    value_key = _normalize_key(
        value
    )

    for existing in parts:
        existing_key = (
            _normalize_key(
                existing
            )
        )

        if (
            existing_key == value_key
            or value_key in existing_key
        ):
            return

    parts.append(value)


def _build_masraf(
    item: dict,
    raw_name: str,
) -> str:
    """
    API'de varsa alt işlem grubu ve kanal bilgisini ücret adıyla
    birleştirir. Ana kategori KATEGORİ kolonunda ayrıca tutulur.
    """
    parts: List[str] = []

    subgroup = _first_value(
        item,
        SUBGROUP_KEYS,
    )

    channel = _first_value(
        item,
        CHANNEL_KEYS,
    )

    _append_unique(
        parts,
        subgroup,
    )

    _append_unique(
        parts,
        channel,
    )

    _append_unique(
        parts,
        raw_name,
    )

    masraf = (
        " - ".join(parts)
        if parts
        else raw_name
    )

    combined = _normalize_key(
        " ".join(
            [
                subgroup,
                channel,
                raw_name,
                _first_value(
                    item,
                    CATEGORY_KEYS,
                ),
            ]
        )
    )

    # VakıfBank'ta SWIFT işlemi açıklama/başlık olarak
    # "Uluslararası Fon Transferi" veya "Döviz Transferi"
    # biçiminde yayınlanabilir.
    swift_alias = (
        "uluslararasi fon transfer" in combined
        or "yurtdisi doviz transfer" in combined
        or "yurt disi doviz transfer" in combined
    )

    if (
        swift_alias
        and not _has_transfer_term(
            masraf,
            "swift",
        )
    ):
        masraf = (
            f"SWIFT - {masraf}"
        )

    return _normalize(
        masraf
    )


# =========================================================
# API ITEM PARSER
# =========================================================

def _parse_fee_list(
    fee_list: List[dict],
) -> Tuple[
    List[UcretSatiri],
    Dict[str, int],
]:
    stats: Dict[str, int] = {
        "raw_items": len(fee_list),
        "non_dict": 0,
        "missing_name": 0,
        "empty_record": 0,
        "parsed_before_dedup": 0,
        "duplicates": 0,
        "missing_category": 0,
        "missing_date": 0,
    }

    rows: List[
        UcretSatiri
    ] = []

    seen: Set[
        Tuple[
            str,
            str,
            str,
            str,
            str,
            str,
            str,
            str,
        ]
    ] = set()

    for item in fee_list:
        if not isinstance(
            item,
            dict,
        ):
            stats[
                "non_dict"
            ] += 1
            continue

        raw_name = _first_value(
            item,
            FEE_NAME_KEYS,
        )

        if not raw_name:
            stats[
                "missing_name"
            ] += 1
            continue

        kategori = _first_value(
            item,
            CATEGORY_KEYS,
        )

        if not kategori:
            kategori = "Genel"
            stats[
                "missing_category"
            ] += 1

        currency = _first_value(
            item,
            (
                "CurrencyCode",
                "currencyCode",
                "Currency",
                "currency",
            ),
        )

        asgari_tutar = _with_currency(
            _normalize_tutar(
                item.get(
                    "MinimumAmount",
                    item.get(
                        "minimumAmount",
                        "",
                    ),
                )
            ),
            currency,
        )

        asgari_oran = (
            _normalize_tutar(
                item.get(
                    "MinimumRate",
                    item.get(
                        "minimumRate",
                        "",
                    ),
                )
            )
        )

        azami_tutar = _with_currency(
            _normalize_tutar(
                item.get(
                    "MaximumAmount",
                    item.get(
                        "maximumAmount",
                        "",
                    ),
                )
            ),
            currency,
        )

        azami_oran = (
            _normalize_tutar(
                item.get(
                    "MaximumRate",
                    item.get(
                        "maximumRate",
                        "",
                    ),
                )
            )
        )

        aciklama = _first_value(
            item,
            (
                "Description",
                "description",
                "Explanation",
                "explanation",
            ),
        )

        bsmv = _first_value(
            item,
            (
                "BSMV",
                "bsmv",
            ),
        )

        if bsmv:
            bsmv_text = f"BSMV: {bsmv}"

            if bsmv_text not in aciklama:
                aciklama = (
                    f"{aciklama} | {bsmv_text}"
                    if aciklama
                    else bsmv_text
                )

        site_tarihi = (
            _normalize_tarih(
                item.get(
                    "UpdateDate",
                    item.get(
                        "updateDate",
                        item.get(
                            "LastUpdateDate",
                            item.get(
                                "lastUpdateDate",
                                "",
                            ),
                        ),
                    ),
                )
            )
        )

        if not site_tarihi:
            stats[
                "missing_date"
            ] += 1

        # İsim dışında hiçbir ücret bilgisi yoksa yine kaydı koruyoruz;
        # çünkü bazı tarife satırları yalnız açıklama / grup bilgisi taşıyabilir.
        has_any_detail = any(
            [
                asgari_tutar,
                asgari_oran,
                azami_tutar,
                azami_oran,
                aciklama,
                site_tarihi,
            ]
        )

        if not has_any_detail:
            stats[
                "empty_record"
            ] += 1

        masraf = _build_masraf(
            item,
            raw_name,
        )

        stats[
            "parsed_before_dedup"
        ] += 1

        key = (
            kategori,
            masraf,
            asgari_tutar,
            asgari_oran,
            azami_tutar,
            azami_oran,
            aciklama,
            site_tarihi,
        )

        if key in seen:
            stats[
                "duplicates"
            ] += 1
            continue

        seen.add(key)

        rows.append(
            UcretSatiri(
                kategori=kategori,
                masraf=masraf,
                asgari_tutar=asgari_tutar,
                asgari_oran=asgari_oran,
                azami_tutar=azami_tutar,
                azami_oran=azami_oran,
                aciklama=aciklama,
                site_guncelleme_tarihi=site_tarihi,
            )
        )

    return rows, stats


# =========================================================
# PLAYWRIGHT API INTERCEPT
# =========================================================

def _is_target_api(
    url: str,
) -> bool:
    key = _normalize_key(url)

    return (
        "getproductserviceprices"
        in key
        or "productserviceprice"
        in key
    )


def _capture_api_json(
    page_url: str,
) -> Tuple[
    List[dict],
    List[dict],
]:
    from playwright.sync_api import (
        sync_playwright,
    )

    captured: List[dict] = []
    request_meta: List[
        dict
    ] = []

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True
        )

        try:
            context = (
                browser.new_context(
                    user_agent=HEADERS[
                        "User-Agent"
                    ],
                    viewport={
                        "width": 1440,
                        "height": 1000,
                    },
                    locale="tr-TR",
                    timezone_id=(
                        "Europe/Istanbul"
                    ),
                    extra_http_headers={
                        "Accept-Language": (
                            HEADERS[
                                "Accept-Language"
                            ]
                        ),
                    },
                )
            )

            page = context.new_page()

            def handle_response(
                response,
            ):
                if not _is_target_api(
                    response.url
                ):
                    return

                method = ""

                try:
                    method = (
                        response.request.method
                    )
                except Exception:
                    pass

                print(
                    f"[vakifbank] Hedef API: "
                    f"{method or '?'} "
                    f"{response.url} "
                    f"({response.status})",
                    file=sys.stderr,
                )

                request_meta.append(
                    {
                        "url": response.url,
                        "status": (
                            response.status
                        ),
                        "method": method,
                    }
                )

                if response.status != 200:
                    return

                try:
                    data = response.json()

                    captured.append(
                        {
                            "url": response.url,
                            "data": data,
                        }
                    )

                    print(
                        "[vakifbank] "
                        "JSON API yanıtı yakalandı.",
                        file=sys.stderr,
                    )

                except Exception as exc:
                    try:
                        body = response.body()
                        text = body.decode(
                            "utf-8",
                            errors="replace",
                        )

                        data = json.loads(
                            text
                        )

                        captured.append(
                            {
                                "url": response.url,
                                "data": data,
                            }
                        )

                        print(
                            "[vakifbank] "
                            "API body JSON olarak yakalandı.",
                            file=sys.stderr,
                        )

                    except Exception:
                        print(
                            f"[vakifbank][UYARI] "
                            f"API JSON okunamadı: "
                            f"{exc}",
                            file=sys.stderr,
                        )

            page.on(
                "response",
                handle_response,
            )

            print(
                "[vakifbank] "
                "Ürün ve Hizmet Ücretleri "
                "sayfası yükleniyor...",
                file=sys.stderr,
            )

            page.goto(
                page_url,
                timeout=120000,
                wait_until=(
                    "domcontentloaded"
                ),
            )

            # API çoğu durumda ilk birkaç saniyede gelir.
            deadline = (
                time.time()
                + 12
            )

            while (
                not captured
                and time.time()
                < deadline
            ):
                page.wait_for_timeout(
                    400
                )

            # Gelmediyse sayfadaki olası accordion/tabları tetikle.
            if not captured:
                print(
                    "[vakifbank][UYARI] "
                    "İlk yüklemede API gelmedi; "
                    "dinamik öğeler tetikleniyor.",
                    file=sys.stderr,
                )

                try:
                    page.evaluate("""
                    () => {
                        const targets = new Set();

                        document.querySelectorAll(
                            "[aria-expanded='false'], "
                            "[data-bs-toggle='collapse'], "
                            "[role='button'], "
                            ".accordion-button"
                        ).forEach(el => {
                            targets.add(el);
                        });

                        for (const el of targets) {
                            try {
                                el.click();
                            } catch (_) {}
                        }

                        window.scrollTo(
                            0,
                            document.documentElement
                                ? document.documentElement.scrollHeight
                                : 0
                        );
                    }
                    """)
                except Exception:
                    pass

                deadline = (
                    time.time()
                    + 6
                )

                while (
                    not captured
                    and time.time()
                    < deadline
                ):
                    page.wait_for_timeout(
                        400
                    )

            # Hâlâ gelmediyse yalnız bir kez reload.
            if not captured:
                print(
                    "[vakifbank][UYARI] "
                    "API hâlâ gelmedi; "
                    "sayfa bir kez yenileniyor.",
                    file=sys.stderr,
                )

                try:
                    page.reload(
                        timeout=90000,
                        wait_until=(
                            "domcontentloaded"
                        ),
                    )

                    deadline = (
                        time.time()
                        + 8
                    )

                    while (
                        not captured
                        and time.time()
                        < deadline
                    ):
                        page.wait_for_timeout(
                            400
                        )

                except Exception as exc:
                    print(
                        f"[vakifbank][UYARI] "
                        f"Reload başarısız: "
                        f"{exc}",
                        file=sys.stderr,
                    )

        finally:
            browser.close()

    return captured, request_meta


# =========================================================
# RAPORLAR
# =========================================================

def _print_schema_report(
    fee_list: List[dict],
    json_path: str,
) -> None:
    print(
        "",
        file=sys.stderr,
    )

    print(
        "[vakifbank] ===== API ŞEMA KONTROLÜ =====",
        file=sys.stderr,
    )

    print(
        f"[vakifbank] Fee listesi JSON yolu: "
        f"{json_path or '?'}",
        file=sys.stderr,
    )

    print(
        f"[vakifbank] Ham Fee item sayısı: "
        f"{len(fee_list)}",
        file=sys.stderr,
    )

    keys: Set[str] = set()

    for item in fee_list[:50]:
        if isinstance(
            item,
            dict,
        ):
            keys.update(
                str(k)
                for k in item.keys()
            )

    print(
        "[vakifbank] İlk item'larda görülen alanlar: "
        + ", ".join(
            sorted(keys)
        ),
        file=sys.stderr,
    )

    channels = sorted({
        _first_value(
            item,
            CHANNEL_KEYS,
        )
        for item in fee_list
        if isinstance(item, dict)
        and _first_value(
            item,
            CHANNEL_KEYS,
        )
    })

    currencies = sorted({
        _first_value(
            item,
            (
                "CurrencyCode",
                "currencyCode",
                "Currency",
                "currency",
            ),
        )
        for item in fee_list
        if isinstance(item, dict)
        and _first_value(
            item,
            (
                "CurrencyCode",
                "currencyCode",
                "Currency",
                "currency",
            ),
        )
    })

    print(
        f"[vakifbank] Kanal çeşitleri "
        f"({len(channels)}): "
        + ", ".join(channels[:30]),
        file=sys.stderr,
    )

    print(
        f"[vakifbank] Para birimleri "
        f"({len(currencies)}): "
        + ", ".join(currencies[:30]),
        file=sys.stderr,
    )

    print(
        "[vakifbank] =============================",
        file=sys.stderr,
    )


def _print_category_report(
    rows: List[UcretSatiri],
) -> None:
    counts: Dict[
        str,
        int,
    ] = {}

    for row in rows:
        counts[row.kategori] = (
            counts.get(
                row.kategori,
                0,
            )
            + 1
        )

    print(
        "",
        file=sys.stderr,
    )

    print(
        "[vakifbank] ===== KATEGORİ RAPORU =====",
        file=sys.stderr,
    )

    for category, count in sorted(
        counts.items(),
        key=lambda item: (
            -item[1],
            item[0],
        ),
    ):
        print(
            f"[vakifbank] "
            f"{category} -> "
            f"{count} kayıt",
            file=sys.stderr,
        )

    print(
        "[vakifbank] ===========================",
        file=sys.stderr,
    )


def _print_integrity_report(
    stats: Dict[str, int],
    result_count: int,
) -> None:
    raw_items = stats.get(
        "raw_items",
        0,
    )

    accounted = (
        stats.get(
            "parsed_before_dedup",
            0,
        )
        + stats.get(
            "non_dict",
            0,
        )
        + stats.get(
            "missing_name",
            0,
        )
    )

    print(
        "",
        file=sys.stderr,
    )

    print(
        "[vakifbank] ===== BÜTÜNLÜK KONTROLÜ =====",
        file=sys.stderr,
    )

    print(
        f"[vakifbank] Ham API item: "
        f"{raw_items}",
        file=sys.stderr,
    )

    print(
        f"[vakifbank] Parse edilen "
        f"(dedup öncesi): "
        f"{stats.get('parsed_before_dedup', 0)}",
        file=sys.stderr,
    )

    print(
        f"[vakifbank] Duplicate: "
        f"{stats.get('duplicates', 0)}",
        file=sys.stderr,
    )

    print(
        f"[vakifbank] Dict olmayan item: "
        f"{stats.get('non_dict', 0)}",
        file=sys.stderr,
    )

    print(
        f"[vakifbank] Ücret adı olmayan item: "
        f"{stats.get('missing_name', 0)}",
        file=sys.stderr,
    )

    print(
        f"[vakifbank] Ana kategori olmayan item: "
        f"{stats.get('missing_category', 0)}",
        file=sys.stderr,
    )

    print(
        f"[vakifbank] Tarih olmayan item: "
        f"{stats.get('missing_date', 0)}",
        file=sys.stderr,
    )

    print(
        f"[vakifbank] Değer/açıklama/tarih boş item: "
        f"{stats.get('empty_record', 0)}",
        file=sys.stderr,
    )

    print(
        f"[vakifbank] Excel'e gidecek "
        f"benzersiz satır: "
        f"{result_count}",
        file=sys.stderr,
    )

    if (
        raw_items == accounted
        and stats.get(
            "missing_name",
            0,
        ) == 0
    ):
        print(
            "[vakifbank] BÜTÜNLÜK: OK - "
            "API item'larının tamamı açıklandı.",
            file=sys.stderr,
        )
    else:
        print(
            "[vakifbank] BÜTÜNLÜK: UYARI",
            file=sys.stderr,
        )

    print(
        "[vakifbank] ===============================",
        file=sys.stderr,
    )


def _print_transfer_report(
    rows: List[UcretSatiri],
) -> None:
    print(
        "",
        file=sys.stderr,
    )

    print(
        "[vakifbank] ===== PARA AKTARMA KONTROLÜ =====",
        file=sys.stderr,
    )

    for label, term in [
        ("FAST", "fast"),
        ("EFT", "eft"),
        ("Havale", "havale"),
        ("SWIFT", "swift"),
    ]:
        found = [
            row
            for row in rows
            if _has_transfer_term(
                row.masraf,
                term,
            )
        ]

        print(
            f"[vakifbank] {label}: "
            f"{len(found)} kayıt",
            file=sys.stderr,
        )

        for row in found[:12]:
            print(
                f"    - "
                f"[{row.kategori}] "
                f"{row.masraf}",
                file=sys.stderr,
            )

        if not found:
            print(
                f"[vakifbank][UYARI] "
                f"{label} MASRAF alanında "
                "hiç bulunamadı.",
                file=sys.stderr,
            )

    print(
        "[vakifbank] =================================",
        file=sys.stderr,
    )


def _print_extra_product_report(
    rows: List[UcretSatiri],
) -> None:
    print(
        "",
        file=sys.stderr,
    )

    print(
        "[vakifbank] ===== EK ÜRÜN KONTROLÜ =====",
        file=sys.stderr,
    )

    for label, term in [
        ("Fatura", "fatura"),
        ("HGS", "hgs"),
        ("Kiralık Kasa", "kiralik kasa"),
    ]:
        found = [
            row
            for row in rows
            if (
                term in _normalize_key(
                    row.masraf
                )
                or term in _normalize_key(
                    row.kategori
                )
            )
        ]

        print(
            f"[vakifbank] {label}: "
            f"{len(found)} kayıt",
            file=sys.stderr,
        )

        for row in found[:8]:
            print(
                f"    - "
                f"[{row.kategori}] "
                f"{row.masraf}",
                file=sys.stderr,
            )

    print(
        "[vakifbank] =============================",
        file=sys.stderr,
    )


# =========================================================
# ANA SCRAPER
# =========================================================

def scrape_vakifbank(
    page_url: str = VAKIFBANK_PAGE_URL,
) -> List[UcretSatiri]:
    print(
        f"[vakifbank] SÜRÜM: "
        f"{SCRAPER_VERSION}",
        file=sys.stderr,
    )

    print(
        "[vakifbank] "
        "Playwright API intercept başlıyor...",
        file=sys.stderr,
    )

    captured, request_meta = (
        _capture_api_json(
            page_url
        )
    )

    if not captured:
        if request_meta:
            details = "; ".join(
                f"{x.get('status')} "
                f"{x.get('method')} "
                f"{x.get('url')}"
                for x in request_meta
            )

            raise ScraperError(
                "VakıfBank hedef API isteği görüldü "
                "ama kullanılabilir 200 JSON yanıtı alınamadı. "
                f"İstekler: {details}"
            )

        raise ScraperError(
            "VakıfBank getProductServicePrices "
            "API yanıtı yakalanamadı."
        )

    best_fee_list: List[
        dict
    ] = []

    best_path = ""
    best_url = ""

    for capture in captured:
        fee_list, path = (
            _extract_fee_list(
                capture.get(
                    "data"
                )
            )
        )

        print(
            f"[vakifbank] API adayı: "
            f"{capture.get('url')} | "
            f"Fee={len(fee_list)} | "
            f"JSON yolu={path or '?'}",
            file=sys.stderr,
        )

        if len(fee_list) > len(
            best_fee_list
        ):
            best_fee_list = fee_list
            best_path = path
            best_url = _normalize(
                capture.get(
                    "url"
                )
            )

    if not best_fee_list:
        raise ScraperError(
            "VakıfBank API JSON'u yakalandı "
            "ancak ücret listesi bulunamadı."
        )

    print(
        f"[vakifbank] Kullanılan API yanıtı: "
        f"{best_url}",
        file=sys.stderr,
    )

    _print_schema_report(
        best_fee_list,
        best_path,
    )

    rows, stats = _parse_fee_list(
        best_fee_list
    )

    if not rows:
        raise ScraperError(
            "VakıfBank Fee listesi "
            "parse edilemedi."
        )

    rows = sorted(
        rows,
        key=lambda row: (
            _normalize_key(
                row.kategori
            ),
            _normalize_key(
                row.masraf
            ),
        ),
    )

    _print_category_report(
        rows
    )

    _print_integrity_report(
        stats,
        len(rows),
    )

    _print_transfer_report(
        rows
    )

    _print_extra_product_report(
        rows
    )

    print(
        f"[vakifbank] Toplam "
        f"{len(rows)} benzersiz "
        "satır bulundu.",
        file=sys.stderr,
    )

    return rows


# =========================================================
# TEST
# =========================================================

if __name__ == "__main__":
    try:
        veriler = scrape_vakifbank()

        print()
        print("=" * 70)
        print("VAKIFBANK SCRAPER")
        print("=" * 70)
        print(
            f"Toplam çekilen ücret: "
            f"{len(veriler)}"
        )
        print()

        for i, row in enumerate(
            veriler[:40],
            start=1,
        ):
            print(
                f"{i}. "
                f"[{row.kategori}] "
                f"{row.masraf}"
            )
            print(
                f"   Asgari Tutar : "
                f"{row.asgari_tutar}"
            )
            print(
                f"   Asgari Oran  : "
                f"{row.asgari_oran}"
            )
            print(
                f"   Azami Tutar  : "
                f"{row.azami_tutar}"
            )
            print(
                f"   Azami Oran   : "
                f"{row.azami_oran}"
            )
            print(
                f"   Açıklama     : "
                f"{row.aciklama}"
            )
            print(
                f"   Tarih        : "
                f"{row.site_guncelleme_tarihi}"
            )
            print("-" * 70)

    except Exception as exc:
        print(
            f"[vakifbank][HATA] "
            f"{exc}",
            file=sys.stderr,
        )
        sys.exit(1)
