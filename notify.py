"""
Komisyon değişikliklerini tespit edip mail atan modül.
Microsoft 365 / Outlook SMTP uyumlu.
"""

import os
import ssl
import smtplib
import pandas as pd

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


# =========================================================
# EXCEL OKUMA
# =========================================================

def load_excel(path: str) -> pd.DataFrame:
    try:
        return (
            pd.read_excel(
                path,
                sheet_name="KOMISYONLAR",
                dtype=str
            )
            .fillna("")
        )
    except Exception as exc:
        print(f"[notify] Excel okunamadı: {path} -> {exc}")
        return pd.DataFrame()


# =========================================================
# DEĞİŞİKLİK TESPİTİ
# =========================================================

def detect_changes(
    old_df: pd.DataFrame,
    new_df: pd.DataFrame
) -> pd.DataFrame:

    if old_df.empty:
        return new_df

    key_cols = [
        "BANKA",
        "KATEGORİ",
        "MASRAF"
    ]

    value_cols = [
        "ASGARİ TUTAR",
        "ASGARİ ORAN",
        "AZAMİ TUTAR",
        "AZAMİ ORAN"
    ]

    key_cols = [
        c for c in key_cols
        if c in old_df.columns and c in new_df.columns
    ]

    value_cols = [
        c for c in value_cols
        if c in old_df.columns and c in new_df.columns
    ]

    if not key_cols:
        print("[notify] Karşılaştırma için ortak anahtar kolon bulunamadı.")
        return pd.DataFrame()

    old_indexed = (
        old_df
        .set_index(key_cols)
        .sort_index()
    )

    new_indexed = (
        new_df
        .set_index(key_cols)
        .sort_index()
    )

    degisiklikler = []

    for idx in new_indexed.index:

        # -------------------------------------------------
        # YENİ SATIR
        # -------------------------------------------------

        if idx not in old_indexed.index:

            raw = new_indexed.loc[idx]

            if isinstance(raw, pd.DataFrame):
                raw = raw.iloc[0]

            row = raw.copy()

            row["DEĞİŞİKLİK"] = "YENİ EKLENDİ"

            if isinstance(idx, tuple):
                for k, v in zip(key_cols, idx):
                    row[k] = v
            else:
                row[key_cols[0]] = idx

            degisiklikler.append(row)

        # -------------------------------------------------
        # MEVCUT SATIRDA DEĞİŞİKLİK
        # -------------------------------------------------

        else:

            old_raw = old_indexed.loc[idx]
            new_raw = new_indexed.loc[idx]

            if isinstance(old_raw, pd.DataFrame):
                old_raw = old_raw.iloc[0]

            if isinstance(new_raw, pd.DataFrame):
                new_raw = new_raw.iloc[0]

            farklar = []

            for col in value_cols:

                if col in old_raw and col in new_raw:

                    eski = str(old_raw[col]).strip()
                    yeni = str(new_raw[col]).strip()

                    if eski != yeni:
                        farklar.append(
                            f"{col}: {eski} → {yeni}"
                        )

            if farklar:

                row = new_raw.copy()

                row["DEĞİŞİKLİK"] = " | ".join(farklar)

                if isinstance(idx, tuple):
                    for k, v in zip(key_cols, idx):
                        row[k] = v
                else:
                    row[key_cols[0]] = idx

                degisiklikler.append(row)

    if not degisiklikler:
        return pd.DataFrame()

    try:

        frames = []

        for item in degisiklikler:

            if isinstance(item, pd.Series):
                frames.append(item.to_frame().T)
            else:
                frames.append(pd.DataFrame([item]))

        result = pd.concat(
            frames,
            ignore_index=True
        )

        all_cols = result.columns.tolist()

        front_cols = [
            c
            for c in key_cols
            + value_cols
            + ["DEĞİŞİKLİK"]
            if c in all_cols
        ]

        other_cols = [
            c
            for c in all_cols
            if c not in front_cols
        ]

        return result[
            front_cols + other_cols
        ]

    except Exception as exc:

        print(
            f"[notify] DataFrame oluşturma hatası: {exc}"
        )

        return pd.DataFrame()


# =========================================================
# HTML TABLO
# =========================================================

def build_html_table(df: pd.DataFrame) -> str:

    rows_html = ""

    for _, row in df.iterrows():

        degisiklik = row.get(
            "DEĞİŞİKLİK",
            ""
        )

        renk = (
            "#fff3cd"
            if "YENİ" in str(degisiklik)
            else "#fde8e8"
        )

        cells = "".join(
            f"""
            <td style="
                padding:6px 10px;
                border:1px solid #ddd
            ">
                {row.get(c, '')}
            </td>
            """
            for c in df.columns
        )

        rows_html += (
            f"<tr style='background:{renk}'>"
            f"{cells}"
            f"</tr>"
        )

    headers = "".join(
        f"""
        <th style="
            padding:8px 10px;
            background:#1a3c5e;
            color:white;
            border:1px solid #ddd
        ">
            {c}
        </th>
        """
        for c in df.columns
    )

    return f"""
    <table style="
        border-collapse:collapse;
        font-family:Arial,sans-serif;
        font-size:13px;
        width:100%
    ">
        <thead>
            <tr>
                {headers}
            </tr>
        </thead>

        <tbody>
            {rows_html}
        </tbody>
    </table>
    """


# =========================================================
# MAIL GÖNDERME
# =========================================================

def send_mail(
    changes_df: pd.DataFrame,
    new_excel_path: str,
    test_mode: bool = False
):

    # -----------------------------------------------------
    # GitHub Secrets
    # -----------------------------------------------------

    mail_user = os.environ["MAIL_USER"].strip()
    mail_pass = os.environ["MAIL_PASS"].strip()

    mail_to = [
        adres.strip()
        for adres in os.environ["MAIL_TO"].split(",")
        if adres.strip()
    ]

    # Outlook / Microsoft 365
    smtp_host = os.getenv(
        "SMTP_HOST",
        "smtp.office365.com"
    )

    smtp_port = int(
        os.getenv(
            "SMTP_PORT",
            "587"
        )
    )

    if not mail_to:
        raise ValueError(
            "MAIL_TO boş. En az bir alıcı tanımlanmalı."
        )

    # =====================================================
    # TEST MAILİ
    # =====================================================

    if test_mode:

        konu = (
            "✅ TEST — Komisyon bildirim sistemi çalışıyor"
        )

        html_body = """
        <html>
        <body style="
            font-family:Arial,sans-serif;
            color:#333
        ">

            <h2 style="color:#1a3c5e">
                ✅ Test Bildirimi
            </h2>

            <p>
                Komisyon değişiklik bildirim sistemi
                başarıyla çalışıyor.
            </p>

            <p>
                Bu mail Microsoft 365 / Outlook
                üzerinden gönderilmiştir.
            </p>

            <p style="
                color:#888;
                font-size:12px
            ">
                Bu bir test mailidir.
                Gerçek değişiklik olduğunda otomatik
                bildirim gelecektir.
            </p>

        </body>
        </html>
        """

    # =====================================================
    # GERÇEK DEĞİŞİKLİK MAILİ
    # =====================================================

    else:

        sayi = len(changes_df)

        konu = (
            f"⚠️ Komisyon Değişikliği Bildirimi "
            f"— {sayi} değişiklik tespit edildi"
        )

        html_table = build_html_table(
            changes_df
        )

        html_body = f"""
        <html>
        <body style="
            font-family:Arial,sans-serif;
            color:#333
        ">

            <h2 style="color:#1a3c5e">
                Komisyon Değişiklik Bildirimi
            </h2>

            <p>
                Bugün yapılan komisyon güncellemesinde
                <strong>{sayi} değişiklik</strong>
                tespit edildi.
            </p>

            {html_table}

            <br>

            <p style="
                color:#888;
                font-size:12px
            ">
                Güncel Excel dosyası ekte
                yer almaktadır.
            </p>

        </body>
        </html>
        """

    # =====================================================
    # MAIL MESAJI
    # =====================================================

    msg = MIMEMultipart("mixed")

    msg["From"] = mail_user
    msg["To"] = ", ".join(mail_to)
    msg["Subject"] = konu

    msg.attach(
        MIMEText(
            html_body,
            "html",
            "utf-8"
        )
    )

    # =====================================================
    # EXCEL EKİ
    # =====================================================

    if os.path.exists(new_excel_path):

        with open(new_excel_path, "rb") as f:

            part = MIMEBase(
                "application",
                "octet-stream"
            )

            part.set_payload(
                f.read()
            )

        encoders.encode_base64(
            part
        )

        part.add_header(
            "Content-Disposition",
            'attachment; filename="komisyonlar_guncel.xlsx"'
        )

        msg.attach(
            part
        )

    else:

        print(
            "[notify] UYARI: "
            f"Ek dosya bulunamadı: {new_excel_path}"
        )

    # =====================================================
    # OUTLOOK / MICROSOFT 365 SMTP
    # =====================================================

    try:

        print(
            "[notify] Microsoft 365 SMTP bağlantısı kuruluyor..."
        )

        print(
            f"[notify] SMTP: {smtp_host}:{smtp_port}"
        )

        print(
            f"[notify] Kullanıcı: {mail_user}"
        )

        context = ssl.create_default_context()

        with smtplib.SMTP(
            smtp_host,
            smtp_port,
            timeout=60
        ) as server:

            server.ehlo()

            server.starttls(
                context=context
            )

            server.ehlo()

            print(
                "[notify] SMTP TLS bağlantısı kuruldu."
            )

            print(
                "[notify] Microsoft hesabına giriş yapılıyor..."
            )

            server.login(
                mail_user,
                mail_pass
            )

            print(
                "[notify] SMTP girişi başarılı."
            )

            server.sendmail(
                mail_user,
                mail_to,
                msg.as_string()
            )

        print(
            f"[notify] Mail başarıyla gönderildi. "
            f"Alıcılar: {mail_to}"
        )

    except smtplib.SMTPAuthenticationError as exc:

        print(
            "[notify] MICROSOFT 365 GİRİŞ HATASI."
        )

        print(
            "[notify] Kullanıcı adı/parola yanlış olabilir "
            "veya kurum SMTP AUTH kullanımını engelliyor olabilir."
        )

        print(
            f"[notify] SMTP cevabı: {exc}"
        )

        raise

    except Exception as exc:

        print(
            f"[notify] MAIL HATASI: {exc}"
        )

        raise


# =========================================================
# ANA KONTROL
# =========================================================

def check_and_notify(
    old_excel: str,
    new_excel: str,
    test_mode: bool = False
):

    print(
        "[notify] Değişiklik kontrolü başlıyor..."
    )

    # =====================================================
    # TEST MODU
    # =====================================================

    if test_mode:

        print(
            "[notify] TEST MODU — mail gönderiliyor..."
        )

        send_mail(
            pd.DataFrame(),
            new_excel,
            test_mode=True
        )

        return

    # =====================================================
    # NORMAL MOD
    # =====================================================

    old_df = load_excel(
        old_excel
    )

    new_df = load_excel(
        new_excel
    )

    changes = detect_changes(
        old_df,
        new_df
    )

    if changes.empty:

        print(
            "[notify] Değişiklik yok, mail gönderilmedi."
        )

        return

    print(
        f"[notify] {len(changes)} değişiklik bulundu, "
        "mail gönderiliyor..."
    )

    send_mail(
        changes,
        new_excel
    )
