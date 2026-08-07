"""
Komisyon değişikliklerini tespit edip mail atan modül.
"""

import os
import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


def load_excel(path: str) -> pd.DataFrame:
    try:
        return pd.read_excel(path, sheet_name="KOMISYONLAR", dtype=str).fillna("")
    except Exception:
        return pd.DataFrame()


def detect_changes(old_df: pd.DataFrame, new_df: pd.DataFrame) -> pd.DataFrame:
    if old_df.empty:
        return new_df

    key_cols = ["BANKA", "KATEGORİ", "MASRAF"]
    value_cols = ["ASGARİ TUTAR", "ASGARİ ORAN", "AZAMİ TUTAR", "AZAMİ ORAN"]

    key_cols   = [c for c in key_cols   if c in old_df.columns and c in new_df.columns]
    value_cols = [c for c in value_cols if c in old_df.columns and c in new_df.columns]

    if not key_cols:
        return pd.DataFrame()

    old_indexed = old_df.set_index(key_cols)
    new_indexed = new_df.set_index(key_cols)

    degisiklikler = []

    for idx in new_indexed.index:
        if idx not in old_indexed.index:
            raw = new_indexed.loc[idx]
            if isinstance(raw, pd.DataFrame):
                raw = raw.iloc[0]
            row = raw.copy()
            row["DEĞİŞİKLİK"] = "YENİ EKLENDI"
            if isinstance(idx, tuple):
                for k, v in zip(key_cols, idx):
                    row[k] = v
            else:
                row[key_cols[0]] = idx
            degisiklikler.append(row)
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
                    if str(old_raw[col]).strip() != str(new_raw[col]).strip():
                        farklar.append(f"{col}: {old_raw[col]} → {new_raw[col]}")
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
        result = pd.concat([d.to_frame().T if isinstance(d, pd.Series) else pd.DataFrame([d])
                           for d in degisiklikler], ignore_index=True)
        all_cols = result.columns.tolist()
        front_cols = [c for c in key_cols + value_cols + ["DEĞİŞİKLİK"] if c in all_cols]
        other_cols = [c for c in all_cols if c not in front_cols]
        result = result[front_cols + other_cols]
        return result
    except Exception as e:
        print(f"[notify] DataFrame oluşturma hatası: {e}")
        return pd.DataFrame()


def build_html_table(df: pd.DataFrame) -> str:
    rows_html = ""
    for _, row in df.iterrows():
        degisiklik = row.get("DEĞİŞİKLİK", "")
        renk = "#fff3cd" if "YENİ" in str(degisiklik) else "#fde8e8"
        cells = "".join(f"<td style='padding:6px 10px;border:1px solid #ddd'>{row.get(c,'')}</td>"
                        for c in df.columns)
        rows_html += f"<tr style='background:{renk}'>{cells}</tr>"

    headers = "".join(f"<th style='padding:8px 10px;background:#1a3c5e;color:white;border:1px solid #ddd'>{c}</th>"
                      for c in df.columns)

    return f"""
    <table style='border-collapse:collapse;font-family:Arial,sans-serif;font-size:13px;width:100%'>
        <thead><tr>{headers}</tr></thead>
        <tbody>{rows_html}</tbody>
    </table>
    """


def send_mail(changes_df: pd.DataFrame, new_excel_path: str, test_mode: bool = False):
    mail_user = os.environ["MAIL_USER"]
    mail_pass = os.environ["MAIL_PASS"]
    mail_to   = os.environ["MAIL_TO"].split(",")

    if test_mode:
        konu = "✅ TEST — Komisyon bildirim sistemi çalışıyor"
        html_body = """
        <html><body style='font-family:Arial,sans-serif;color:#333'>
            <h2 style='color:#1a3c5e'>✅ Test Bildirimi</h2>
            <p>Komisyon değişiklik bildirim sistemi başarıyla çalışıyor.</p>
            <p>Bugün herhangi bir komisyon değişikliği tespit edilmedi.</p>
            <p style='color:#888;font-size:12px'>Bu bir test mailidir. Gerçek değişiklik olduğunda otomatik bildirim gelecektir.</p>
        </body></html>
        """
    else:
        sayi = len(changes_df)
        konu = f"⚠️ Komisyon Değişikliği Bildirimi — {sayi} değişiklik tespit edildi"
        html_table = build_html_table(changes_df)
        html_body = f"""
        <html><body style='font-family:Arial,sans-serif;color:#333'>
            <h2 style='color:#1a3c5e'>Komisyon Değişiklik Bildirimi</h2>
            <p>Bugün yapılan komisyon güncellemesinde <strong>{sayi} değişiklik</strong> tespit edildi.</p>
            {html_table}
            <br>
            <p style='color:#888;font-size:12px'>Güncel Excel dosyası ekte yer almaktadır.</p>
        </body></html>
        """

    msg = MIMEMultipart("mixed")
    msg["From"]    = mail_user
    msg["To"]      = ", ".join(mail_to)
    msg["Subject"] = konu
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    if os.path.exists(new_excel_path):
        with open(new_excel_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition",
                            f'attachment; filename="komisyonlar_guncel.xlsx"')
            msg.attach(part)

    try:
        print(f"[notify] SMTP bağlantısı kuruluyor... Kullanıcı: {mail_user}")
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(mail_user, mail_pass)
            server.sendmail(mail_user, mail_to, msg.as_string())
        print(f"[notify] Mail gönderildi. Alıcılar: {mail_to}")
    except Exception as e:
        print(f"[notify] MAIL HATASI: {e}")
        raise


def check_and_notify(old_excel: str, new_excel: str, test_mode: bool = False):
    print("[notify] Değişiklik kontrolü başlıyor...")

    if test_mode:
        print("[notify] TEST MODU — mail gönderiliyor...")
        send_mail(pd.DataFrame(), new_excel, test_mode=True)
        return

    old_df = load_excel(old_excel)
    new_df = load_excel(new_excel)

    changes = detect_changes(old_df, new_df)

    if changes.empty:
        print("[notify] Değişiklik yok, mail gönderilmedi.")
        return

    print(f"[notify] {len(changes)} değişiklik bulundu, mail gönderiliyor...")
    send_mail(changes, new_excel)
