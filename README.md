# komisyon_analiz

Bankaların komisyon ücretlerini günlük olarak takip eden ve Excel'e yazan otomasyon aracı.

## Ne yapar?

Bu proje şu anda **Garanti BBVA** için çalışmaktadır:

- [Garanti BBVA Ürün ve Hizmet Ücretleri](https://www.garantibbva.com.tr/urun-ve-hizmet-ucretleri) sayfasındaki tüm kategori ve masraf kalemlerini otomatik olarak çeker.
- Çekilen verileri `garanti_komisyonlar.xlsx` dosyasına yazar.
- Değer değişmemişse dokunmaz; değer değişmişse ilgili hücreyi günceller ve **GÜNCELLEME TARİHİ** kolonuna o günün tarihini yazar.
- Yeni bir masraf kalemi sitede çıkarsa, otomatik olarak yeni bir satır ekler.
- Sitede artık bulunmayan kalemler Excel'den silinmez, olduğu gibi kalır.

## Excel yapısı

| KATEGORİ | MASRAF | ASGARİ TUTAR | ASGARİ ORAN | AZAMİ TUTAR | AZAMİ ORAN | AÇIKLAMA | GÜNCELLEME TARİHİ |
|---|---|---|---|---|---|---|---|

## Manuel çalıştırma

```bash
pip install -r requirements.txt
playwright install --with-deps chromium
python main.py
```

Çalıştırdıktan sonra `garanti_komisyonlar.xlsx` dosyası oluşur veya güncellenir.

## Otomatik çalışma (GitHub Actions)

`.github/workflows/daily-scrape.yml` dosyası her gün **Türkiye saati ile 08:00**'de (UTC 05:00) otomatik olarak çalışır:

1. Scraper'ı çalıştırır.
2. Excel dosyasını günceller.
3. Değişiklik varsa otomatik olarak commit edip repoya push eder.

Manuel olarak da tetiklemek isterseniz: **Actions** sekmesi → **Günlük Komisyon Güncelleme** → **Run workflow**.

## İleride eklenecekler

- Diğer bankaların (Ziraat, İş Bankası, Yapı Kredi, Akbank vb.) eklenmesi.
- Kod, banka bazlı scraper fonksiyonlarına ayrılacak şekilde genişletilebilir yapıda tasarlanmıştır.
