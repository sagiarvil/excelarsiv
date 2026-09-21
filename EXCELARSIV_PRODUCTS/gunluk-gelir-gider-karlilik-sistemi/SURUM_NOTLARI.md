# SÜRÜM NOTLARI — GÜNLÜK GELİR-GİDER VE GERÇEK KÂRLILIK SİSTEMİ

**Ürün:** Günlük Gelir-Gider ve Gerçek Kârlılık Sistemi
**Sürüm:** 1.0.0 (10.08.2026)
**Üretici betik:** `gelir_gider_karliik.py`
**Denetim:** **0 KALDI** — G01–G24 (İşçilik) + Ö1, Ö1b, Ö2 (Dayanıklılık) + D01–D14 (Değer)
**Satış fiyatı:** 1.490 TL
**SHA-256:** `6d19ace18c80109214c7773a7e39586b42ebcc98875676e532a2007bb40498c4`

## Bu sürümde neler var

- **Hesap planı:** Gelir/gider hesap listesi, tür ve varsayılan sınıf girişi; veri durumu otomatik izlenir.
- **Gelir-gider kayıt defteri:** Günlük gelir/gider, tahakkuk ve nakit tarihleri, KDV işareti, ortak çekişi/finansman/yatırım ayrımı; nakit ve kârlılık etkisi otomatik hesaplanır.
- **Satış-maliyet karması:** Ürün bazında satış, maliyet, brüt kâr ve marj tablosu.
- **Tahakkuk-nakit köprüsü:** Kârın nakde dönüşümünü kanıtlayan köprü, stok ve yuvarlama toleransı kontrollü.
- **Günlük özet ve aylık kârlılık:** 100 günlük net nakit, kümülatif nakit ve aylık kâr/zarar tablosu.
- **Başabaş analizi:** Katkı payı, başabaş cirosu, güvenlik marjı ve hedef kâr cirosu.
- **Karar motoru:** 7 kural — marj, nakit dönüşümü, ortak çekişi, finansman yükü ve olağandışı gelir; RİSKLİ/DİKKAT/NORMAL kararı, gerekçe ve aksiyon önerisi; veri yokken VERİ YOK.
- **Analitik modüller:** T1, T2, T3, T4, O1, O2, O3, O4, I1, I2, I6 kataloğundan 11 modül; derinlik puanı 175 (eşik ≥100, ileri modül 3).
- **Pano:** 374 canlı KPI, 8 grafik (tahmin katmanı dahil), karar ve gerekçe şeridi, senaryo motoru.
- **Rapor:** PDF'e hazır tek sayfa yönetici raporu; varsayımlar, uyarılar ve aksiyonlar.

## Kalite kanıtları (manda v4)

- G kapıları (G01–G24): **0 KALDI**
- Ölçek kapıları (Ö1, Ö1b, Ö2): **0 KALDI** — 600 satır / 1.800 formül, hesap 3.5 sn, 7 uç durum temiz
- Değer kapıları (D01–D14): **0 KALDI** — derinlik 175, D01 8.500 ≥ 4.470, D02 6 ayrım
- Altın çıktı regresyonu: 0 fark

## Bilinen sınırlamalar ve varsayımlar

- Eşikler, oranlar ve toleranslar işletme varsayımıdır; kendi sektörünüze göre AYARLAR'dan değiştirilmelidir (varsayım listesi PARAMETRE_KAYNAKLARI.md'de).
- Vergi oranı ve mevzuat parametreleri sabitlenmemiştir; güncel değerlerle güncellenmelidir.
- Kârlılık hesabı tahakkuk esasına, nakit hesabı nakit tarihine dayanır; ortak çekişi, kredi anapara ve yatırım harcamaları kârlılığı etkilemez (yalnızca nakdi etkiler).
- Hesaplar mali danışmanlık görüşünün yerine geçmez; karar destek amaçlıdır.
- Kişisel veri işlenmez; dosya tamamen çevrimdışı çalışır; verileriniz cihazınızdan çıkmaz.
