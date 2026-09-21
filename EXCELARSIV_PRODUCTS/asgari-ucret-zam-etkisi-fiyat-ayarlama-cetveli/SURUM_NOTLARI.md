# SÜRÜM NOTLARI — ASGARİ ÜCRET ZAM ETKİSİ & FİYAT AYARLAMA CETVELİ

**Ürün:** Asgari Ücret Zam Etkisi & Fiyat Ayarlama Cetveli
**Sürüm:** 1.0.0 (09.08.2026)
**Satış fiyatı:** 1.490 TL
**SHA-256:** `f7935923f4ac20ff75354188f27b92dd55d49f1c038649de9aaef774a17264d9`

## Bu sürümde neler var

- **Çalışan kartı:** Ad, işe giriş tarihi, zam öncesi brüt ücret, ücret türü (Asgari / Ust Oran /
  Sabit) ve zam oranı girişi; tablo kapasitesi 1.000 satır.
- **Zam motoru:** Çalışan bazında saatlik ücret, zam sonrası brüt ücret, işveren maliyeti
  (SGK + işsizlik işveren payları dâhil), maliyet farkı ve artış oranı; kıdem günü hesabı.
- **Fiyat ayarlama cetveli:** Ürün/hizmet bazında işçilik oranı, fiyat yansıtma oranı,
  önerilen yeni fiyat, yeni satış hasılatı ve işçilik yükü; tablo kapasitesi 2.000 satır.
- **Karar motoru:** UYGUN / İNCELE / DURDUR kararı, makine üretimi gerekçe ve 3 aksiyon;
  veri yokken VERİ YOK kararı.
- **Analitik modüller:** T1-T6, O1-O8, I1-I7 katalogundan 15 modül; derinlik puanı 250
  (eşik ≥100, ileri modül ≥2).
- **Pano:** 22 canlı KPI, 6 grafik, karar ve gerekçe şeridi, en kritik 3 içgörü.
- **Senaryo/duyarlılık:** Kötümser/baz/iyimser senaryo motoru, tornado, yoğunlaşma analizi.
- **Rapor:** PDF'e hazır tek sayfa yönetici raporu; varsayımlar, uyarılar ve aksiyonlar.

## Kalite kanıtları (manda v4)

- G kapıları (G01-G24): **0 KALDI**
- Ölçek kapıları (Ö1, Ö1b, Ö2): **0 KALDI** — 1.200 satır, hesap süresi eşik altında
- Değer kapıları (D01-D14): **0 KALDI** — derinlik 250, D01 7.000 ≥ 4.470, D02 8 ayrım
- Altın çıktı regresyonu: 0 fark

## Bilinen sınırlamalar ve varsayımlar

- Asgari ücret değerleri (eski/yeni) ve SGK/işsizlik prim oranları Resmi Gazete duyurusu sonrası
  güncellenmelidir; dosya AYARLAR'da ayrı parametreler olarak tutar.
- Fiyat yansıtma oranı işçilik oranıyla doğrusal varsayılır; pazar koşulları farklılaştırabilir.
- Hesaplar mali danışmanlık görüşünün yerine geçmez; karar destek amaçlıdır.
- Kişisel veri işlenmez; dosya tamamen çevrimdışı çalışır; verileriniz cihazınızdan çıkmaz.
