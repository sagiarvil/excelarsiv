# SÜRÜM NOTLARI — FAZLA MESAİ VE İŞÇİ DAVA RİSKİ TESPİT DOSYASI

**Ürün:** Fazla Mesai ve İşçi Dava Riski Tespit Dosyası
**Sürüm:** 1.0.0 (09.08.2026)
**Satış fiyatı:** 2.490 TL
**SHA-256:** `9132a9860eb89098fcdf1166bd957aee7a07797e66b448081fb3aec55d24d431`

## Bu sürümde neler var

- **Çalışan kartı:** Ad, işe giriş tarihi, aylık brüt ücret, haftalık çalışma saati ve aylık fazla
  mesai saati girişi; tablo kapasitesi 1.000 satır.
- **Mesai ücreti motoru:** 4857 s. İş Kanunu md.41 kuralıyla (saat ücreti + %50) fazla çalışma
  ücreti; md.42 (%25) fazla sürelerle çalışma katsayısı; yıllık 270 saat sınırı ve aşım tutarı.
- **Kayıt mutabakatı:** Bildirilen ile belgelenen mesai saati karşılaştırması, fark saati ve
  fark oranı; kayıt eksikliği riski tutarı.
- **Dava riski modeli:** Zamanaşımı penceresi (4857 s. md.32/8, 5 yıl) risk tutarı, dava risk
  oranı, NPV ile bugünkü değer.
- **Karar motoru:** UYGUN / İNCELE / DURDUR kararı, makine üretimi gerekçe ve 4 aksiyon;
  veri yokken VERİ YOK kararı.
- **Analitik modüller:** T1-T6, O1-O8, I1-I7 katalogundan 15 modül; derinlik puanı 240
  (eşik ≥100, ileri modül ≥2).
- **Pano:** 23 canlı KPI, 6 grafik, karar ve gerekçe şeridi, en kritik 3 içgörü.
- **Senaryo/duyarlılık:** Kötümser/baz/iyimser senaryo motoru, tornado, yoğunlaşma analizi.
- **Rapor:** PDF'e hazır tek sayfa yönetici raporu; varsayımlar, uyarılar ve aksiyonlar.

## Kalite kanıtları (manda v4)

- G kapıları (G01-G24): **0 KALDI**
- Ölçek kapıları (Ö1, Ö1b, Ö2): **0 KALDI** — 1.200 satır, hesap süresi eşik altında
- Değer kapıları (D01-D14): **0 KALDI** — derinlik 240, D01 11.500 ≥ 7.470, D02 8 ayrım
- Altın çıktı regresyonu: 0 fark

## Bilinen sınırlamalar ve varsayımlar

- Zamanaşımı süresi (5 yıl) ve bütçe eşiği kullanıcı tarafından kendi koşullarına göre
  doğrulanmalıdır; dosya AYARLAR'da ayrı parametreler olarak tutar.
- Kayıtların "belgelenen" kısmı işverenin elindeki resmi kayda göre girilir; eksik belge dava
  riskini artırır, dosya bunu oran ve tutar olarak raporlar.
- Hesaplar iş hukuku görüşünün yerine geçmez; karar destek amaçlıdır.
- Kişisel veri işlenmez; dosya tamamen çevrimdışı çalışır; verileriniz cihazınızdan çıkmaz.
