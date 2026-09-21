# SÜRÜM NOTLARI — İTHALAT DEPO TESLİM (RAFA GELEN) NET BİRİM MALİYET

**Ürün:** İthalat Depo Teslim (Rafa Gelen) Net Birim Maliyet
**Sürüm:** 1.0.0 (09.08.2026)
**Satış fiyatı:** 1.490 TL
**SHA-256:** `4d69f63d6656a15ba9f69aee08a9c1395c730f83ab65e06c0d85102106c622d8`

## Bu sürümde neler var

- **İthal ürün girişi:** Adet, FOB birim fiyat (döviz), döviz cinsi, kur, navlun/sigorta,
  gümrük vergisi, ÖTV ve KDV oranı girişi; tablo kapasitesi 2.000 satır.
- **Vergi matrah zinciri:** GV = CIF×oran, ÖTV = (CIF+GV)×oran, KDV = (CIF+GV+ÖTV)×oran
  (4458 s., 3065 s. md.21, 4760 s. dayanaklı); FOB payıyla masraf dağıtımı.
- **Yurt içi masraf tablosu:** Liman, müşavirlik, iç nakliye gibi masraflar TL girilir;
  ürünlere FOB payıyla otomatik dağıtılır; tablo kapasitesi 1.000 satır.
- **Karar motoru:** UYGUN / İNCELE / DURDUR kararı, makine üretimi gerekçe ve 4 aksiyon;
  veri yokken VERİ YOK kararı.
- **Analitik modüller:** T1-T6, O1-O8, I1-I7 katalogundan 16 modül; derinlik puanı 250
  (eşik ≥100, ileri modül ≥2).
- **Pano:** 25 canlı KPI, 6 grafik, karar ve gerekçe şeridi, en kritik 3 içgörü.
- **Senaryo/duyarlılık:** Kötümser/baz/iyimser kur ve oran senaryoları, tornado,
  yoğunlaşma analizi, kur dönem farkı.
- **Rapor:** PDF'e hazır tek sayfa yönetici raporu; varsayımlar, uyarılar ve aksiyonlar.

## Kalite kanıtları (manda v4)

- G kapıları (G01-G24): **0 KALDI**
- Ölçek kapıları (Ö1, Ö1b, Ö2): **0 KALDI** — 2.400 satır, hesap süresi eşik altında
- Değer kapıları (D01-D14): **0 KALDI** — derinlik 250, D01 6.500 ≥ 4.470, D02 8 ayrım
- Altın çıktı regresyonu: 0 fark

## Bilinen sınırlamalar ve varsayımlar

- Vergi oranları (gümrük, ÖTV, KDV) GTİP ve mevzuat değişikliklerinde AYARLAR'dan
  güncellenmelidir; dosya varsayılan oranlarla gelir.
- Döviz kurları kullanıcı tarafından güncel girilmelidir; varsayılan kur TCMB gösterge
  kuru bazlıdır ve doğrulanmalıdır.
- Yurt içi masraf dağıtımı FOB payıyla yapılır; dağıtım anahtarı değiştirilebilir.
- Hesaplar gümrük müşavirliği veya mali danışmanlık görüşünün yerine geçmez; karar
  destek amaçlıdır.
- Kişisel veri işlenmez; dosya tamamen çevrimdışı çalışır; verileriniz cihazınızdan çıkmaz.
