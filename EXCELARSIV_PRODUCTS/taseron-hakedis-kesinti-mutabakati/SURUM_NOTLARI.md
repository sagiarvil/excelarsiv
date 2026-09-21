# SÜRÜM NOTLARI — TAŞERON/ALT YÜKLENİCİ HAKEDİŞ & KESİNTİ MUTABAKATI

**Ürün:** Taşeron/Alt Yüklenici Hakediş & Kesinti Mutabakatı
**Sürüm:** 1.0.0 (09.08.2026)
**Satış fiyatı:** 1.490 TL
**SHA-256:** `0a6bce31e2f81e0cbcae9a265fbb47b45f784b49f741d907c80a8d3db893fefd`

## Bu sürümde neler var

- **Sözleşme kartı:** Yüklenici/taşeron/iş bilgileri, sözleşme bedeli, tarihler, iş türü ve
  KDV tevkifatı seçimi.
- **Hakediş motoru:** Her hakediş için KDV, KDV tevkifatı (4/10), gelir vergisi stopajı, SGK
  kesintisi, teminat kesintisi ve net hakediş otomatik hesaplanır.
- **Mutabakat:** Hak edilen net tutar ile ödenen tutarın farkı ve fark oranı üretilir;
  tahsilat oranı izlenir.
- **Geç ödeme ve gecikme maliyeti:** Ödeme tarihi vadeyle karşılaştırılır; gecikme günü ve
  maliyeti yıllık faiz oranından hesaplanır.
- **Karar motoru:** UYGUN / İNCELE / DURDUR kararı, makine üretimi gerekçe ve 4 aksiyon.
- **Analitik modüller:** T1-T6, O1-O8, I1-I7 katalogundan 15 modül; derinlik puanı 240
  (eşik ≥100, ileri modül ≥2).
- **Pano:** 18 canlı KPI, 6 grafik, karar ve gerekçe şeridi.
- **Senaryo/duyarlılık:** Kötümser/baz/iyimser senaryo motoru, tornado, kırılım analizi.
- **Rapor:** PDF'e hazır tek sayfa yönetici raporu; varsayımlar, uyarılar ve aksiyonlar.

## Kalite kanıtları (manda v4)

- G kapıları (G01-G24): **0 KALDI**
- Ölçek kapıları (Ö1, Ö1b, Ö2): **0 KALDI** — 2.400 satır / 9.592 formül, 3.1 sn hesap
- Değer kapıları (D01-D14): **0 KALDI** — derinlik 240, D01 7.000 ≥ 4.470, D02 7 ayrım
- Altın çıktı regresyonu: 0 fark

## Bilinen sınırlamalar ve varsayımlar

- Kesinti oranları sözleşme ve mevzuat tebliğlerinden teyit edilmelidir; AYARLAR'dan
  güncellenebilir.
- SGK kesintisi, alt işveren prim sorumluluğu kapsamında sözleşme hükmüne göre yapılan
  ayrıştırmadır; işveren bordrosu yerine geçmez.
- Gecikme faizi oranı yıllık sabittir; sözleşmedeki faiz hükmüyle farklılaşabilir.
- Kişisel veri işlenmez; dosya tamamen çevrimdışı çalışır.
