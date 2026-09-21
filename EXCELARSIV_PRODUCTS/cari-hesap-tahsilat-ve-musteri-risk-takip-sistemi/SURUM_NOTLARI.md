# SÜRÜM NOTLARI — CARİ HESAP, TAHSİLAT VE MÜŞTERİ RİSK TAKİP SİSTEMİ

**Ürün:** Cari Hesap, Tahsilat ve Müşteri Risk Takip Sistemi
**Sürüm:** 2.0.0 (10.08.2026)
**Satış fiyatı:** 1.490 TL
**SHA-256:** `bd8f773843f13582c20417be574be4a2dcdf2f87cbe0c347c734804e27038aab`

## Bu sürümde neler var

- **Müşteri kartı:** Müşteri kodu, limit, teminat, vade, sektör ve durum girişi; tablo kapasitesi 200 satır.
- **Fatura / tahsilat / teminat / iletişim takibi:** Kayıt bazlı giriş; kesinleşme ve vade kontrolü.
- **Risk skoru motoru:** Gecikme, vade, limit, teminat, itiraz ve statü bileşenlerinden 0–100 puan;
  müşteri bazında güven ve risk seviyesi.
- **Yaşlandırma ve tahsilat aksiyonu:** Dilim bazlı açık tutar, gecikme günü ve önceliklendirilmiş
  aksiyon planı.
- **Karar motoru:** DEVAM / İZLE / LİMİT AZALT / VADELİ SATIŞI DURDUR kararı, makine üretimi gerekçe ve
  son tarih; veri yokken VERİ YOK kararı.
- **Analitik modüller:** T1-T6, O1-O8, I1-I7 katalogundan 12 modül; derinlik puanı 240 (eşik ≥100, ileri modül 3).
- **Pano:** 39 canlı KPI, 8 grafik, karar ve gerekçe şeridi, en kritik 3 içgörü.
- **Senaryo/duyarlılık:** Kötümser/baz/iyimser senaryo motoru, tornado, güven senaryosu.
- **Rapor:** PDF'e hazır tek sayfa yönetici raporu; varsayımlar, uyarılar ve aksiyonlar.

## Kalite kanıtları (manda v4)

- G kapıları (G01-G24): **0 KALDI**
- Ölçek kapıları (Ö1, Ö1b, Ö2): **0 KALDI** — 1.800 satır / 13.400 formül, hesap 4.5 sn (sınır 25 sn)
- Değer kapıları (D01-D14): **0 KALDI** — derinlik 240, D01 8.500 ≥ 4.470, D02 6 ayrım
- Altın çıktı regresyonu: 0 fark

## Bilinen sınırlamalar ve varsayımlar

- Risk skoru ağırlıkları ve karar eşikleri işletme varsayımıdır; kendi portföyünüze göre AYARLAR'dan
  değiştirilmelidir (varsayım listesi PARAMETRE_KAYNAKLARI.md'de).
- Finansman maliyeti oranı (%30) gecikme maliyetinde kullanılır; sektörel fon maliyetinize göre
  güncelleyin.
- Hesaplar mali danışmanlık görüşünün yerine geçmez; karar destek amaçlıdır.
- Kişisel veri işlenmez; dosya tamamen çevrimdışı çalışır; verileriniz cihazınızdan çıkmaz.
