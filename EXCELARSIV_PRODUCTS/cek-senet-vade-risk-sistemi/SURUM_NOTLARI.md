# SÜRÜM NOTLARI — ÇEK SENET VADE VE KARŞILIK RİSK SİSTEMİ

**Ürün:** Çek Senet Vade ve Karşılık Risk Sistemi
**Sürüm:** 1.0.0 (10.08.2026)
**Satış fiyatı:** 1.490 TL
**SHA-256:** `ef5a8de2e06705c8915fec1b7fcf23b26530b406a1c2853716c5df844730d81b`

## Bu sürümde neler var

- **Cari kartı:** Müşteri/tedarikçi kodu, risk skoru, karşılıksız geçmişi ve kredi limiti girişi;
  risk sınıfı otomatik hesaplanır.
- **Çek/senet kayıt defteri:** Belge türü, yön (alınan/verilen), banka/şube/hesap, keşide ve vade
  tarihi, teminat, ciro, portföy durumu, karşılıksız ve hukuki takip bilgisi; risk ağırlıklı TL
  otomatik türetilir.
- **Ciro zinciri ve banka tahsilatı:** Ciro edilen belgelerde sorumluluk takibi, tahsile veriliş ve
  tahsil kayıtları.
- **Vade takvimi:** 61 günlük günlük karşılık açığı, 13 haftalık ödeme yoğunluğu, 7 günlük
  FORECAST tahmin kolonu.
- **Karşılık analizi:** Risk ağırlıklı alınan, karşılık oranı, müşteri/banka yoğunlaşması, karşılıksız
  çek riski ve güvenli yeni çek limiti.
- **Karar motoru:** Haftalık eşik ve riskli müşteri kurallarıyla KRİTİK / RİSKLİ / NORMAL kararı,
  makine üretimi gerekçe ve aksiyon önerisi; veri yokken VERİ YOK kararı.
- **Analitik modüller:** T1, T3, O1, O5, O8, I1, I2 kataloğundan 7 modül; derinlik puanı 115
  (eşik ≥100, ileri modül 2).
- **Pano:** 40 canlı KPI, 9 grafik (tahmin aralığı dahil), karar ve gerekçe şeridi, en kritik 3 içgörü.
- **Rapor:** PDF'e hazır tek sayfa yönetici raporu; varsayımlar, uyarılar ve aksiyonlar.

## Kalite kanıtları (manda v4)

- G kapıları (G01-G24): **0 KALDI**
- Ölçek kapıları (Ö1, Ö1b, Ö2): **0 KALDI** — 6.000 satır / 23.000 formül, hesap 9.4 sn (sınır 25 sn)
- Değer kapıları (D01-D14): **0 KALDI** — derinlik 115, D01 8.500 ≥ 4.470, D02 5 ayrım
- Altın çıktı regresyonu: 0 fark

## Bilinen sınırlamalar ve varsayımlar

- Eşikler, katsayılar ve finansman oranı işletme varsayımıdır; kendi portföyünüze göre AYARLAR'dan
  değiştirilmelidir (varsayım listesi PARAMETRE_KAYNAKLARI.md'de).
- Karşılıksız çek durumu banka bildiriminden sonra elle işaretlenir; sistem bildirimi otomatik almaz.
- Hesaplar mali danışmanlık görüşünün yerine geçmez; karar destek amaçlıdır.
- Kişisel veri işlenmez; dosya tamamen çevrimdışı çalışır; verileriniz cihazınızdan çıkmaz.
