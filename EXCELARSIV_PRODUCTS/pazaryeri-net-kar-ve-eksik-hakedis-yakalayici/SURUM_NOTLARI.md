# SÜRÜM NOTLARI — PAZARYERİ NET KÂR & EKSİK HAKEDİŞ YAKALAYICI

## 1.0.0 — İlk sürüm

### Ürün amacı
Pazaryerinde satılan kalemlerin komisyon, kargo, reklam, iade ve ek gider sonrası net kârını ve kâr marjını hesaplar; beklenen hakediş ile pazaryerinden gelen gerçek hakedişi karşılaştırarak eksik tahsilatı eşiklerle işaretler; pazaryeri bazlı satış/kâr özeti, yoğunlaşma, anomali, senaryo ve tornado analiziyle net kâr projeksiyonu ve GÜÇLÜ/NORMAL/İNCELE kararı üretir.

### Özellikler
- **Kalem listesi:** 1.000 satır kapasiteli tabloda kalem kodu, ürün adı, pazaryeri, satış fiyatı, miktar, birim maliyet ve oranlar girişi; toplam satış, komisyon/kargo/reklam/iade tutarları, toplam maliyet, net kâr, marj, beklenen hakediş, hakediş farkı, fark oranı ve durum canlı formülle.
- **Eksik hakediş motoru:** Beklenen hakediş ile gerçek hakedişi karşılaştırır; fark oranı ve tutar eşiklerini aşan kalemleri `EKSİK HAKEDİŞ`, tahsilatı girmeyenleri `BEKLİYOR` olarak işaretler.
- **Karar kapısı:** GÜÇLÜ / NORMAL / İNCELE / VERİ YOK / DURDUR kararı ve gerekçesi; aksiyon önerileri.
- **Analitik motor:** Ortalama marj, ortalama komisyon, kâr sapması (MAD), en yüksek kâr, senaryo bant genişliği, HHI yoğunlaşması, kalite skoru, tahmin + güven bandı, P90 yüzdelik (10 modül, 170 puan; 3 ileri modül).
- **Panel:** 10 grafik (pazaryeri bazlı satış/net kâr/marj/beklenen hakediş, kalem bazlı net kâr trendi, hakediş farkı, durum dağılımı, projeksiyon bant aralığı, senaryo net kârları) ve 66 canlı KPI.
- **Senaryo ve duyarlılık:** İyi/Baz/Kötü/Kritik senaryo bant aralığı ve komisyon oranı/satış fiyatı tornado analizi.

### Denetim sonuçları
- G kapıları (G01–G24): **0 KALDI**
- Ö kapıları (Ö1, Ö1b, Ö2): **0 KALDI** (3.000 satır / 23.984 formül, 0 hata, 5,0 sn; 7 uç durum temiz)
- D kapıları (D01–D14): **0 KALDI** (D01 ikame 12.000 ₺ ≥ 7.470 ₺; D02 5 ayrım; D03 170 puan; D04 3 ileri modül)

### Altın çıktı
- Altın dosya: `KANIT/altin_cikti_pazaryeri.json`
- İçerik SHA-256: `027a8e93d3752b6fa8bbb4235fefdeb74d676a5ff0a36d97c2b2f6fc90a8783f`
- Fark sayısı: 0

### Dosya parmak izi
- SHA-256 (teslim edilen dosya): `6726685e449d1c3688c30581acddf1db294934330c04d23e0cfc4c3c73575e2b`
- KANIT'taki üç rapor da aynı parmak iziyle imzalanmıştır (RAPOR_DENETIM, RAPOR_OLCEK, RAPOR_DEGER).

### Değişiklik günlüğü
| Tarih | Sürüm | Değişiklik | Neden |
|---|---|---|---|
| 10.08.2026 | 1.0.0 | İlk üretim | Ürünün yayına hazırlanması |
