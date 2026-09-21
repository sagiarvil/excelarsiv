# FİYAT SAVUNMASI — Kargo/Desi Maliyet Optimizasyonu

**Satış fiyatı:** 4.900 TL
**Denetim eşiği:** 3 × satış fiyatı = 14.700 TL

## D01 — İkame Maliyeti Çapası

| Bileşen | Tutar (₺) | Gerekçe |
|---|---|---|
| Y1 — Danışman ikamesi | 10.000 | Desi/tarife optimizasyon danışmanlığı |
| Y2 — Kurulum ikamesi | 5.000 | Sipariş↔tarife eşleştirme, MOTOR, KDO kök neden, 8 grafik |
| **Toplam ikame** | **15.000** | 15.000 ≥ 14.700 ✓ |

## D02 — Serbest Alternatif

| Bedava yöntem | Ayrım | Dosya karşılığı |
|---|---|---|
| Elle desi hesap makinesi | Normalize + duplike sipariş | `kdo_normAnahtar` |
| Taşıyıcı panosu | Dört kova bütünlük | `kdo_kovaButunluk` |
| Muhasebe kargo özeti | TANIMSIZ kök neden | `kdo_tanimsizKova` |
| Sezgisel taşıyıcı seçimi | Eşleşme oranı veriden | `kdo_eslesmeOrani` |
| Sınırsız satır vaadi | 50.000 ölçek | `kdo_olcekHedef` |

## En zayıf nokta

Örnek 35 satırdır; gerçek sipariş hacmi binlerce olabilir. Tablo kapasitesi 500; ölçek sözleşmesi SPEC'te 50.000'dir.
