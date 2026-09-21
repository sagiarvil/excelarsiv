# FİYAT SAVUNMASI — Kat Karşılığı / Hasılat Paylaşımı Simülatör

**Satış fiyatı:** 9.900 TL
**Denetim eşiği:** 3 × satış fiyatı = 29.700 TL (ikame maliyeti bu değerin üzerinde olmalı)

## D01 — İkame Maliyeti Çapası

| Bileşen | Tutar (₺) | Gerekçe |
|---|---|---|
| Y1 — Danışman ikamesi | 22.000 | Kat karşılığı / hasılat paylaşımı sözleşme danışmanlığı: pay oranı, başabaş ve kâr modelinin müteahhit + arsa sahibi + SMMM üçlüsünde kurulması |
| Y2 — Kurulum ikamesi | 9.000 | Çok yıllı kural tablosu, MOTOR zinciri, durum makinesi, senaryo/tornado, altın vakalar ve KANIT_RAPORU'nun sıfırdan Excel'de kurulması |
| **Toplam ikame** | **31.000** | 31.000 ≥ 29.700 ✓ |

**Y1+Y2 ≥ 3 × satış fiyatı:** 31.000 ≥ 29.700 → **GEÇTİ**

## D02 — Serbest Alternatif

| Bedava yöntem | Ayrım | Dosya karşılığı |
|---|---|---|
| Elle Excel pay oranı | Net hasılat + tarafların kârı + başabaş tek motor | `kkh_senaryoKarsilastirma` |
| Sözleşme avukatı notu | Makine üretilmiş gerekçeli karar | `kkh_kararGerekce` |
| Genel fizibilite şablonu | Her MOTOR adımında kural_id + madde atıfı | `kkh_maddeAtifMetni` |
| Elle senaryo tablosu | M-SEN kural değişimi + tornado | `kkh_kuralDegisimSenaryo` |
| Elle süreç defteri | Sözleşme→inşaat→teslim→paylaşım zinciri | `kkh_kararMetni` |

## En zayıf nokta

Hasılat paylaşımı yorumu (A/B) mevzuatta tartışmalıdır; dosya her iki yorumu yan yana gösterir ama hukuki kesin görüş vermez. Ölçek sözleşmesi 20.000 satırdır; dosya içi tablo kapasitesi 1.000'dür.
