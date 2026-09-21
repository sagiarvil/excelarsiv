# FİYAT SAVUNMASI — İSG Risk Değerlendirme Pro (6331)

**Satış fiyatı:** 7.900 TL
**Denetim eşiği:** 3 × satış fiyatı = 23.700 TL (ikame maliyeti bu değerin üzerinde olmalı)

## D01 — İkame Maliyeti Çapası

| Bileşen | Tutar (₺) | Gerekçe |
|---|---|---|
| Y1 — Danışman ikamesi | 18.000 | İSG risk değerlendirme danışmanlığı (envanter + skorlama + eylem planı + denetim dosyası) |
| Y2 — Kurulum ikamesi | 6.000 | L/FK matrisi, durum makinesi, yaşlandırma, 8 grafikli pano ve kanıt raporunun sıfırdan Excel’de kurulması |
| **Toplam ikame** | **24.000** | 24.000 ≥ 23.700 ✓ |

**Y1+Y2 ≥ 3 × satış fiyatı:** 24.000 ≥ 23.700 → **GEÇTİ**

## D02 — Serbest Alternatif

| Bedava yöntem | Ayrım | Dosya karşılığı |
|---|---|---|
| ÇSGB örnek form | L-matris / Fine-Kinney seçimli | `isg_puanlamaMatrisi` |
| Word risk listesi | Eylem planı durum makinesi | `isg_eylemDurum` |
| Danışman PDF'i | Yenileme takvimi | `isg_yenilemeTakvim` |
| Elle takip | Uyum oranı veriden | `isg_uyumOrani` |
| Statik rapor | Revizyon geçmişi bloğu | `isg_revizyonGecmisi` |

## En zayıf nokta

Örnek veri 24 tehlike / 28 eylemdir; gerçek saha yüzlerce satıra çıkabilir. Ölçek sözleşmesi SPEC’te 5.000’dir; dosya içi tablo kapasitesi 200’dür. Matris seçimi (L vs FK) aynı girdide farklı skor üretir — belirsizlik KILAVUZ’da yazılıdır. Geçersiz geçiş makrosuz olduğu için engellenmez.
