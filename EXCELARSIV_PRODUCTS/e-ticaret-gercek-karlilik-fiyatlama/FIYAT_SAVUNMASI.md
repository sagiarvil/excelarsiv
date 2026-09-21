# FİYAT SAVUNMASI — E-Ticaret Gerçek Kârlılık & Fiyatlama

**Satış fiyatı:** 4.900 TL
**Denetim eşiği:** 3 × satış fiyatı = 14.700 TL

## D01 — İkame Maliyeti Çapası

| Bileşen | Tutar (₺) | Gerekçe |
|---|---|---|
| Y1 — Danışman ikamesi | 12.000 | Ürün bazlı gerçek kâr + fiyatlama danışmanlığı |
| Y2 — Kurulum ikamesi | 5.000 | SKU eşleştirme, MOTOR, EGK kök neden, 8 grafik pano |
| **Toplam ikame** | **17.000** | 17.000 ≥ 14.700 ✓ |

## D02 — Serbest Alternatif

| Bedava yöntem | Ayrım | Dosya karşılığı |
|---|---|---|
| Elle komisyon Exceli | Normalize + duplike SKU | `egk_normAnahtar` |
| Pazaryeri panosu | Dört kova bütünlük | `egk_kovaButunluk` |
| Muhasebe kâr özeti | TANIMSIZ kök neden | `egk_tanimsizKova` |
| Sezgisel fiyat | Eşleşme oranı veriden | `egk_eslesmeOrani` |
| Sınırsız satır vaadi | 50.000 ölçek | `egk_olcekHedef` |

## En zayıf nokta

Örnek 35 satırdır; gerçek katalog binlerce SKU olabilir. Tablo kapasitesi 500; ölçek sözleşmesi SPEC'te 50.000'dir.
