# FİYAT SAVUNMASI — Pazaryeri Hakediş Mutabakat Motoru

**Satış fiyatı:** 9.900 TL
**Denetim eşiği:** 3 × satış fiyatı = 29.700 TL (ikame maliyeti bu değerin üzerinde olmalı)

## D01 — İkame Maliyeti Çapası

| Bileşen | Tutar (₺) | Gerekçe |
|---|---|---|
| Y1 — Danışman ikamesi | 22.000 | Ay sonu pazaryeri hakediş mutabakatının operasyon + mali işler + SMMM üçlüsünde danışman eliyle kurulması ve itiraz dosyasının hazırlanması |
| Y2 — Kurulum ikamesi | 8.000 | Normalize anahtar, dört kova motoru, kök-neden kuralları, itiraz metni, 8 grafikli pano ve kanıt raporunun sıfırdan Excel’de kurulması |
| **Toplam ikame** | **30.000** | 30.000 ≥ 29.700 ✓ |

**Y1+Y2 ≥ 3 × satış fiyatı:** 30.000 ≥ 29.700 → **GEÇTİ**

## D02 — Serbest Alternatif

| Bedava yöntem | Ayrım | Dosya karşılığı |
|---|---|---|
| Elle VLOOKUP | Dört kova bütünlüğü denetimli | `pzm_kovaButunluk` |
| SaaS deneme | Normalize + kısa/uzun no | `pzm_normAnahtar` |
| Muhasebeci aylık mutabakat | TANIMSIZ kök-neden görünür | `pzm_tanimsizKova` |
| Sabit oranlı kontrol listesi | Eşleşme oranı veriden türer | `pzm_eslesmeOrani` |
| Sınırsız satır vaadi | 50.000 satır ölçek sözleşmesi | `pzm_olcekHedef` |

## En zayıf nokta

Örnek veri 35 satırdır; gerçek ay sonu dökümü binlerce satıra çıkabilir. Ölçek sözleşmesi SPEC’te 50.000’dir; dosya içi tablo kapasitesi 500’dür ve büyütme kullanıcıya bırakılmıştır. Kök-neden kuralları komisyon/kargo/iade/reklam yakınlığına dayanır; pazaryeri özel kesinti adları TANIMSIZ’a düşer ve elle incelenmelidir.
