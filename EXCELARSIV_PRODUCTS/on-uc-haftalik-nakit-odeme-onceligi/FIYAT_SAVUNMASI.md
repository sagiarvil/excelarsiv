# FİYAT SAVUNMASI — 13 Haftalık Nakit + Ödeme Önceliği

**Satış fiyatı:** 9.900 TL
**Denetim eşiği:** 3 × satış fiyatı = 29.700 TL (ikame maliyeti bu değerin üzerinde olmalı)

## D01 — İkame Maliyeti Çapası

| Bileşen | Tutar (₺) | Gerekçe |
|---|---|---|
| Y1 — Danışman ikamesi | 20.000 | 13 haftalık nakit planlama danışmanlığı (CFO + mali işler) |
| Y2 — Kurulum ikamesi | 10.000 | Durum makinesi, kuyruk, yaşlandırma, öncelik skoru, 8 grafikli pano ve kanıt raporunun sıfırdan kurulması |
| **Toplam ikame** | **30.000** | 30.000 ≥ 29.700 ✓ |

**Y1+Y2 ≥ 3 × satış fiyatı:** 30.000 ≥ 29.700 → **GEÇTİ**

## D02 — Serbest Alternatif

| Bedava yöntem | Ayrım | Dosya karşılığı |
|---|---|---|
| Elle Excel | Durum haritası | `onh_durumHaritasi` |
| ERP nakit | Yaşlandırma | `onh_yasKova` |
| Banka Excel | Geçersiz geçiş | `onh_gecersizGecis` |
| Elle plan | Tahmin aralık | `onh_tahminAralik` |
| Elle senaryo | Senaryo farkı | `onh_senaryoKarsilastirma` |

## En zayıf nokta

Örnek veri 28 akış satırıdır; gerçek nakit kuyruğu yüzlerce satıra çıkabilir. Ölçek sözleşmesi SPEC’te 20.000’dir; dosya içi tablo kapasitesi 250’dir. Geçersiz geçiş makrosuz olduğu için engellenmez; kullanıcı uyarıyı görmezden gelebilir — bu sınır KILAVUZ’da yazılıdır.
