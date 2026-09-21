# FİYAT SAVUNMASI — Puantaj-Vardiya-Fazla Mesai Kanıt Sistemi

**Satış fiyatı:** 7.900 TL
**Denetim eşiği:** 3 × satış fiyatı = 23.700 TL (ikame maliyeti bu değerin üzerinde olmalı)

## D01 — İkame Maliyeti Çapası

| Bileşen | Tutar (₺) | Gerekçe |
|---|---|---|
| Y1 — Danışman ikamesi | 16.000 | Puantaj/FM uyum danışmanlığının İK + bordro + denetçi üçlüsünde kurulması ve dönem kanıt dosyasının hazırlanması |
| Y2 — Kurulum ikamesi | 8.000 | Durum makinesi, yasal tavan, zincir bayrağı, 8 grafikli pano ve KANIT_RAPORU'nun sıfırdan Excel'de kurulması |
| **Toplam ikame** | **24.000** | 24.000 ≥ 23.700 ✓ |

**Y1+Y2 ≥ 3 × satış fiyatı:** 24.000 ≥ 23.700 → **GEÇTİ**

## D02 — Serbest Alternatif

| Bedava yöntem | Ayrım | Dosya karşılığı |
|---|---|---|
| Elle Excel puantaj listesi | Durum haritası AYARLAR'da | `pvf_durumHaritasi` |
| Kart basım CSV + manuel toplam | Geçersiz geçiş görünür uyarı | `pvf_gecersizGecis` |
| Bordro yazılımı vardiya ekranı | Fiili ↔ planlı ↔ FM zincir bayrağı | `pvf_zincirKirik` |
| Elle kuyruk listesi | Kapalı kayıt kuyruktan düşer | `pvf_kapaliDusum` |
| Elle numara defteri | Kimlik disiplini otomatik öneri | `pvf_sonrakiKimlik` |

## En zayıf nokta

Örnek veri 28 puantaj satırıdır; gerçek dönem yüzlerce satıra çıkabilir. Ölçek sözleşmesi SPEC'te 20.000'dır; dosya içi tablo kapasitesi 250'dir. Geçersiz geçiş makrosuz olduğu için engellenmez; kullanıcı uyarıyı görmezden gelebilir — bu sınır KILAVUZ'da yazılıdır.
