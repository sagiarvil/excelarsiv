# FİYAT SAVUNMASI — İnşaat Hakediş Yönetim Sistemi

**Satış fiyatı:** 9.900 TL
**Denetim eşiği:** 3 × satış fiyatı = 29.700 TL (ikame maliyeti bu değerin üzerinde olmalı)

## D01 — İkame Maliyeti Çapası

| Bileşen | Tutar (₺) | Gerekçe |
|---|---|---|
| Y1 — Danışman ikamesi | 22.000 | Hakediş takip danışmanlığının şantiye muhasebesi + proje müdürü + SMMM üçlüsünde kurulması ve dönem kanıt dosyasının hazırlanması |
| Y2 — Kurulum ikamesi | 9.000 | Durum makinesi, kuyruk, yaşlandırma, zincir denetimi, 8 grafikli pano ve kanıt raporunun sıfırdan Excel’de kurulması |
| **Toplam ikame** | **31.000** | 31.000 ≥ 29.700 ✓ |

**Y1+Y2 ≥ 3 × satış fiyatı:** 31.000 ≥ 29.700 → **GEÇTİ**

## D02 — Serbest Alternatif

| Bedava yöntem | Ayrım | Dosya karşılığı |
|---|---|---|
| Elle Excel takip | Durum haritası AYARLAR'da | `ihy_durumHaritasi` |
| ERP hakediş modülü | Geçersiz geçiş görünür uyarı | `ihy_gecersizGecis` |
| Taşeron defteri | Yaşlandırma kovaları | `ihy_yasKova` |
| Elle kuyruk listesi | Kapalı kayıt kuyruktan düşer | `ihy_kapaliDusum` |
| Elle numara defteri | Kimlik disiplini otomatik öneri | `ihy_sonrakiKimlik` |

## En zayıf nokta

Örnek veri 28 akış satırıdır; gerçek şantiye dönemi yüzlerce satıra çıkabilir. Ölçek sözleşmesi SPEC’te 20.000’dir; dosya içi tablo kapasitesi 250’dir. Geçersiz geçiş makrosuz olduğu için engellenmez; kullanıcı uyarıyı görmezden gelebilir — bu sınır KILAVUZ’da yazılıdır.
