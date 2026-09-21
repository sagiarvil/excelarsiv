# FİYAT SAVUNMASI — Tahsilat Riski & Vade Komuta Paneli

**Satış fiyatı:** 7.900 TL
**Denetim eşiği:** 3 × satış fiyatı = 23.700 TL (ikame maliyeti bu değerin üzerinde olmalı)

## D01 — İkame Maliyeti Çapası

| Bileşen | Tutar (₺) | Gerekçe |
|---|---|---|
| Y1 — Danışman ikamesi | 16.000 | Tahsilat risk danışmanlığının CFO + tahsilat müdürü + SMMM üçlüsünde kurulması ve dönem kanıt dosyasının hazırlanması |
| Y2 — Kurulum ikamesi | 8.000 | Durum makinesi, yaşlandırma kovaları, tahsilat olasılığı, beklenen kayıp, 8 grafikli pano ve kanıt raporunun sıfırdan Excel’de kurulması |
| **Toplam ikame** | **24.000** | 24.000 ≥ 23.700 ✓ |

**Y1+Y2 ≥ 3 × satış fiyatı:** 24.000 ≥ 23.700 → **GEÇTİ**

## D02 — Serbest Alternatif

| Bedava yöntem | Ayrım | Dosya karşılığı |
|---|---|---|
| Elle Excel tahsilat listesi | Durum haritası AYARLAR'da | `trv_durumHaritasi` |
| ERP cari yaşlandırma | Geçersiz geçiş görünür uyarı | `trv_gecersizGecis` |
| Banka vade defteri | Yaşlandırma kovaları 0-30/31-60/61-90/90+ | `trv_yasKova` |
| Elle kuyruk listesi | Kapalı kayıt kuyruktan düşer | `trv_kapaliDusum` |
| Elle numara defteri | Kimlik disiplini otomatik öneri | `trv_sonrakiKimlik` |

## En zayıf nokta

Örnek veri 28 akış satırıdır; gerçek tahsilat dönemi yüzlerce satıra çıkabilir. Ölçek sözleşmesi SPEC’te 20.000’dir; dosya içi tablo kapasitesi 250’dir. Geçersiz geçiş makrosuz olduğu için engellenmez; kullanıcı uyarıyı görmezden gelebilir — bu sınır KILAVUZ’da yazılıdır.
