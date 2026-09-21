# FİYAT SAVUNMASI — Fason Üretim Takip Sistemi

**Satış fiyatı:** 7.900 TL
**Denetim eşiği:** 3 × satış fiyatı = 23.700 TL (ikame maliyeti bu değerin üzerinde olmalı)

## D01 — İkame Maliyeti Çapası

| Bileşen | Tutar (₺) | Gerekçe |
|---|---|---|
| Y1 — Danışman ikamesi | 16.000 | Fason takip danışmanlığının üretim planlama + fabrika müdürü + SMMM üçlüsünde kurulması ve dönem kanıt dosyasının hazırlanması |
| Y2 — Kurulum ikamesi | 8.000 | Durum makinesi, verilen↔dönen+fire zinciri, fire/gecikme eşikleri, 8 grafikli pano ve kanıt raporunun sıfırdan Excel’de kurulması |
| **Toplam ikame** | **24.000** | 24.000 ≥ 23.700 ✓ |

**Y1+Y2 ≥ 3 × satış fiyatı:** 24.000 ≥ 23.700 → **GEÇTİ**

## D02 — Serbest Alternatif

| Bedava yöntem | Ayrım | Dosya karşılığı |
|---|---|---|
| Elle Excel fason defteri | Durum haritası AYARLAR'da | `fut_durumHaritasi` |
| ERP iş emri ekranı | Geçersiz geçiş görünür uyarı | `fut_gecersizGecis` |
| WhatsApp teslim listesi | Açık bakiye ve fire% motoru | `fut_acikBakiye` |
| Elle kuyruk listesi | Kapalı kayıt kuyruktan düşer | `fut_kapaliDusum` |
| Elle numara defteri | Kimlik disiplini otomatik öneri | `fut_sonrakiKimlik` |

## En zayıf nokta

Örnek veri 28 akış satırıdır; gerçek fason dönemi yüzlerce satıra çıkabilir. Ölçek sözleşmesi SPEC’te 20.000’dir; dosya içi tablo kapasitesi 250’dir. Geçersiz geçiş makrosuz olduğu için engellenmez; kullanıcı uyarıyı görmezden gelebilir — bu sınır KILAVUZ’da yazılıdır.
