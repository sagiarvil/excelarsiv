# FİYAT SAVUNMASI — Mini-MRP (BOM + Malzeme İhtiyacı + Kapasite)

**Satış fiyatı:** 9.900 TL
**Denetim eşiği:** 3 × satış fiyatı = 29.700 TL (ikame maliyeti bu değerin üzerinde olmalı)

## D01 — İkame Maliyeti Çapası

| Bileşen | Tutar (₺) | Gerekçe |
|---|---|---|
| Y1 — Danışman ikamesi | 20.000 | MRP/BOM planlama danışmanlığının üretim + satın alma + mali işler üçlüsünde kurulması ve dönem kanıt dosyasının hazırlanması |
| Y2 — Kurulum ikamesi | 10.000 | Durum makinesi, brüt/net ihtiyaç, kapasite yükü, önerilen sipariş, 8 grafikli pano ve kanıt raporunun sıfırdan Excel'de kurulması |
| **Toplam ikame** | **30.000** | 30.000 ≥ 29.700 ✓ |

**Y1+Y2 ≥ 3 × satış fiyatı:** 30.000 ≥ 29.700 → **GEÇTİ**

## D02 — Serbest Alternatif

| Bedava yöntem | Ayrım | Dosya karşılığı |
|---|---|---|
| Elle Excel BOM listesi | Durum haritası AYARLAR'da | `mmr_durumHaritasi` |
| ERP MRP ekranı | Geçersiz geçiş görünür uyarı | `mmr_gecersizGecis` |
| Satın alma sipariş defteri | Kapasite yükü % ve aşım | `mmr_kapasiteYuku` |
| Elle kuyruk listesi | Kapalı kayıt kuyruktan düşer | `mmr_kapaliDusum` |
| Elle numara defteri | Kimlik disiplini otomatik öneri | `mmr_sonrakiKimlik` |

## En zayıf nokta

Örnek veri 28 akış satırıdır; gerçek MRP dönemi yüzlerce satıra çıkabilir. Ölçek sözleşmesi SPEC'te 20.000'dır; dosya içi tablo kapasitesi 250'dir. Geçersiz geçiş makrosuz olduğu için engellenmez; kullanıcı uyarıyı görmezden gelebilir — bu sınır KILAVUZ'da yazılıdır.
