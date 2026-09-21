# FİYAT SAVUNMASI — DÖVİZ AÇIK POZİSYONU VE KUR RİSKİ STRES TESTİ

**Satış fiyatı:** 2.490 TL
**Denetim eşiği:** 3 × satış fiyatı = 7.470 TL (ikame maliyeti bu değerin üzerinde olmalı)

## D01 — İkame Maliyeti Çapası

| Bileşen | Tutar (₺) | Gerekçe |
|---|---|---|
| Y1 — Danışman ikamesi | 18.000 | Dövizli alacak-borç yapısını döviz cinsi bazında netleştiren, kur şoku stres testi kuran ve açık pozisyon politikası öneren bir kur riski/treasury danışmanının üç aylık sözleşme bedeli; kapsam olarak bu paketin ürettiği pozisyon-stres raporunun eşdeğeri |
| **Y2 — Kurulum ikamesi** | 6.000 | Aynı modelin Excel'de sıfırdan kurulması: 1.000 satırlık döviz pozisyon tablosu, cinsi bazlı hesap motoru, 8 grafikli pano, stres senaryoları, tornado, tahmin bandı ve veri kalite kontrollerinin tasarlanıp test edilmesi |
| **Toplam ikame** | **24.000** | 24.000 ≥ 7.470 ✓ |

**Y1+Y2 ≥ 3 × satış fiyatı:** 24.000 ≥ 7.470 → **GEÇTİ**

## D02 — Serbest Alternatif (Alıcı bunu bedavaya nasıl yapar?)

| Alıcının bedava yöntemi | Bu paketin ayrıştığı nokta | Dosyada karşılığı |
|---|---|---|
| Döviz pozisyonlarını ayrı ekranlarda elle toplamak | Döviz cinsi bazında varlık, borç ve net açık pozisyon aynı motorda canlı hesaplanır | `NetAcikPozisyon` |
| Kur şokunda etkiyi "kör tahminle" söylemek | Kötümser ve kritik senaryolarda TL etkisi stres testiyle üretilir | `modulO3Senaryo` |
| Hangi döviz cinsinin riski en çok artırdığını bilmemek | Tornado analizi kur/pozisyon etkisini önceliklendirir | `modulO2Tornado` |
| Gelecek kur seviyesini tahmin edememek | FORECAST ile tahmin + güven bandı | `modulI1KurTahmin` |
| Yoğunlaşmayı fark etmemek | Döviz yoğunlaşması HHI ile ölçülür | `modulO6Hhi` |
| Senaryoları ayrı dosyalarda manuel kurmak | İyimser/baz/kötümser/kritik senaryo tek ekranda karşılaştırılır | `modulO3Senaryo` |

## En zayıf nokta (dürüst beyan)

Örnek veri 10 pozisyon satırıdır; işlem sayısı ve dönem derinleştikçe tahmin ve yüzdelik katmanları daha güvenilir okunur. Kur riski hesabı, işlem kurlarının doğru girilmesine bağlıdır; kurlar güncel TCMB kurlarından girilmelidir. Net pozisyon varlık-borç farkı olduğu için, varlık ve borcun dengelenmesi durumunda mutlak pozisyon küçük görünebilir; bu durumda işlem hacmi eşiği devreye girer.

## Fiyat-performans gerekçesi

2.490 TL = bir kur riski danışmanının yarım günlük bedeline denk gelir; karşılığında net açık pozisyon, stres etkisi, senaryo bandı, tornado, tahmin ve karar kapısı kalıcı olarak kullanıcının cihazında çalışır ve her pozisyon girişinde yeniden kullanılır.
