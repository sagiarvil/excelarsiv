# FİYAT SAVUNMASI — STOK, SATIŞ VE NAKİT BAĞLANMA SİSTEMİ

**Satış fiyatı:** 1.490 TL
**Denetim eşiği:** 3 × satış fiyatı = 4.470 TL (ikame maliyeti bu değerin üzerinde olmalı)

## D01 — İkame Maliyeti Çapası

| Bileşen | Tutar (₺) | Gerekçe |
|---|---|---|
| Y1 — Danışman ikamesi | 18.000 | Stok devir hızı, stoğa bağlanan nakit, ürün bazlı stok değeri, yenileme kararı ve nakit yönetimi modeli kuran bir işletme/mali danışmanın üç aylık sözleşme bedeli; kapsam olarak bu paketin ürettiği stok-nakit raporunun eşdeğeri |
| **Y2 — Kurulum ikamesi** | 6.000 | Aynı modelin Excel'de sıfırdan kurulması: 1.000 satırlık hareket tablosu, ürün bazlı hesap motoru, 8 grafikli pano, karar kapısı, senaryo/tornado katmanı ve veri kalite kontrollerinin tasarlanıp test edilmesi |
| **Toplam ikame** | **24.000** | 24.000 ≥ 4.470 ✓ |

**Y1+Y2 ≥ 3 × satış fiyatı:** 24.000 ≥ 4.470 → **GEÇTİ**

## D02 — Serbest Alternatif (Alıcı bunu bedavaya nasıl yapar?)

| Alıcının bedava yöntemi | Bu paketin ayrıştığı nokta | Dosyada karşılığı |
|---|---|---|
| Depo fişlerini elle toplayıp kalan stok, ortalama maliyet ve stok değerini ayrı araçlarda hesaplamak | Ürün bazında kalan stok, ortalama maliyet, stok değeri ve devir hızı aynı tablodan canlı hesaplanır | `StokDevirHizi` |
| Stoğa bağlanan nakdi görmemek | Stok değeri ve nakit etkisi otomatik ölçülür; eşik aşımında karar kapısı uyarır | `StokBagNakit` |
| Gelecek ay stok değerini "kör tahminle" söylemek | FORECAST ile tahmin + güven bandı; boş veride karar VERİ YOK | `modulI1Tahmin` |
| Hangi değişkenin stok değerini en çok etkilediğini bilmeden karar vermek | Tornado analizi fiyat/adet etkisini önceliklendirir | `modulO2Tornado` |
| Yoğunlaşmayı fark etmemek | Ürün yoğunlaşması HHI ile ölçülür | `modulO6Hhi` |
| Senaryoları ayrı dosyalarda manuel kurmak | İyimser/baz/kötümser senaryo tek ekranda karşılaştırılır | `modulO3Senaryo` |

## En zayıf nokta (dürüst beyan)

Örnek veri 6 hareket satırıdır; hareket adedi ve dönem derinleştikçe tahmin ve yüzdelik katmanları daha güvenilir okunur. Devir hızı, çıkışların doğru girilmesine bağlıdır; çıkışlar eksik girilirse devir hızı ve karar etkilenir.

## Fiyat-performans gerekçesi

1.490 TL = bir işletme danışmanının bir günlük bedelinin altına denk gelir; karşılığında stok devir raporu, stoğa bağlanan nakit görünümü, ürün bazlı stok değeri, senaryo bandı, tornado ve karar kapısı kalıcı olarak kullanıcının cihazında çalışır ve her hareket girişinde yeniden kullanılır.
