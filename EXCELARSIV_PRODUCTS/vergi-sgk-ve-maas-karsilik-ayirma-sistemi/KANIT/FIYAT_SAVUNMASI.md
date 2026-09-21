# FİYAT SAVUNMASI — VERGİ, SGK VE MAAŞ KARŞILIK AYIRMA SİSTEMİ

**Satış fiyatı:** 990 TL
**Denetim eşiği:** 3 × satış fiyatı = 2.970 TL (ikame maliyeti bu değerin üzerinde olmalı)

## D01 — İkame Maliyeti Çapası

| Bileşen | Tutar (₺) | Gerekçe |
|---|---|---|
| Y1 — Danışman ikamesi | 18.000 | Dönem gelirine göre KDV, kurumlar vergisi, SGK, işsizlik, stopaj, damga ve maaş karşılıklarını ayırıp nakit planını koruyan bir mali müşavir/danışmanın üç aylık sözleşme bedeli; kapsam olarak bu paketin ürettiği karşılık ve nakit raporunun eşdeğeri |
| **Y2 — Kurulum ikamesi** | 6.000 | Aynı modelin Excel'de sıfırdan kurulması: 1.000 satırlık dönem/bordro tabloları, karşılık motoru, 8 grafikli pano, karar kapısı, senaryo/tornado ve tahmin katmanının tasarlanıp test edilmesi |
| **Toplam ikame** | **24.000** | 24.000 ≥ 2.970 ✓ |

**Y1+Y2 ≥ 3 × satış fiyatı:** 24.000 ≥ 2.970 → **GEÇTİ**

## D02 — Serbest Alternatif (Alıcı bunu bedavaya nasıl yapar?)

| Alıcının bedava yöntemi | Bu paketin ayrıştığı nokta | Dosyada karşılığı |
|---|---|---|
| Vergi/SGK/maaş karşılıklarını elle Excel'de ayrı ayrı hesaplamak (işçilik saati) | KDV, KV, SGK, işsizlik, stopaj, damga ve maaş karşılıkları tek motorda aynı anda hesaplanır | `KarsilikToplam` |
| Karşılık sonrası nakit etkisini görmemek | Nakit fazlası otomatik hesaplanır; eşik altında karar kapısı uyarır | `KarsilikNakitFazla` |
| Gelecek ay gelirini "kör tahminle" söylemek | FORECAST ile tahmin + güven bandı; boş veride karar VERİ YOK | `modulI1Tahmin` |
| Hangi değişkenin karşılığı en çok etkilediğini bilmeden karar vermek | Tornado analizi KDV/KV/SGK/gelir etkisini önceliklendirir | `modulO2Tornado` |
| Senaryoları ayrı dosyalarda manuel kurmak | İyimser/baz/kötümser senaryo tek ekranda karşılaştırılır | `modulO3Senaryo` |
| Dönemler arası trendi ve anomaliyi görememek | Trend farkı ve anomali oranı otomatik ölçülür | `modulO1Anomali` |

## En zayıf nokta (dürüst beyan)

Örnek veri 6 aylık tek dönemdir; dönem derinleştikçe tahmin ve yüzdelik katmanları daha güvenilir okunur. Kararın doğruluğu dönem gelir ve giderinin doğru girilmesine bağlıdır; eksik veya gecikmiş girişler karşılık oranını ve kararı etkiler. Nakit karşılık oranı (%25) işletme bazlı varsayımdır; sektörel olarak güncellenmelidir.

## Fiyat-performans gerekçesi

990 TL = bir mali danışmanın bir günlük bedelinin altına denk gelir; karşılığında vergi/SGK/maaş karşılık ayrımı, nakit görünümü, senaryo bandı, tornado, tahmin ve karar kapısı kalıcı olarak kullanıcının cihazında çalışır ve her dönem girişinde yeniden kullanılır.
