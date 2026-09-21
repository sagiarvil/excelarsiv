# FİYAT SAVUNMASI — PROJE VE İŞ BAZINDA GERÇEK KÂRLILIK SİSTEMİ

**Satış fiyatı:** 2.490 TL
**Denetim eşiği:** 3 × satış fiyatı = 7.470 TL (ikame maliyeti bu değerin üzerinde olmalı)

## D01 — İkame Maliyeti Çapası

| Bileşen | Tutar (₺) | Gerekçe |
|---|---|---|
| Y1 — Danışman ikamesi | 18.000 | İş/proje bazında kâr, marj, günlük kâr ayrıştırması, zararlı/riskli iş tespiti, senaryo bandı, tornado ve gelecek iş tahmini üreten bir mali müşavir/işletme danışmanının üç aylık sözleşme bedeli; kapsam olarak bu paketin ürettiği iş kârlılık raporunun eşdeğeri |
| **Y2 — Kurulum ikamesi** | 6.000 | Aynı modelin Excel'de sıfırdan kurulması: 1.000 satırlık iş tablosu, 40 adımlı hesap motoru, 8 grafikli pano, karar kapısı, senaryo/tornado katmanı ve veri kalite kontrollerinin tasarlanıp test edilmesi |
| **Toplam ikame** | **24.000** | 24.000 ≥ 7.470 ✓ |

**Y1+Y2 ≥ 3 × satış fiyatı:** 24.000 ≥ 7.470 → **GEÇTİ**

## D02 — Serbest Alternatif (Alıcı bunu bedavaya nasıl yapar?)

| Alıcının bedava yöntemi | Bu paketin ayrıştığı nokta | Dosyada karşılığı |
|---|---|---|
| İş kârlarını Excel'de elle toplamak; marj ve günlük kârı ayrı ayrı hesap etmek | İş bazında kâr, marj ve günlük kâr aynı tablodan canlı hesaplanır | `IsAgirlikliMarj` |
| Zararlı işleri rapor çıktısında görene dek fark etmemek | Eşik kuralları riskli/zararlı işleri otomatik işaretler; karar kapısı DURDUR üretir | `IsRiskliAdet` |
| Gelecek iş kârını "kör tahminle" söylemek | FORECAST/TREND ile tahmin + güven bandı; boş veride karar VERİ YOK | `modulI1Tahmin` |
| Hangi değişkenin kârlılığı en çok etkilediğini bilmeden fiyatlama yapmak | Tornado analizi gelir/maliyet etkisini önceliklendirir | `modulO2Tornado` |
| Senaryoları ayrı dosyalarda manuel kurmak | İyimser/baz/kötümser senaryo tek ekranda karşılaştırılır | `modulO3Senaryo` |
| Kararı kendi sezgisiyle vermek | UYGUN/İNCELE/DURDUR karar kapısı gerekçesiyle çalışır | `IsKarar` |

## En zayıf nokta (dürüst beyan)

Örnek veri 10 iş satırıdır; iş adedi ve dönem derinleştikçe tahmin ve yüzdelik katmanları daha güvenilir okunur. Günlük kâr, sürenin doğru girilmesine bağlıdır; süre yanlış girilirse günlük kâr ve karar etkilenir.

## Fiyat-performans gerekçesi

2.490 TL = bir işletme danışmanının bir günlük bedelinin altına denk gelir; karşılığında iş bazlı kârlılık raporu, riskli iş tespiti, senaryo bandı, tornado ve karar kapısı kalıcı olarak kullanıcının cihazında çalışır ve her iş kapanışında yeniden kullanılır.
