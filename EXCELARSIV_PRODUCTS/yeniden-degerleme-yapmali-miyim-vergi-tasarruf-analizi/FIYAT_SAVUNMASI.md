# FİYAT SAVUNMASI — YENİDEN DEĞERLEME YAPMALI MIYIM? (VERGİ TASARRUF ANALİZİ)

**Satış fiyatı:** 2.490 TL
**Denetim eşiği:** 3 × satış fiyatı = 7.470 TL (ikame maliyeti bu değerin üzerinde olmalı)

## D01 — İkame Maliyeti Çapası

| Bileşen | Tutar (₺) | Gerekçe |
|---|---|---|
| Y1 — Danışman ikamesi | 15.000 | Kıymet listesinden yeniden değerleme katsayısıyla değerlenmiş net değer, değer artışı, %2 fon vergisi ve ek amortismanın vergi tasarrufunun bugünkü değerini kıymet bazında hesaplaması için mali müşavir/danışmanın dönemlik ücreti; kapsam olarak bu paketin ürettiği NBD, net fayda, senaryo bandı ve karar çıktılarının eşdeğeri |
| **Y2 — Kurulum ikamesi** | 6.000 | Aynı modelin Excel'de sıfırdan kurulması: 1.000 satırlık kıymet listesi, %2 fon hesabı, iskonto edilmiş tasarruf modeli, senaryo/tornado, tahmin bandı, 8 grafikli pano ve karar kapısının tasarlanıp test edilmesi |
| **Toplam ikame** | **21.000** | 21.000 ≥ 7.470 ✓ |

**Y1+Y2 ≥ 3 × satış fiyatı:** 21.000 ≥ 7.470 → **GEÇTİ**

## D02 — Serbest Alternatif (Alıcı bunu bedavaya nasıl yapar?)

| Alıcının bedava yöntemi | Bu paketin ayrıştığı nokta | Dosyada karşılığı |
|---|---|---|
| Kıymet bazında değer artışı, %2 fon vergisi ve ek amortismanı elle, ayrı kalemlerde hesaplamak | Kıymet bazında değer artışı, %2 fon vergisi ve ek amortisman hesabı tek tabloda formülle üretilir | `YenidegerDegerArtisiToplam` |
| Ek amortismanın kalan ömür boyunca vergi tasarrufunu iskonto etmeden "kabaca" tahmin etmek | Ek amortismanın kalan ömür boyunca sağladığı vergi tasarrufu iskonto edilerek bugünkü değerine indirgenir; alternatifte NBD hesabı yoktur | `YenidegerTasarrufBdToplam` |
| İskonto/fon oranı değişimini hiç test etmeden tek senaryoyla karar vermek | İyimser/baz/kötümser/kritik senaryo bandı ve tornado ile duyarlılık gösterilir; alternatifte duyarlılık analizi yoktur | `modulO3Senaryo` |
| Kıymet türü yoğunlaşmasını ve uç değerleri analiz etmeden portföyü tek tek incelemek | Kıymet türü bazında değer artışı yoğunlaşması (HHI) ve net fayda anomali sapması ölçülür; alternatifte bu analiz yoktur | `modulO6Hhi` |
| Net fayda ve veri kalitesi skoru olmadan yalnızca katsayıyla karar vermek | Karar kapısı net fayda ve değer artışı eşiklerine göre UYGUN/İNCELE/DURDUR üretir ve aksiyon önerir; veri kalite skoru eksik girişleri yakalar | `modulO8KaliteSkor` |

## En zayıf nokta (dürüst beyan)

Örnek veri 10 kıymet satırıdır; kıymet sayısı ve kalan faydalı ömür bilgisi arttıkça tahmin ve yüzdelik katmanları daha güvenilir okunur. Yeniden değerleme katsayısı her yıl resmi olarak ilan edilir; kullanıcının AYARLAR'dan güncellemesi gerekir (dosya güncel katsayıyı kendiliğinden getirmez). Karar kapısı net faydanın pozitifliğine bakar; fon oranı ve kurumlar vergisi oranı değiştiğinde sonuç değişir. Dosya kesin vergi görüşü değil, karar destek aracıdır.

## Fiyat-performans gerekçesi

2.490 TL = bir mali müşavirin yeniden değerleme dosyası hazırlama ücretinin küçük bir kısmına denk gelir; karşılığında kıymet bazlı net fayda, %2 fon vergisi, iskonto edilmiş tasarruf, senaryo bandı, tornado, tahmin ve karar kapısı kalıcı olarak kullanıcının cihazında çalışır ve her dönem veri girişinde yeniden kullanılır.
