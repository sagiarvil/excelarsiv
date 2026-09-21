# FİYAT SAVUNMASI — ŞİRKET ÖZ KAYNAĞI ERİDİ Mİ? (TTK 376 SERMAYE TAMAMLAMA CETVELİ)

**Satış fiyatı:** 7.900 TL
**Denetim eşiği:** 3 × satış fiyatı = 23.700 TL (ikame maliyeti bu değerin üzerinde olmalı)

## D01 — İkame Maliyeti Çapası

| Bileşen | Tutar (₺) | Gerekçe |
|---|---|---|
| Y1 — Danışman ikamesi | 18.000 | Aylık bilanço bileşenlerinden öz kaynak hesabı, 1/3-2/3 sermaye kaybı izleme, borca batıklık tespiti ve genel kurul/karar dokümanlarının hazırlanması için mali müşavir/hukuk danışmanının dönemlik ücreti; kapsam olarak bu paketin ürettiği öz kaynak, kayıp oranı, kayıp tutarı, eşik kırılımları ve karar çıktılarının eşdeğeri |
| **Y2 — Kurulum ikamesi** | 6.000 | Aynı modelin Excel'de sıfırdan kurulması: 1.000 satırlık bilanço tablosu, borç listesi, eşik motoru, 8 grafikli pano, senaryo/tornado, tahmin bandı, karar kapısı ve veri kalite kontrollerinin tasarlanıp test edilmesi |
| **Toplam ikame** | **24.000** | 24.000 ≥ 23.700 ✓ |

**Y1+Y2 ≥ 3 × satış fiyatı:** 24.000 ≥ 23.700 → **GEÇTİ**

## D02 — Serbest Alternatif (Alıcı bunu bedavaya nasıl yapar?)

| Alıcının bedava yöntemi | Bu paketin ayrıştığı nokta | Dosyada karşılığı |
|---|---|---|
| Aylık bilanço bileşenlerinden öz kaynağı, sermaye kaybını ve 1/3-2/3 eşiklerini elle hesaplamak | Aylık bilanço bileşenleri tek tabloda toplanır; öz kaynak, kayıp tutarı ve kayıp oranı formülle canlı üretilir | `TtkKayipOrani` |
| Borca batıklığı ve yönetim kurulu bildirim yükümlülüğünü genel kurula kadar hiç izlememek | TTK 376 1/3 ve 2/3 eşikleri ve borca batıklık mevzuat oranlarıyla otomatik kontrol edilir; alternatifte hesap elle yapılır | `TtkBorcaBatik` |
| Sermaye artırımı/azaltımı/infisah senaryolarını ayrı dosyalarda tahmini değerlerle kurmak | Borç türü kırılımında yoğunlaşma (HHI) ve anomali analizleri üretilir; alternatifte bu analiz yoktur | `modulO6Hhi` |
| Borç yoğunlaşmasını ve kayıp oranı trendini hiç analiz etmeden dönem sonunu beklemek | İyimser/baz/kötümser senaryolarda kayıp oranı bandı ve gelecek dönem öz kaynak tahmini sunulur; alternatifte tahmin yoktur | `modulI1OzKaynakTahmin` |
| Gelecek dönem öz kaynak tahminini yapmadan mali tablo sunumu hazırlamak | Karar kapısı 1/3-2/3 eşiklerini aşan veya borca batık durumda İNCELE/DURDUR üretir ve genel kurul aksiyonu önerir | `modulO8KaliteSkor` |
| Güven skoru olmadan hatalı/eksik bilanço girişleriyle karar almak | Güven skoru eksik/negatif girişleri yakalayarak veri kalitesini puanlar; alternatifte hatalı giriş fark edilmez | `modulO8KaliteSkorYorum` |

## En zayıf nokta (dürüst beyan)

Örnek veri 12 aylık bilanço satırı ve 6 borç satırıdır; veri derinleştikçe tahmin ve yüzdelik katmanları daha güvenilir okunur. Öz kaynak ve kayıp oranı hesabı, dönem bilanço bileşenlerinin ve borç listesinin eksiksiz kaydedilmesine bağlıdır; eksik veya yanlış giriş kayıp oranını ve karar kapısını değiştirir (veri kalite skoru ve karar kapısı bunu İNCELE olarak işaretler). Eşik oranları (%33,33 / %66,67) TTK 376'dan mevzuat kaynaklıdır; maddeye ilişkin güncel içtihat ve tebliğ değişikliklerini kullanıcının takip etmesi beklenir. Dosya kesin hukuki tespit değil, karar destek aracıdır.

## Fiyat-performans gerekçesi

7.900 TL = bir mali müşavirin dönemlik öz kaynak/sermaye kaybı analizi ve genel kurul dokümantasyonu ücretinin küçük bir kısmına denk gelir; karşılığında aylık bilanço takibi, kayıp oranı, eşik kırılımı, borca batıklık, senaryo bandı, tornado, tahmin ve karar kapısı kalıcı olarak kullanıcının cihazında çalışır ve her dönem veri girişinde yeniden kullanılır.
