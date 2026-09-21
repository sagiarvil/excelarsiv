# FİYAT SAVUNMASI — POS, KOMİSYON VE NET TAHSİLAT KONTROL SİSTEMİ

**Satış fiyatı:** 990 TL
**Denetim eşiği:** 3 × satış fiyatı = 2.970 TL (ikame maliyeti bu değerin üzerinde olmalı)

## D01 — İkame Maliyeti Çapası

| Bileşen | Tutar (₺) | Gerekçe |
|---|---|---|
| Y1 — Danışman ikamesi | 6.000 | Çoklu POS kanalının brüt/iade/komisyon ayrıştırmasını, en pahalı kanal tespitini, senaryo ve tahmin bandını üreten bir perakende finans danışmanının bir aylık sözleşme bedeli; kapsam olarak bu paketin ürettiği kanal maliyet raporu ve tahmin katmanının eşdeğeri |
| **Y2 — Kurulum ikamesi** | 2.500 | Aynı modelin Excel'de sıfırdan kurulması: 1.000 satırlık işlem tablosu, 49 adımlı hesap motoru, 6 grafikli pano, karar kapısı ve veri kalite kontrollerinin tasarlanıp test edilmesi |
| **Toplam ikame** | **8.500** | 8.500 ≥ 2.970 ✓ |

**Y1+Y2 ≥ 3 × satış fiyatı:** 8.500 ≥ 2.970 → **GEÇTİ**

## D02 — Serbest Alternatif (Alıcı bunu bedavaya nasıl yapar?)

| Alıcının bedava yöntemi | Bu paketin ayrıştığı nokta | Dosyada karşılığı |
|---|---|---|
| Kanal bazlı net tahsilatın elle toplanması; komisyon/iade ayrımı yapılmadan tek toplam takibi | Kanal bazlı brüt, iade ve komisyon ayrıştırması otomatik; net tahsilat canlı | `PosNetTahsilatToplam` |
| POS ekstrelerini PDF arşivleyip aylık komisyon kontrolünü manuel yapmak | En pahalı kanal otomatik belirlenir; komisyon eşiği aşımında karar DURDUR | `PosEnPahaliKanal` |
| Tüm kanallara aynı komisyon eşiği uygulamak | İyimser/baz/kötümser senaryo bandı + tornado ile kanal etkisi önceliklendirilir | `PosKomisyonOrani` |
| Komisyon oranlarını sözleşme metinlerinde aramak | İade ve komisyon risk eşikleri AYARLAR'dan tek yerden yönetilir, aşım uyarır | `PosIadeOrani` |
| İade riskini görmeden yüksek iade kanalına aynı kotayı vermek | FORECAST ile gelecek ay tahmini + güven bandı; boş veride karar VERİ YOK | `modulI1Tahmin` |

## En zayıf nokta (dürüst beyan)

Örnek veri 12 işlem satırıdır; kanal bazlı aylık düzenlilik veri derinleştikçe daha güvenilir okunur. Tahmin katmanı mevcut eğilime dayanır; mevsimsel POS dalgalanması uzun dönem veriyle daha iyi modellenir.

## Fiyat-performans gerekçesi

990 TL = bir perakende finans danışmanının bir günlük bedelinin altına denk gelir; karşılığında kanal maliyet raporu, en pahalı kanal tespiti, senaryo bandı ve veri kalitesi kontrolü kalıcı olarak kullanıcının cihazında çalışır.
