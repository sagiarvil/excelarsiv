# FİYAT SAVUNMASI — KOBİ FİNANS YÖNETİM PAKETİ

**Satış fiyatı:** 7.900 TL
**Denetim eşiği:** 3 × satış fiyatı = 23.700 TL (ikame maliyeti bu değerin üzerinde olmalı)

## D01 — İkame Maliyeti Çapası

| Bileşen | Tutar (₺) | Gerekçe |
|---|---|---|
| Y1 — Danışman ikamesi | 12.000 | Aylık nakit akışını dört tahsilat kanalıyla konsolide eden, senaryo + tahmin + risk skoru üreten bir finans danışmanının bir aylık sözleşme bedeli; kapsam olarak bu paketin ürettiği aylık yönetim raporu, senaryo bandı ve tahmin katmanının eşdeğeri |
| **Y2 — Kurulum ikamesi** | 12.000 | Aynı modelin Excel'de sıfırdan kurulması: tablo mimarisi, 41 adımlı hesap motoru, 8 grafikli pano, karar motoru ve veri kalite kontrollerinin tasarlanıp test edilmesi |
| **Toplam ikame** | **24.000** | 24.000 ≥ 23.700 ✓ |

**Y1+Y2 ≥ 3 × satış fiyatı:** 24.000 ≥ 23.700 → **GEÇTİ**

## D02 — Serbest Alternatif (Alıcı bunu bedavaya nasıl yapar?)

| Alıcının bedava yöntemi | Bu paketin ayrıştığı nokta | Dosyada karşılığı |
|---|---|---|
| Gelir/giderlerin ayrı Excel listelerinde elle tutulması | Dört tahsilat kanalı tek modelde konsolide, kapanış otomatik | `KobiNakitGiris` |
| Banka internet şubesi dökümüyle nakit takibi | Kasa/cari/POS kanalları da görünür; risk skoru eşiğe bağlı | `KobiRiskSkoru` |
| Dönem sonunda mali müşavirden tablo beklemek | İyimser/baz/kötümser senaryo bandı anlık | `KobiIyiSenaryo` |
| Nakit akışını yalnızca giriş-çıkış farkıyla izlemek | Tornado ile hangi değişkenin etkisi büyük önceliklendirilir | `KobiTornadoTahsilat` |
| Aylık bütçe tablosunu sıfırdan yeniden kurmak | FORECAST tahmini + veri kalitesine bağlı karar (VERİ YOK) | `KobiTahminiKar` |

## En zayıf nokta (dürüst beyan)

Örnek veri 6 aylık olduğundan uzun dönemli (12 ay üstü) sezonellik tam gözlenemez; tahmin katmanı mevcut dönem eğilimine dayanır. Veri derinleştikçe tahmin aralığı güvenilirliği artar.

## Fiyat-performans gerekçesi

7.900 TL = bir finans danışmanının iki günlük bedelinin altına denk gelir; karşılığında aylık yönetim raporu, senaryo bandı, risk skoru ve veri kalitesi kontrolü kalıcı olarak kullanıcının cihazında çalışır.
