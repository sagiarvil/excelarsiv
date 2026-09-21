# FİYAT SAVUNMASI — KDV İADESİ AZAMİ ALACAK HESABI & DOSYA HAZIRLAYICI

**Satış fiyatı:** 7.900 TL
**Denetim eşiği:** 3 × satış fiyatı = 23.700 TL (ikame maliyeti bu değerin üzerinde olmalı)

## D01 — İkame Maliyeti Çapası

| Bileşen | Tutar (₺) | Gerekçe |
|---|---|---|
| Y1 — Danışman ikamesi | 18.000 | KDV iadesi dosyasını hazırlayan, devreden KDV ile iade sınırını hesaplayan ve belge kontrolü yapan bir mali müşavir/danışmanın sezonluk danışmanlık bedeli; kapsam olarak bu paketin ürettiği iade-alacak ve dosya hazırlık çıktısının eşdeğeri |
| **Y2 — Kurulum ikamesi** | 6.000 | Aynı modelin Excel'de sıfırdan kurulması: 1.000 satırlık KDV belgesi tablosu, belge türü bazlı hesap motoru, 8 grafikli pano, senaryo/tornado, tahmin bandı, karar kapısı ve veri kalite kontrollerinin tasarlanıp test edilmesi |
| **Toplam ikame** | **24.000** | 24.000 ≥ 23.700 ✓ |

**Y1+Y2 ≥ 3 × satış fiyatı:** 24.000 ≥ 23.700 → **GEÇTİ**

## D02 — Serbest Alternatif (Alıcı bunu bedavaya nasıl yapar?)

| Alıcının bedava yöntemi | Bu paketin ayrıştığı nokta | Dosyada karşılığı |
|---|---|---|
| Devreden KDV ve iade sınırını ayrı ekranlarda elle hesaplamak | Belge türü bazında hesaplanan/indirilecek/devreden KDV aynı motorda canlı hesaplanır | `AzamiIadeAlacak` |
| Azami iade alacağını tahmin etmeden beyan dönemini beklemek | Devreden KDV ile ihracat iade sınırının kesişimi otomatik üretilir | `HesaplananKdv` |
| Eksik belgeyi iade süreci başladıktan sonra fark etmek | Eksik belge sayısı canlı listelenir, karar kapısı İNCELE der | `EksikBelgeSayisi` |
| Senaryoları ayrı dosyalarda manuel kurmak | İyimser/baz/kötümser iade senaryosu tek ekranda karşılaştırılır | `modulO3Senaryo` |
| Hangi değişkenin iadeyi en çok etkilediğini bilmemek | Tornado analizi devreden/ihracat etkisini önceliklendirir | `modulO2Tornado` |
| Gelecek dönem devreden KDV'yi tahmin edememek | FORECAST ile tahmin + güven bandı | `modulI1IadeTahmin` |

## En zayıf nokta (dürüst beyan)

Örnek veri 10 belge satırıdır; belge sayısı ve dönem derinleştikçe tahmin ve yüzdelik katmanları daha güvenilir okunur. Azami iade tutarı, belge türü girişlerinin doğruluğuna bağlıdır; yanlış belge türü ayrışması sonucu bozar (bu durum veri kalite skoru ve karar kapısı tarafından İNCELE olarak işaretlenir). İade tutarı vergi dairesinin inceleme ve belge teyidine bağlıdır; dosya kesin iade tutarı taahhüt etmez.

## Fiyat-performans gerekçesi

7.900 TL = bir mali müşavirin KDV iade dosyası hazırlık ücretinin bir dönemlik kısmına denk gelir; karşılığında azami iade alacağı, devreden KDV analizi, senaryo bandı, tornado, tahmin, eksik belge kontrolü ve karar kapısı kalıcı olarak kullanıcının cihazında çalışır ve her dönem belge girişinde yeniden kullanılır.
