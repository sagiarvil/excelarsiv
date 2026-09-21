# FİYAT SAVUNMASI — KKEG VE FİNANSMAN GİDER KISITLAMASI VERGİ SAVUNMA SETİ

**Satış fiyatı:** 7.900 TL
**Denetim eşiği:** 3 × satış fiyatı = 23.700 TL (ikame maliyeti bu değerin üzerinde olmalı)

## D01 — İkame Maliyeti Çapası

| Bileşen | Tutar (₺) | Gerekçe |
|---|---|---|
| Y1 — Danışman ikamesi | 18.000 | KKEG kapsamının, kapsam dışı giderlerin, örtülü sermaye ve finansman gider kısıtlamasının dönem dönem hesaplanması ve incelemeye karşı savunma dosyasının hazırlanması için mali müşavir/vergi danışmanının dönemlik ücreti; kapsam olarak bu paketin ürettiği öz/yabancı kaynak ortalaması, aşan kısma isabet eden gider, KKEG tutarı, vergi etkisi ve karar çıktılarının eşdeğeri |
| **Y2 — Kurulum ikamesi** | 6.000 | Aynı modelin Excel'de sıfırdan kurulması: 1.000 satırlık dönem tablosu, kredi listesi, kısıtlama motoru, 8 grafikli pano, senaryo/tornado, tahmin bandı, karar kapısı ve veri kalite kontrollerinin tasarlanıp test edilmesi |
| **Toplam ikame** | **24.000** | 24.000 ≥ 23.700 ✓ |

**Y1+Y2 ≥ 3 × satış fiyatı:** 24.000 ≥ 23.700 → **GEÇTİ**

## D02 — Serbest Alternatif (Alıcı bunu bedavaya nasıl yapar?)

| Alıcının bedava yöntemi | Bu paketin ayrıştığı nokta | Dosyada karşılığı |
|---|---|---|
| KKEG kapsamındaki gideri mali müşavire her dönem ücret ödeyerek hesaplatmak | Dönem bazlı öz kaynak, yabancı kaynak ve finansman gideri tek tabloda toplanır; ortalamalar formülle canlı üretilir | `OrtalamaOzKaynak` |
| Yabancı kaynak/öz kaynak ortalamasını ve aşan kısma isabet eden gideri ayrı dosyalarda elle kurmak | Kredi bazlı gider listesi ile kapsam dışı işaretleri ayrı tabloda tutulur; kapsamdaki gider kredi kırılımında gösterilir | `ToplamKapsamGider` |
| Kredi listesini ve kapsam dışı işaretlerini defter kayıtlarında hiç izlememek | Aşan kısma isabet eden gider ve KKEG tutarı mevzuat oranlarıyla hesaplanır; alternatifte hesap elle yapılır | `KkegTutari` |
| Kısıtlama oranı değişiminin vergi yüküne etkisini tahmin etmeden dönem sonunu beklemek | Kısıtlama oranı ve yabancı kaynak değişiminin KKEG'e etkisi tornado ile önceliklendirilir | `modulO2Tornado` |
| Finansman gider ayrıştırmasını inceleme gelene kadar hiç belgelememek | İyimser/baz/kötümser senaryolarda KKEG bandı ve gelecek dönem gider tahmini sunulur | `modulI1GiderTahmin` |
| Yabancı kaynak yoğunluğunu ve kalite bozukluğunu hiç izlemeyip vergi incelemesine kadar fark etmemek | Karar kapısı, kapsamdaki gider ve KKEG eşiklerini aşan durumda İNCELE/DURDUR üretir ve savunma aksiyonu önerir | `modulO8KaliteSkor` |

## En zayıf nokta (dürüst beyan)

Örnek veri 6 aylık dönem satırı ve 2 kredi satırıdır; veri derinleştikçe tahmin ve yüzdelik katmanları daha güvenilir okunur. KKEG hesabı, dönem kaynak değerlerinin ve kredi listesinin eksiksiz kaydedilmesine bağlıdır; eksik veya yanlış kapsam dışı işareti KKEG tutarını değiştirir (veri kalite skoru ve karar kapısı bunu İNCELE olarak işaretler). Kısıtlama oranı (%25) ve kurumlar vergisi oranı güncel mevzuata göre AYARLAR'da doğrulanmalıdır; dosya oranların güncelliğini taahhüt etmez, kullanıcının mevzuat güncellemelerini takip etmesi beklenir.

## Fiyat-performans gerekçesi

7.900 TL = bir mali müşavirin dönemlik KKEG ve finansman gider kısıtlaması hesabı ücretinin küçük bir kısmına denk gelir; karşılığında dönem bazlı öz/yabancı kaynak ortalaması, kısıtlama kapsamı, KKEG tutarı, kurumlar vergisi etkisi, kredi kırılımı, senaryo bandı, tornado, tahmin ve karar kapısı kalıcı olarak kullanıcının cihazında çalışır ve her dönem veri girişinde yeniden kullanılır.
