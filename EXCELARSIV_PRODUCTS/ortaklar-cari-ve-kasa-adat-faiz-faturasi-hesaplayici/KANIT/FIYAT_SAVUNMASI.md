# FİYAT SAVUNMASI — ORTAKLAR CARİ & KASA ADAT FAİZ FATURASI HESAPLAYICI

**Satış fiyatı:** 2.490 TL
**Denetim eşiği:** 3 × satış fiyatı = 7.470 TL (ikame maliyeti bu değerin üzerinde olmalı)

## D01 — İkame Maliyeti Çapası

| Bileşen | Tutar (₺) | Gerekçe |
|---|---|---|
| Y1 — Danışman ikamesi | 9.000 | Ortak cari bakiyelerinin dönem sonu mutabakatını yapan, adat ve faiz hesabını üreten ve örtülü sermaye/KKEG riskini değerlendiren bir mali müşavir/danışmanın dönemlik danışmanlık bedeli; kapsam olarak bu paketin ürettiği adat, faiz, fatura tutarı ve risk kararının eşdeğeri |
| **Y2 — Kurulum ikamesi** | 4.500 | Aynı modelin Excel'de sıfırdan kurulması: 1.000 satırlık ortak cari tablosu, gün bazlı adat motoru, kasa adat modülü, 8 grafikli pano, senaryo/tornado, tahmin bandı, karar kapısı ve veri kalite kontrollerinin tasarlanıp test edilmesi |
| **Toplam ikame** | **13.500** | 13.500 ≥ 7.470 ✓ |

**Y1+Y2 ≥ 3 × satış fiyatı:** 13.500 ≥ 7.470 → **GEÇTİ**

## D02 — Serbest Alternatif (Alıcı bunu bedavaya nasıl yapar?)

| Alıcının bedava yöntemi | Bu paketin ayrıştığı nokta | Dosyada karşılığı |
|---|---|---|
| Ortak cari bakiyelerini ayrı ekranlarda elle toplayıp faiz hesabını hesap makinesiyle yapmak | Hareketler tek tabloda tarih sıralı tutulur; bakiyeyi her işlemden sonra formül canlı üretir | `OrtakCariNetBakiye` |
| Kasa gün sonu bakiyelerini deftere elle işleyip adat formülünü kendisi kurmak | Kasa adatı ayrı tabloda gün bazlı hesaplanır; toplam tek motorda birleşir | `ToplamAdat` |
| Dönem sonunda faturayı tahmini tutarla kesip mutabakat farkını sonradan düzeltmek | Faiz faturası tutarı net hesaplanır; girişler eksiksizse fark doğmaz | `BrutFaizTutari` |
| Örtülü sermaye ve KKEG riskini hiç izlemeyip vergi incelemesine kadar fark etmemek | Ortak cari/öz kaynak oranı eşikte ise karar kapısı İNCELE der | `modulO1Anomali` |
| Faiz oranını emsal verilerle karşılaştırmadan sabit bir oranla kesmek | Oran ve vade değişimlerinin fatura tutarına etkisi tornado ile önceliklendirilir | `modulO2Tornado` |
| Gelecek dönem faizini tahmin edememek | FORECAST ile tahmin + güven bandı sunulur | `modulI1AdatTahmin` |

## En zayıf nokta (dürüst beyan)

Örnek veri 6 hareket satırıdır; hareket sayısı ve dönem derinleştikçe tahmin ve yüzdelik katmanları daha güvenilir okunur. Adat hesabı, gün sonu bakiyelerinin eksiksiz kaydedilmesine bağlıdır; eksik kasa kaydı adatı düşürür (veri kalite skoru ve karar kapısı bunu İNCELE olarak işaretler). Faiz oranı ve eşikler kullanıcı varsayımıdır; dosya oranların mevzuata uygunluğunu taahhüt etmez, emsal mutabakatını kullanıcının yapması beklenir.

## Fiyat-performans gerekçesi

2.490 TL = bir mali müşavirin dönemlik adat/faiz mutabakat ücretinin küçük bir kısmına denk gelir; karşılığında ortak cari bakiyesi, gün bazlı adat, faiz faturası tutarı, örtülü sermaye/KKEG risk izleme, senaryo bandı, tornado, tahmin ve karar kapısı kalıcı olarak kullanıcının cihazında çalışır ve her dönem hareket girişinde yeniden kullanılır.
