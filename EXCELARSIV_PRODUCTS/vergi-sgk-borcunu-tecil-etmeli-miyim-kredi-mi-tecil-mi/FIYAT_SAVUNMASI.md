# FİYAT SAVUNMASI — VERGİ/SGK BORCUNU TECİL ETMELİ MİYİM? (KREDİ Mİ TECİL Mİ?)

**Satış fiyatı:** 990 TL
**Denetim eşiği:** 3 × satış fiyatı = 2.970 TL (ikame maliyeti bu değerin üzerinde olmalı)

## D01 — İkame Maliyeti Çapası

| Bileşen | Tutar (₺) | Gerekçe |
|---|---|---|
| Y1 — Danışman ikamesi | 4.000 | Tecil başvurusu öncesi tecil, kredi ve peşin ödeme seçeneklerini taksit formülüyle karşılaştırıp NBD'ye indirgeyerek en düşük maliyetli yolu raporlayan mali müşavir/danışman çalışmasının ücreti; kapsam olarak bu paketin ürettiği maliyet, NBD, senaryo bandı ve karar çıktılarının eşdeğeri |
| **Y2 — Kurulum ikamesi** | 1.000 | Aynı modelin Excel'de sıfırdan kurulması: 1.000 satırlık borç listesi, taksit formülleri, NBD karşılaştırması, senaryo/tornado, tahmin bandı, 9 grafikli pano ve karar kapısının tasarlanıp test edilmesi |
| **Toplam ikame** | **5.000** | 5.000 ≥ 2.970 ✓ |

**Y1+Y2 ≥ 3 × satış fiyatı:** 5.000 ≥ 2.970 → **GEÇTİ**

## D02 — Serbest Alternatif (Alıcı bunu bedavaya nasıl yapar?)

| Alıcının bedava yöntemi | Bu paketin ayrıştığı nokta | Dosyada karşılığı |
|---|---|---|
| Tecil taksitini ve toplam maliyeti hesap makinesiyle taksit formülüyle elle hesaplamak | Borç aslı, gecikme zammı ve vade yapısı tek tabloda toplanır; geciken ay ve aylık zam yükü formülle üretilir; alternatifte hesap elle ve dağınık kalemlerde yapılır | `TecilToplamBorc` |
| Tecil, kredi ve peşin seçeneklerini sezgisel olarak karşılaştırmak | Üç seçeneğin toplam maliyeti taksit formülüyle tek motor üzerinde üretilir; alternatifte taksit ve maliyet elle hesap makinesiyle çıkarılır | `TecilToplamMaliyet` |
| Seçenekleri bugünkü değere indirgemeden "hangisi ucuz" diye bakmak | Seçenekler iskonto oranıyla NBD'ye indirgenir ve en düşük maliyetli yol gerekçesiyle seçilir; alternatifte NBD hesabı yoktur | `NbdFarki` |
| Faiz/komisyon/terkin değişimlerinin maliyete etkisini hiç ölçmeden işlem yapmak | İyimser/baz/kötümser/kritik senaryo bandı ve tornado ile duyarlılık gösterilir; alternatifte duyarlılık analizi yoktur | `modulO3Senaryo` |
| Borç türü dağılımını ve veri kalitesini izlemeden tek seçeneğe karar vermek | Borç türü yoğunlaşması (HHI), anomali sapması ve veri kalite skoru karar kapısına güven katmanı ekler; alternatifte bu analizler yoktur | `modulO8KaliteSkor` |

## En zayıf nokta (dürüst beyan)

Tecil faizi ve terkin oranı idari düzenlemelere göre değişir; dosya güncel resmi oranı kendiliğinden getirmez, kullanıcının AYARLAR'dan güncellemesi gerekir. Kredi faizi ve komisyonu banka teklifine göre girilir; girilen oran doğrulanamaz. NBD iskonto oranı varsayımsaldır ve kullanıcı sermaye maliyetini bilmiyorsa sonuçlar yalnızca yaklaşıktır. Karar kapısı kredi onayı bilgisini kullanıcıdan alır; gerçek kredi onayı ve tecil sonucu idarenin/bankaların kararıdır. Dosya kesin mali görüş değil, karar destek aracıdır.

## Fiyat-performans gerekçesi

990 TL = bir mali müşavirin tecil/kredi karşılaştırma çalışmasının küçük bir kısmına denk gelir; karşılığında borç listesi, üç seçenekli maliyet ve NBD motoru, senaryo bandı, tornado, tahmin ve karar kapısı kalıcı olarak kullanıcının cihazında çalışır ve her yeni borç/faiz verisiyle yeniden kullanılır.
