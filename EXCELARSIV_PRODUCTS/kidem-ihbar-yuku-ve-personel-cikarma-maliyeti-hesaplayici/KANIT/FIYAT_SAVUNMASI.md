# FİYAT SAVUNMASI — KIDEM–İHBAR YÜKÜ VE PERSONEL ÇIKARMA MALİYETİ HESAPLAYICI

**Satış fiyatı:** 1.490 TL
**Manda v4 D01 eşiği:** Y1 (danışman ikamesi) + Y2 (kurulum ikamesi) ≥ 3 × 1.490 = **4.470 TL**

## Y1 — Danışman ikamesi (ne yapıldığında para ödenir)

Kıdem/ihbar tazminatı ve personel çıkarma maliyeti hesabı dış kaynaklı bir İK danışmanı veya
mali müşavire yaptırıldığında ödenecek bedel:

| Hizmet | İçerik | Tutar |
|---|---|---|
| Toplu çıkarma maliyeti danışmanlığı | Çalışan bazında kıdem, ihbar, kullanılmamış izin hesapları; bütçe karşılaştırması, uç değer/anomali denetimi, senaryo ve duyarlılık analizi, yönetici raporu | 4.500 TL |
| **Y1 toplam** | | **4.500 TL** |

## Y2 — Kurulum ikamesi (iç kaynağın kurma maliyeti)

Şirket içi bir kişinin bu sistemi sıfırdan kurması (tavan sınırlı kıdem hesabı, ihbar kademe
mantığı, izin hakkı tablosu, bütçe karşılaştırma, senaryo/duyarlılık, kontrol paneli, rapor düzeni)
için gereken asgari emek bedeli:

| Kalem | Tutar |
|---|---|
| Kıdem/ihbar/izin katmanlarının formülle kurulması ve doğrulanması | 2.500 TL |
| **Y2 toplam** | **2.500 TL** |

## Toplam değer hesabı

**Y1 + Y2 = 4.500 + 2.500 = 7.000 TL** ≥ 4.470 TL eşiği → **GEÇTİ**

## D02 serbest alternatif karşılığı (özet)

- Kıdem hesabı 1475 s. md.14 kuralıyla (tam yıl + artan gün oranı, tavan sınırlı) formülden üretilir; elle oran kurmaya gerek yoktur. (`ToplamKidem`)
- İhbar süresi 4857 s. md.17 kıdem dilimlerine göre otomatik seçilir. (`ToplamIhbar`)
- Kullanılmamış izin ücreti 4857 s. md.53 haklarından hesaplanır. (`ToplamIzinUcreti`)
- Toplam maliyet bütçe eşiğiyle karşılaştırılır; kullanım oranı ve açık köprüsü üretilir. (`ButceKullanimOrani`)
- Ücret anomalileri Z skoruyla denetlenir. (`modulO1AnomaliOran`)
- Tornado duyarlılıkla maliyeti en çok etkileyen değişken sıralanır. (`modulO2TornadoZirve`)
- Kötümser/baz/iyimser senaryo kararları tek ekranda üretilir. (`modulO3SenaryoBazFark`)
- Trend tahmini ve P90 aralığı ad tanımlı ve raporlanabilir. (`modulI1TahminOrta`)

## Dürüstlük notu

- Y3 (önlenen hata) bu hesaba dahil edilmemiştir.
- Kıdem tavanı güncel değeri kullanıcı doğrulamalıdır; dosya yıllık duyurudan güncellenmek üzere
  AYARLAR'da ayrı bir parametre olarak tutulur.
