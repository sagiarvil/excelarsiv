# FİYAT SAVUNMASI — GÜNLÜK GELİR-GİDER VE GERÇEK KÂRLILIK SİSTEMİ

**Satış fiyatı:** 1.490 TL
**Manda v4 D01 eşiği:** Y1 (danışman ikamesi) + Y2 (kurulum ikamesi) ≥ 3 × 1.490 = **4.470 TL**

## Y1 — Danışman ikamesi (ne yapıldığında para ödenir)

Kârlılık ve nakit danışmanlığı kapsamında şu hizmetler dışarıdan alındığında ödenecek bedel:

| Hizmet | İçerik | Tutar |
|---|---|---|
| Finans/kârlılık danışmanlığı | Tahakkuk-nakit köprüsü, brüt/faaliyet marjı analizi, başabaş cirosu, nakit dönüşümü, senaryo/duyarlılık ve yönetici raporu hazırlanması | 6.000 TL |
| **Y1 toplam** | | **6.000 TL** |

## Y2 — Kurulum ikamesi (iç kaynağın kurma maliyeti)

Şirket içi bir kişinin bu sistemi sıfırdan kurması (hesap planı, kayıt defteri, tahakkuk/nakit eşleme motoru, aylık kârlılık ve günlük özet, karar kuralları, analitik modüller, pano ve rapor düzeni) için gereken asgari emek bedeli:

| Kalem | Tutar |
|---|---|
| Kayıt defteri + tahakkuk/nakit köprüsü ve günlük/aylık özet motorunun formülle kurulması ve doğrulanması | 1.700 TL |
| Pano, başabaş, senaryo modülleri ve rapor katmanlarının kurulması | 800 TL |
| **Y2 toplam** | **2.500 TL** |

## Toplam değer hesabı

**Y1 + Y2 = 6.000 + 2.500 = 8.500 TL** ≥ 4.470 TL eşiği → **GEÇTİ**

## D02 — Serbest alternatifle örtüşme riski

Alıcının "bunu bedavaya kendim yaparım" diyebileceği en güçlü nokta: **gelir ve giderleri deftere ya da tek Excel listesine elle işleyip dönem sonunda toplam almak**. Sistemin bu alternatife karşı ayrıştığı maddeler dosyada ad tanımlarıyla kanıtlanır: tahakkuk ile nakit arasındaki köprü (`NakitDonusumOrani`), başabaş cirosu (`BasabasCiro`), nakdi gerçekleşmemiş satış yaşlandırması (`modulT1Yaslandirma`), aylık trend (`modulT2Trend`), satış hacmi ±%15 kâr bandı (`modulO3Senaryo`) ve veri yetersizken "VERİ YOK" diyen karar katmanı (`KayitSayisi`). Elle toplam tek başına "kâr var ama para yok" sorusuna cevap vermez; bu dosya verir.
