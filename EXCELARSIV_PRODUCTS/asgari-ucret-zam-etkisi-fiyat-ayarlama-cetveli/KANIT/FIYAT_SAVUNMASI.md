# FİYAT SAVUNMASI — ASGARİ ÜCRET ZAM ETKİSİ & FİYAT AYARLAMA CETVELİ

**Satış fiyatı:** 1.490 TL
**Manda v4 D01 eşiği:** Y1 (danışman ikamesi) + Y2 (kurulum ikamesi) ≥ 3 × 1.490 = **4.470 TL**

## Y1 — Danışman ikamesi (ne yapıldığında para ödenir)

Asgari ücret zammı sonrası işçilik maliyeti etkisi ve fiyat ayarlama çalışması dış kaynaklı bir
mali müşavir veya ücret/fiyatlandırma danışmanına yaptırıldığında ödenecek bedel:

| Hizmet | İçerik | Tutar |
|---|---|---|
| Zam etkisi ve fiyat ayarlama danışmanlığı | Çalışan bazında zam öncesi/sonrası işveren maliyeti (SGK + işsizlik payları dâhil), maliyet artış oranı, ürün bazında fiyat yansıtma cetveli, senaryo/duyarlılık çalışması, yönetici raporu | 4.500 TL |
| **Y1 toplam** | | **4.500 TL** |

## Y2 — Kurulum ikamesi (iç kaynağın kurma maliyeti)

Şirket içi bir kişinin bu sistemi sıfırdan kurması (asgari ücret ve SGK/İşsizlik prim oranları,
çalışan bazında işveren maliyeti, fiyat yansıtma cetveli, karar motoru, kontrol paneli, rapor düzeni)
için gereken asgari emek bedeli:

| Kalem | Tutar |
|---|---|
| Zam etkisi ve fiyat yansıtma katmanlarının formülle kurulması ve doğrulanması | 2.500 TL |
| **Y2 toplam** | **2.500 TL** |

## Toplam değer hesabı

**Y1 + Y2 = 4.500 + 2.500 = 7.000 TL** ≥ 4.470 TL eşiği → **GEÇTİ**

## D02 serbest alternatif karşılığı (özet)

- Asgari ücret zammı çalışan bazında işveren maliyetine dönüştürülür; SGK ve işsizlik işveren
  payları otomatik eklenir. (`ToplamMaliyetArtis`)
- Zam öncesi ve sonrası işçilik maliyeti dönem karşılaştırmasıyla üretilir. (`modulT4DonemFark`)
- İşçilik maliyeti artışının ürün/hizmet fiyatına yansıtma oranı cetvel olarak üretilir. (`FiyatYansitmaOrani`)
- Kötümser/baz/iyimser zam yansıtma senaryolarında fiyat ve marj tek ekranda üretilir. (`modulO3SenaryoBazFark`)
- Tornado duyarlılıkla maliyet artışını en çok etkileyen değişken sıralanır. (`modulO2TornadoZirve`)
- Kayıtlı ile hesaplanan işçilik maliyeti farkı mutabakat denetimiyle üretilir. (`modulT5MutabakatFark`)
- Maliyet artışı tahmini ve P90 aralığı ad tanımlı ve raporlanabilir. (`modulI1MaliyetTahminOrta`)
- Maliyet artışının net bugünkü değeri ve nakit açığı köprüsü ad tanımlıdır. (`modulI3MaliyetNpv`)

## Dürüstlük notu

- Y3 (önlenen hata) bu hesaba dahil edilmemiştir.
- Asgari ücret değerleri, SGK/işsizlik prim oranları ve eşik değerler kullanıcı doğrulamalıdır;
  dosya AYARLAR'da ayrı parametreler olarak tutulur.
- Dosya mali danışmanlık görüşü yerine geçmez; karar destek amaçlıdır.
