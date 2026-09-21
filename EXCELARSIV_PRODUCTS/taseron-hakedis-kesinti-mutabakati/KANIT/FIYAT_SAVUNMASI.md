# FİYAT SAVUNMASI — TAŞERON/ALT YÜKLENİCİ HAKEDİŞ & KESİNTİ MUTABAKATI

**Satış fiyatı:** 1.490 TL
**Manda v4 D01 eşiği:** Y1 (danışman ikamesi) + Y2 (kurulum ikamesi) ≥ 3 × 1.490 = **4.470 TL**

## Y1 — Danışman ikamesi (ne yapıldığında para ödenir)

Taşeron hakediş mutabakatı dış kaynaklı bir mali müşavir/danışmana yaptırıldığında ödenecek
tek seferlik mutabakat ve denetim bedeli:

| Hizmet | İçerik | Tutar |
|---|---|---|
| Taşeron mutabakat ve kesinti denetimi | Her hakedişin KDV tevkifatı, stopaj, SGK ve teminat kesintilerinin hesabı; hakediş-ödeme mutabakat çıkarımı, geç ödeme/gecikme maliyeti tespiti, uç dönem analizi ve yönetici raporu | 4.500 TL |
| **Y1 toplam** | | **4.500 TL** |

## Y2 — Kurulum ikamesi (iç kaynağın kurma maliyeti)

Şirket içi bir kişinin bu mutabakat sistemini sıfırdan kurması (kesinti formül mimarisi,
mutabakat motoru, senaryo/duyarlılık, kontrol paneli, rapor düzeni) için gereken asgari emek
bedeli:

| Kalem | Tutar |
|---|---|
| Kesinti, mutabakat, gecikme ve karar katmanlarının formülle kurulması | 2.500 TL |
| **Y2 toplam** | **2.500 TL** |

## Toplam değer hesabı

**Y1 + Y2 = 4.500 + 2.500 = 7.000 TL** ≥ 4.470 TL eşiği → **GEÇTİ**

## Y3 — Önlenen hata (bilgi amaçlı, D01 hesabına dahil değil)

Yanlış kesinti uygulaması veya gözden kaçan mutabakat farkının ceza ve finansman maliyeti
hesaplanmamıştır; D01 hesabı yalnızca Y1 ve Y2 üzerinden kurulur.

## Serbest alternatif ayrımı (D02 özeti)

Aşağıdaki her madde, dosyada ad tanımlı bir hücreye karşılık gelir ve "bunu bedavaya kendim
yaparım" diyen alıcıya karşı savunulabilir:

1. **Otomatik kesinti hesabı** — Her hakedişin KDV tevkifatı, gelir vergisi stopajı, SGK
   kesintisi ve teminatı sözleşme kartındaki oranlardan formülle hesaplanır; elle oran kurmayı
   gerektirmez. (`ToplamKdvTevkifat`)
2. **Mutabakat farkı** — Hakediş net tutarı ile ödenen tutarı tek tabloda karşılaştırır;
   farkı ve fark oranını formülle üretir. (`MutabakatFarki`)
3. **Gecikme maliyeti** — Geç ödeme gününü vadeyle karşılaştırır, gecikme maliyetini gecikme
   faizi oranından türetir. (`GecikmeMaliyetiToplam`)
4. **Anomali denetimi** — Hakediş anomali oranı Z skoruyla denetlenir; uç dönemler mutabakat
   gerekliliği olarak raporlanır. (`modulO1AnomaliOran`)
5. **Tornado duyarlılık** — Mutabakat farkını en çok etkileyen değişken sıralanır. (`modulO2TornadoZirve`)
6. **Senaryo motoru** — Kötümser/baz/iyimser fark, gecikme ve karar tek ekranda üretilir. (`modulO3SenaryoBazFark`)
7. **Tahmin aralığı** — Hakediş trendinden sonraki dönem tahmini ve P90 aralığı üretilir; sonuç
   ad tanımlı ve raporlanabilir. (`modulI1HakedisTahminOrta`)

## Sonuç

- D01: **GEÇTİ** (7.000 ≥ 4.470)
- D02: **GEÇTİ** (7 ayrım maddesi ≥ 5, her biri ad tanımlı çıktı)
- Satış fiyatı 1.490 TL; ürün üzerinde kurulan değer mimarisi **ikame bedelinin 4,7 katı**
  değer üretir (7.000 TL ikame / 1.490 TL satış).
