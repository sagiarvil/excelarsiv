# FİYAT SAVUNMASI — YILLARA SARİ İNŞAAT STOPAJ & NAKİT AKIŞ PLANLAYICI

**Satış fiyatı:** 2.490 TL
**Manda v4 D01 eşiği:** Y1 (danışman ikamesi) + Y2 (kurulum ikamesi) ≥ 3 × 2.490 = **7.470 TL**

## Y1 — Danışman ikamesi (ne yapıldığında para ödenir)

Yıllara sari bir inşaat işinin stopaj, tevkifat, vergi planı ve nakit akışı dış kaynaklı bir mali
müşavir/danışmana yaptırıldığında ödenecek tek seferlik danışman bedeli:

| Hizmet | İçerik | Tutar |
|---|---|---|
| Yıllara sari iş vergi ve nakit planı | KDV tevkifatı + GV stopajı hesabı, dönemlik kümülatif nakit planı, geçici/kurumlar vergisi öngörüsü, yönetici raporu | 7.500 TL |
| **Y1 toplam** | | **7.500 TL** |

## Y2 — Kurulum ikamesi (iç kaynağın kurma maliyeti)

Şirket içi bir kişinin bu planlayıcıyı sıfırdan kurması (formül mimarisi, senaryo/duyarlılık,
kontrol paneli, rapor düzeni) için gereken asgari emek bedeli:

| Kalem | Tutar |
|---|---|
| Kesinti, vergi, nakit ve karar katmanlarının formülle kurulması | 4.000 TL |
| **Y2 toplam** | **4.000 TL** |

## Toplam değer hesabı

**Y1 + Y2 = 7.500 + 4.000 = 11.500 TL** ≥ 7.470 TL eşiği → **GEÇTİ**

## Y3 — Önlenen hata (bilgi amaçlı, D01 hesabına dahil değil)

Yanlış stopaj/tevkifat uygulaması veya gözden kaçan nakit açığının ceza ve finansman maliyeti
hesaplanmamıştır; D01 hesabı yalnızca Y1 ve Y2 üzerinden kurulur.

## Serbest alternatif ayrımı (D02 özeti)

Aşağıdaki her madde, dosyada ad tanımlı bir hücreye karşılık gelir ve "bunu bedavaya kendim
yaparım" diyen alıcıya karşı savunulabilir:

1. **Otomatik kesinti hesabı** — Her hakedişin KDV tevkifatı ve gelir vergisi stopajı sözleşme
   kartındaki oranlardan formülle hesaplanır; elle oran kurmayı gerektirmez. (`ToplamKdvTevkifat`)
2. **Kümülatif nakit takibi** — Dönem dönem net nakit akışı ve kümülatif pozisyon formülle
   işlenir; açık dönemleri ve açık tutarını gösterir. (`NakitAcik`)
3. **Vergi planı ve stopaj mahsubu** — Yıllık matrah, geçici/kurumlar vergisi ve stopaj mahsubu
   üretilir; vergi yükü kâra oranlanır. (`ToplamOdenekVergi`)
4. **Anomali denetimi** — Hakediş anomali oranı Z skoruyla denetlenir; uç dönemler mutabakat
   gerekliliği olarak raporlanır. (`modulO1AnomaliOran`)
5. **Tornado duyarlılık** — Kâr oranını en çok etkileyen değişken sıralanır. (`modulO2TornadoZirve`)
6. **Senaryo motoru** — Kötümser/baz/iyimser kâr, DSCR ve karar tek ekranda üretilir. (`modulO3SenaryoBazKar`)
7. **Tahmin aralığı** — Hakediş trendinden sonraki dönem tahmini ve P90 aralığı üretilir; sonuç
   ad tanımlı ve raporlanabilir. (`modulI1HakedisTahminOrta`)

## Sonuç

- D01: **GEÇTİ** (11.500 ≥ 7.470)
- D02: **GEÇTİ** (7 ayrım maddesi ≥ 5, her biri ad tanımlı çıktı)
- Satış fiyatı 2.490 TL; ürün üzerinde kurulan değer mimarisi **ikame bedelinin 4,6 katı**
  değer üretir (11.500 TL ikame / 2.490 TL satış).
