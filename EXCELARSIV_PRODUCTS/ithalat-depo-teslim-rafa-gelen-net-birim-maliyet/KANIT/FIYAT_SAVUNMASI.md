# FİYAT SAVUNMASI — İTHALAT DEPO TESLİM (RAFA GELEN) NET BİRİM MALİYET

**Satış fiyatı:** 1.490 TL
**Manda v4 D01 eşiği:** Y1 (danışman ikamesi) + Y2 (kurulum ikamesi) ≥ 3 × 1.490 = **4.470 TL**

## Y1 — Danışman ikamesi (ne yapıldığında para ödenir)

Depo teslim birim maliyet hesaplaması dış kaynaklı bir dış ticaret/gümrük danışmanına
yaptırıldığında ödenecek bedel:

| Hizmet | İçerik | Tutar |
|---|---|---|
| İthalat maliyeti danışmanlığı | FOB→navlun/sigorta→CIF→gümrük vergisi→ÖTV→KDV matrah zinciri, yurt içi masraf dağıtımı, vergi yükü ve kur senaryo/duyarlılık çalışması, yönetici raporu | 4.000 TL |
| **Y1 toplam** | | **4.000 TL** |

## Y2 — Kurulum ikamesi (iç kaynağın kurma maliyeti)

Şirket içi bir kişinin bu sistemi sıfırdan kurması (4458 s. Gümrük Kanunu kıymet mantığı,
3065 s. KDV md.21 matrah zinciri, 4760 s. ÖTV matrahı, FOB payıyla masraf dağıtımı, kontrol
paneli, rapor düzeni) için gereken asgari emek bedeli:

| Kalem | Tutar |
|---|---|
| Vergi matrah zinciri ve masraf dağıtım katmanlarının formülle kurulması ve doğrulanması | 2.500 TL |
| **Y2 toplam** | **2.500 TL** |

## Toplam değer hesabı

**Y1 + Y2 = 4.000 + 2.500 = 6.500 TL** ≥ 4.470 TL eşiği → **GEÇTİ**

## D02 serbest alternatif karşılığı (özet)

- FOB kıymetten depo teslim maliyete giden adım zinciri (FOB→navlun→CIF→GV→ÖTV→KDV→masraf)
  ürün satırında tek formülle üretilir. (`ToplamDepoTeslimMaliyeti`)
- Yurt içi operasyon masrafları ürünlerin FOB payıyla orantılı dağıtılır; masraf tablosu ayrıdır. (`ToplamYurtIciMasraf`)
- Bileşik vergi yükü (GV+ÖTV+KDV matrah zinciri) ürün bazında raporlanır. (`modulT6VergiYukOran`)
- Kötümser/baz/iyimser kur ve oran senaryolarında depo teslim maliyeti ve karar değişimi üretilir. (`modulO3SenaryoFark`)
- Tornado duyarlılıkla depo teslim maliyetini en çok etkileyen değişken sıralanır. (`modulO2TornadoZirve`)
- Hesaplanan toplam maliyet ile giriş toplamları arasındaki mutabakat farkı denetlenir. (`modulT5MutabakatFark`)
- Birim maliyet tahmini, P90 aralığı ve NPV değeri ad tanımlı ve raporlanabilir. (`modulI1MaliyetTahmin`)
- Peşin ödenen ithal vergilerinin nakit açığı köprüsü ad tanımlıdır. (`modulI7NakitAcik`)

## Dürüstlük notu

- Y3 (önlenen hata) bu hesaba dahil edilmemiştir.
- Vergi oranları ve kurlar kullanıcı tarafından doğrulanmalıdır; dosya AYARLAR'da ayrı
  parametreler olarak tutulur.
- Dosya gümrük müşavirliği veya mali danışmanlık görüşü yerine geçmez; karar destek amaçlıdır.
