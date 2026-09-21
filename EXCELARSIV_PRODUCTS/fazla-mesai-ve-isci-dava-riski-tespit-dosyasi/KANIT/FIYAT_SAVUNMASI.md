# FİYAT SAVUNMASI — FAZLA MESAİ VE İŞÇİ DAVA RİSKİ TESPİT DOSYASI

**Satış fiyatı:** 2.490 TL
**Manda v4 D01 eşiği:** Y1 (danışman ikamesi) + Y2 (kurulum ikamesi) ≥ 3 × 2.490 = **7.470 TL**

## Y1 — Danışman ikamesi (ne yapıldığında para ödenir)

Fazla mesai yükü ve işçi dava riski değerlendirmesi dış kaynaklı bir iş hukuku avukatına veya
İK/ücret danışmanına yaptırıldığında ödenecek bedel:

| Hizmet | İçerik | Tutar |
|---|---|---|
| Fazla mesai ve dava riski danışmanlığı | Çalışan bazında fazla mesai ücreti (%50 zamlı, 270 saat sınırı), bildirilen/belgelenen kayıt farkı denetimi, zamanaşımı penceresi riski, anomali ve yoğunlaşma analizi, senaryo/duyarlılık çalışması, yönetici raporu | 7.500 TL |
| **Y1 toplam** | | **7.500 TL** |

## Y2 — Kurulum ikamesi (iç kaynağın kurma maliyeti)

Şirket içi bir kişinin bu sistemi sıfırdan kurması (4857 s. md.41/42/43 fazla çalışma mantığı,
270 saat sınırı denetimi, kayıt mutabakatı, zamanaşımı risk modeli, kontrol paneli, rapor düzeni)
için gereken asgari emek bedeli:

| Kalem | Tutar |
|---|---|
| Mesai ücreti ve dava riski katmanlarının formülle kurulması ve doğrulanması | 4.000 TL |
| **Y2 toplam** | **4.000 TL** |

## Toplam değer hesabı

**Y1 + Y2 = 7.500 + 4.000 = 11.500 TL** ≥ 7.470 TL eşiği → **GEÇTİ**

## D02 serbest alternatif karşılığı (özet)

- Fazla mesai ücreti 4857 s. md.41 kuralıyla (%50 zamlı, 270 saat sınırlı) formülden hesaplanır;
  elle oran kurmaya gerek yoktur. (`ToplamFazlaMesaiUcreti`)
- Bildirilen ile belgelenen mesai saati karşılaştırılır; kayıt eksikliği oranı üretilir. (`KayitFarkiOrani`)
- Yıllık 270 saat sınırı aşımı çalışan bazında denetlenir. (`SinirAsimCalisan`)
- 5 yıllık zamanaşımı penceresindeki risk tutarı hesaplanır. (`modulO4ZamanAsimiRisk`)
- Tornado duyarlılıkla riski en çok etkileyen değişken sıralanır. (`modulO2TornadoZirve`)
- Kötümser/baz/iyimser senaryo kararları tek ekranda üretilir. (`modulO3SenaryoBazFark`)
- Mesai trendinden sonraki dönem tahmini ve P90 aralığı ad tanımlı ve raporlanabilir. (`modulI1MesaiTahminOrta`)
- Risk tutarının net bugünkü değeri ve açık köprüsü ad tanımlıdır. (`modulI3RiskNpv`)

## Dürüstlük notu

- Y3 (önlenen hata) bu hesaba dahil edilmemiştir.
- Zamanaşımı süresi (5 yıl) ve sınır değerler kullanıcı doğrulamalıdır; dosya AYARLAR'da ayrı
  parametreler olarak tutulur.
- Dosya hukuki görüş yerine geçmez; karar destek amaçlıdır.
