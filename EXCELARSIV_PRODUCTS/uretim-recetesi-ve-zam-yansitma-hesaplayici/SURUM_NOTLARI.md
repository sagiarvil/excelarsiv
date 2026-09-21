# SÜRÜM NOTLARI — ÜRETİM REÇETESİ & ZAM YANSITMA HESAPLAYICI

## 1.0.0 — İlk sürüm

### Ürün amacı
Reçetedeki hammadde kalemlerinin miktar, birim fiyat, fire ve zam oranlarıyla eski ve yeni reçete maliyetini üretir; hammadde zamlarının ürün bazında maliyet artışına dönüşümünü ve hedef kâr marjını koruyan önerilen satış fiyatını hesaplar; genel maliyet artışını eşiklerle karşılaştırarak ZAM YANSIT/KISMİ YANSIT/FİYATI KORU kararını gerekçesiyle üretir ve önerilen aksiyonları gösterir.

### Özellikler
- **Reçete listesi:** 1.000 satır kapasiteli tabloda ürün kodu/adı, hammadde, birim, miktar, birim fiyat, fire ve zam oranı girişi; hammadde tutarı, fire dahil tutar, zam sonrası birim fiyat, zam sonrası tutar, fark ve açıklama canlı formülle.
- **Maliyet motoru:** Ürün bazında eski reçete maliyeti, fire dahil eski maliyet, zam sonrası maliyet, maliyet artışı ve artış oranı; işçilik ve GÜG dahil hedef satış fiyatı (maliyet / (1 − marj)).
- **Karar kapısı:** ZAM YANSIT / KISMİ YANSIT / FİYATI KORU / VERİ YOK / DURDUR kararı, gerekçesi ve önerilen aksiyonlar.
- **Analitik motor:** Ortalama fire, ortalama zam, kalem fark sapması (MAD), en yüksek kalem farkı, senaryo bant genişliği, HHI yoğunlaşması, kalite skoru, gelecek dönem artış tahmini, P90 yüzdelik (9 modül, 145 puan; 2 ileri modül).
- **Panel:** 8 grafik (ürün bazlı eski/yeni maliyet, maliyet artışı, artış oranı, hammadde payı, hammadde karşılaştırma, panel trendi, projeksiyon bant aralığı, senaryo maliyet artışları) ve 85 canlı KPI.
- **Senaryo ve duyarlılık:** İyi/Baz/Kötü/Kritik senaryo bant aralığı ve hammadde fiyatı/fire oranı tornado analizi.

### Denetim sonuçları
- G kapıları (G01–G24): **0 KALDI**
- Ö kapıları (Ö1, Ö1b, Ö2): **0 KALDI** (3.000 satır / 17.994 formül, 0 hata, 6,2 sn; 7 uç durum temiz)
- D kapıları (D01–D14): **0 KALDI** (D01 ikame 7.500 ₺ ≥ 4.470 ₺; D02 5 ayrım; D03 145 puan; D04 2 ileri modül)

### Altın çıktı
- Altın dosya: `KANIT/altin_cikti.json`
- İçerik SHA-256: `118f1c801c53499214d65729d818677433dad0f7d9b1d3dec1eaea2a46a8bf00`
- Fark sayısı: 0

### Dosya parmak izi
- SHA-256 (teslim edilen dosya): `ae2b100f33c9f333c0b9263d21515c6567d41bc192b836ad244b024826c0f3bc`
- KANIT'taki üç rapor da aynı parmak iziyle imzalanmıştır (RAPOR_DENETIM, RAPOR_OLCEK, RAPOR_DEGER).

### Değişiklik günlüğü
| Tarih | Sürüm | Değişiklik | Neden |
|---|---|---|---|
| 10.08.2026 | 1.0.0 | İlk üretim | Ürünün yayına hazırlanması |
