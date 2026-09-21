# SÜRÜM NOTLARI — NAKLİYE MALİYETİ HESAPLAYICI

## 1.0.0 — İlk sürüm

### Ürün amacı
Filo bazında sefer maliyetlerini yakıt, bakım, sürücü ve amortisman kalemlerine ayırarak hesaplar; hasılatla karşılaştırarak sefer kârlılığını ve kâr marjını çıkarır; birim maliyeti km ve ton-km bazında üretir; senaryo ve tornado analiziyle gelecek dönem maliyet tahmini ve KARLI/İNCELE/ZARARLI kararını gerekçesiyle üretir.

### Özellikler
- **Araç listesi:** 1.000 satır kapasiteli tabloda plaka, araç tipi, yakıt tüketimi (lt/100 km), bakım oranı, sürücü ücreti, amortisman ve tonaj girişi.
- **Sefer kayıtları:** 1.000 satır kapasiteli tabloda sefer numarası, araç, güzergâh, mesafe, tonaj, hasılat ve yakıt tüketimi düzeltmesi girişi; yakıt, bakım, sürücü ve amortisman maliyetleri canlı formülle.
- **Karar kapısı:** KARLI / İNCELE / ZARARLI / VERİ YOK / DURDUR kararı, gerekçesi ve önerilen aksiyonlar.
- **Analitik motor:** Ortalama sefer maliyeti, ortalama kâr marjı, maliyet sapması (MAD), en yüksek maliyet kalemi, senaryo bandı, HHI hasılat yoğunlaşması, kalite skoru, gelecek dönem maliyet tahmini, P90 yüzdelik (9 modül, 145 puan; 2 ileri modül).
- **Panel:** 8 grafik (araç bazlı maliyet/hasılat/kâr, birim maliyet, maliyet bileşenleri, hasılat yoğunlaşması, panel trendi, projeksiyon bandı, senaryo net kârları) ve 54 canlı KPI.
- **Senaryo ve duyarlılık:** İyi/Baz/Kötü/Kritik senaryo yakıt fiyatı bandı ve yakıt fiyatı/bakım oranı tornado analizi.
- **Örnek veri:** Araç ve sefer örnek blokları; kullanıcı ürünü örnek veriyle hemen deneyebilir.

### Denetim sonuçları
- G kapıları (G01–G24): **0 KALDI** (G15 taşma uyarısı: 2 hücre, içerik 19 karakter/kolon 8 — anlam kaybı yok)
- Ö kapıları (Ö1, Ö1b, Ö2): **0 KALDI** (6.000 satır / 31.489 formül, 0 hata, 9,8 sn; 7 uç durum temiz)
- D kapıları (D01–D14): **0 KALDI** (D01 ikame 7.500 ₺ ≥ 2.970 ₺; D02 5 ayrım; D03 145 puan; D04 2 ileri modül)

### Altın çıktı
- Altın dosya: `KANIT/altin_cikti_nakliye.json`
- İçerik SHA-256: `e9757b4471392da4c84aa817cec0d1dd0798bfde594edcac4c5bfddef5e57be3`
- Fark sayısı: 0

### Dosya parmak izi
- SHA-256 (teslim edilen dosya): `f8ae7a4a8f3e9adec1695027975680780ca29380c6754691dd6fd835bd895ae5`
- KANIT'taki üç rapor da aynı parmak iziyle imzalanmıştır (RAPOR_DENETIM, RAPOR_OLCEK, RAPOR_DEGER).

### Değişiklik günlüğü
| Tarih | Sürüm | Değişiklik | Neden |
|---|---|---|---|
| 10.08.2026 | 1.0.0 | İlk üretim | Ürünün yayına hazırlanması |
