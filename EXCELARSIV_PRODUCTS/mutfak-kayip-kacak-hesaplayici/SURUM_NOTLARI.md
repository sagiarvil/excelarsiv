# SÜRÜM NOTLARI — MUTFAK KAYIP/KAÇAK HESAPLAYICI

## 1.0.0 — İlk sürüm

### Ürün amacı
Satılan ürünlerin reçetelerinden üretilen teorik hammadde tüketimini, dönem başı stok + alım − dönem sonu stok ile bulunan fiili tüketimle karşılaştırarak mutfaktaki kayıp ve kaçağı hammadde bazında tutar ve oran olarak tespit eder; toplam kayıp maliyetini ve kâr katkısına etkisini hesaplar; kayıp yoğunlaşması, anomali, senaryo ve tornado analiziyle gelecek dönem kayıp tahmini ve KONTROL ALTINDA/İNCELE/KRİTİK KAYIP kararını gerekçesiyle üretir.

### Özellikler
- **Ürün listesi:** 1.000 satır kapasiteli tabloda ürün kodu/adı, satış adedi ve birim satış fiyatı; satış hasılatı ve reçete maliyeti canlı formülle.
- **Reçete listesi:** Ürün-hammadde eşleşmesi, miktar ve birim maliyet girişi; teorik tüketim satış adediyle çarpılarak canlı hesaplanır.
- **Stok sayfası:** Dönem başı stok, alım, dönem sonu stok ve birim maliyet girişi; fiili tüketim, teorik tüketim, kayıp/kaçak miktarı, kayıp oranı, kayıp tutarı canlı formülle.
- **Karar kapısı:** KONTROL ALTINDA / İNCELE / KRİTİK KAYIP / VERİ YOK / DURDUR kararı, gerekçesi ve önerilen aksiyonlar.
- **Analitik motor:** Ortalama kayıp, ortalama maliyet, kayıp sapması (MAD), en yüksek kayıp kalemi, senaryo bant genişliği, HHI yoğunlaşması, kalite skoru, gelecek dönem kayıp tahmini, P90 yüzdelik (9 modül, 145 puan; 2 ileri modül).
- **Panel:** 8 grafik (ürün bazlı hasılat/maliyet/kâr, kâr katkısı oranı, hammadde kayıp payı, hammadde kayıp oranı, panel trendi, projeksiyon bant aralığı, senaryo kayıp tutarları) ve 62 canlı KPI.
- **Senaryo ve duyarlılık:** İyi/Baz/Kötü/Kritik senaryo bant aralığı ve hammadde birim maliyeti/fire oranı tornado analizi.
- **Örnek veri:** Yan yana üç blok halinde örnek stok, reçete ve ürün verisi; kullanıcı ürünü örnek veriyle hemen deneyebilir.

### Denetim sonuçları
- G kapıları (G01–G24): **0 KALDI**
- Ö kapıları (Ö1, Ö1b, Ö2): **0 KALDI** (9.000 satır / 41.986 formül, 0 hata, 11,9 sn; 7 uç durum temiz)
- D kapıları (D01–D14): **0 KALDI** (D01 ikame 12.000 ₺ ≥ 7.470 ₺; D02 5 ayrım; D03 145 puan; D04 2 ileri modül)

### Altın çıktı
- Altın dosya: `KANIT/altin_cikti.json`
- İçerik SHA-256: `f90fd320fa67913aa09a674a87f958dd6d2fb928296ca637f7fde0aabda8f7f6`
- Fark sayısı: 0

### Dosya parmak izi
- SHA-256 (teslim edilen dosya): `34a282585caf230d7e4c63071aeac4b0e6e7810ddd38300b4513052703414385`
- KANIT'taki üç rapor da aynı parmak iziyle imzalanmıştır (RAPOR_DENETIM, RAPOR_OLCEK, RAPOR_DEGER).

### Değişiklik günlüğü
| Tarih | Sürüm | Değişiklik | Neden |
|---|---|---|---|
| 10.08.2026 | 1.0.0 | İlk üretim | Ürünün yayına hazırlanması |
