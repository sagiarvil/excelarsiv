# FİYAT SAVUNMASI — AYLIK PATRON FİNANS PANELİ

**Dosya:** `AylikPatronFinansPaneli.xlsx`
**Sürüm:** 1.0.0
**Üretim tarihi:** 10.08.2026

## 1. D01 — İkame değeri hesabı (fiyat çapası)

Satış fiyatı **2.490 ₺**; manda kuralı: **Y1 + Y2 ≥ 3 × satış fiyatı (7.470 ₺)**.

| Bileşen | Değer (₺) | Açıklama |
|---|---|---|
| Y1 — Danışman ikamesi | 6.000 | Mali müşavir / finans danışmanının aynı analizi (aylık kâr-zarar + nakit + borçluluk + kasa yeterliliği raporu) için asgari 2 aylık emek bedeli |
| Y2 — Kurulum ikamesi | 2.500 | Aynı sistemin sıfırdan kurulması (tasarım, formül motoru, test, dokümantasyon) için iç emek bedeli |
| **Toplam (Y1+Y2)** | **8.500** | ≥ 7.470 → **GEÇTİ (D01)** |

Y3 (önlenen hata — örn. nakit sıkışıklığı veya farkında olunmayan borçluluk nedeniyle kaçırılan
dönem kayıpları) bu çapaya dahil edilmez; yalnızca pazarlama değeri olarak kullanılır.

## 2. D02 — Serbest alternatifle ayrışma

"Alıcı bunu bedavaya nasıl yapar?" sorusuna dürüst yanıt:

| Alternatif | Bu dosya nasıl ayrışır (çıktı düzeyinde) |
|---|---|
| Gelir-giderin elle tutulduğu dağınık hesap tablosu | Tahakkuk kârını nakit akışından ayrıştırır, dönem nakit farkını ölçer (`modulT5TahakkukNakitFark`) |
| Mali müşavirden aylık finansal özet | Kâr + borçluluk + kasa yeterliliğine göre gerekçeli UYGUN/İNCELE/DURDUR kararı üretir (`KararBildirimi`) |
| Ücretsiz genel amaçlı gelir-gider şablonları | Vadesi geçen borçları rapor tarihine göre otomatik yaşlandırır (`modulT1YasliBorc`) |
| Elle yapılan "durum iyi/kötü" yorumu | İyimser/kötümser senaryo ve tornado duyarlılıkla sonucu en çok etkileyen değişkeni sıralar (`modulO2GelirEtki`) |
| Dönem sonunu sürprizle karşılamak | Aylık nakit açığını ve kasa eşiğini önceden görür (`modulI7NakitAcigi`) |
| Her ay farklı kişiden alınan el raporu | Kalite skoruyla eksik veri girişini ölçer, rapor güvenilirliğini gösterir (`modulO8KaliteSkor`) |

Ek ayrımlar: sonraki ay gelir tahmini ve güven aralığı (`modulI1TahminUst`), gelir dağılımı
P90 yüzdelikleri (`modulI2GelirP90`). 6 ayrım maddesinin tamamı dosyada ad tanımıyla
karşılıklıdır → **GEÇTİ (D02)**.

## 3. D11 — Karar dürüstlüğü

- Boş dosyada karar **VERİ YOK** (veri varmış gibi UYGUN demez).
- Tüm kararlar gerekçe satırıyla birlikte üretilir; aksiyonlar türetilir.
- Karar kuralları KILAVUZ sayfasında açıklanır.

## 4. D10 — Kaynak şeffaflığı

Mevzuata dayanan alanlarda kaynak `kurum + madde + tarih` (VUK md.182–198); kaynağı olmayan
her eşik açıkça `Varsayım` işaretlenir ve `KANIT/PARAMETRE_KAYNAKLARI.md` içinde toplu listelenir.

## 5. Dürüstlük maddesi (manda 11)

1. Hangi kapılar geçti? → `RAPOR_DENETIM.md` (G01–G24, 0 KALDI), `RAPOR_OLCEK.md` (Ö1, Ö2; 0 KALDI), `RAPOR_DEGER.md` (D01–D14; 0 KALDI).
2. Uygulanamayan madde yok; üç katman da tamamlandı.
3. En zayıf nokta: karar eşikleri (borç/gelir oranı, kasa katı) işletmeye özgü varsayımdır;
   kurulumda kullanıcı onayı gerekir, sektörden sektöre anlamlı sonuç için eşiklerin
   gözden geçirilmesi önerilir.
4. Y1+Y2 = 8.500 ₺ ≥ 3 × 2.490 ₺ → fiyat çapası sağlanır.
5. Serbest alternatifle en çok örtüşen özellik: aylık gelir-gider kaydı — bu, her işletmede
   elle tutulabilen bir girdidir; fark, kâr/nakit ayrışması, borçluluk ve kasa yeterliliğiyle
   üretilen gerekçeli karar katmanıdır.
6. İçeriği zorladığımız eşik yok; modüller gerçek acıya (dönem sonu nakit sürprizi, farkında
   olunmayan borçluluk) dayanır.
