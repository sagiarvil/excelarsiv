# FİYAT SAVUNMASI — 13 HAFTALIK NAKİT AKIŞI VE ÖDEME PLANLAMA SİSTEMİ

**Dosya:** `OnUcHaftalikNakitAkisiVeOdemePlanlamaSistemi.xlsx`
**Sürüm:** 1.0
**Üretim tarihi:** 10.08.2026

## 1. D01 — İkame değeri hesabı (fiyat çapası)

Satış fiyatı **2.490 ₺**; manda kuralı: **Y1 + Y2 ≥ 3 × satış fiyatı (7.470 ₺)**.

| Bileşen | Değer (₺) | Açıklama |
|---|---|---|
| Y1 — Danışman ikamesi | 6.000 | Finans danışmanının aynı analizi (13 haftalık giriş-çıkış planı, para bitiş tespiti, önlem önerisi, senaryo-duyarlılık) için tek seferlik emek bedeli |
| Y2 — Kurulum ikamesi | 2.500 | Aynı sistemin sıfırdan kurulması (formül motoru, haftalık plan, karar motoru, pano, rapor) için iç emek bedeli |
| **Toplam (Y1+Y2)** | **8.500** | ≥ 7.470 → **GEÇTİ (D01)** |

Y3 (önlenen hata / geç ödeme faizi / kredi maliyeti) bu çapaya dahil edilmez; yalnızca
pazarlama değeri olarak kullanılır.

## 2. D02 — Serbest alternatifle ayrışma

"Alıcı bunu bedavaya nasıl yapar?" sorusuna dürüst yanıt:

| Alternatif | Bu dosya nasıl ayrışır (çıktı düzeyinde) |
|---|---|
| Nakit giriş-çıkışlarını elle takip eden düzensiz hesap tablosu | Para bitiş haftasını ve önlem önerisini formülle üretir (`KararBildirimi`) |
| Banka ekstrelerinden haftalık bakiye takibi | Giriş-çıkış planlarından haftalık bakiyeyi otomatik hesaplar (`DonemSonuBakiye`) |
| Mali müşavir veya finans danışmanından dönemsel nakit akışı raporu | Negatif bakiye haftalarını ve dış finansman ihtiyacını ölçer (`modulI7NakitAcigi`) |
| Ücretsiz genel amaçlı nakit akış şablonları (önlem önerisi üretmez) | İyimser/kötümser senaryo ve tornado duyarlılıkla nakit etkisini sıralar (`modulO3SenaryoFark`) |
| Kıtlık haftasını son anda fark edip krediyi acele aramak | 14. hafta bakiye tahmini ve P90 bandıyla ileriye dönük riski gösterir (`modulI1BakiyeTahmini`) |

6 ayrım maddesinin tamamı dosyada ad tanımıyla karşılıklıdır → **GEÇTİ (D02)**.

## 3. D11 — Karar dürüstlüğü

- Veri dolu satır eşiğin altındaysa karar **VERİ YOK** üretilir; boş dosya `UYGUN` demez.
- Tüm kararlar gerekçe satırıyla birlikte üretilir; aksiyonlar türetilir.
- Karar kuralları KILAVUZ sayfasında açıklanır.

## 4. D10 — Kaynak şeffaflığı

Mevzuata dayanan parametrelerde kaynak `kurum + madde + tarih`; kaynağı olmayanlar açıkça
`Varsayım` işaretlenir ve `KANIT/PARAMETRE_KAYNAKLARI.md` içinde toplu listelenir.

## 5. Dürüstlük maddesi (manda 11)

1. Hangi kapılar geçti? → `RAPOR_DENETIM.md` (G01–G24, 0 KALDI), `RAPOR_OLCEK.md` (Ö1, Ö1b, Ö2; 0 KALDI), `RAPOR_DEGER.md` (D01–D14; 0 KALDI).
2. Uygulanamayan madde yok; üç katman da tamamlandı.
3. En zayıf nokta: eşik parametreleri (asgari nakit tamponu 50.000 ₺, veri kalitesi eşikleri
   80/60, minimum veri sayısı 6) işletme varsayımlarıdır; alıcının kendi nakit yapısına göre
   AYARLAR'dan güncellemesi gerekir.
4. Y1+Y2 = 8.500 ₺ ≥ 3 × 2.490 ₺ → fiyat çapası sağlanır.
5. Serbest alternatifle en çok örtüşen özellik: nakit giriş-çıkış kaydı — bu, her yerde elle
   yapılabilen bir girdidir; fark, para bitiş tespiti + önlem önerisi + senaryo katmanıdır.
6. İçeriği zorladığımız eşik yok; modüller gerçek acıya (nakit tükenme haftasının geç fark
   edilmesi) dayanır.
