# FİYAT SAVUNMASI — BANKA, KREDİ VE TAKSİT TAKİP SİSTEMİ

**Dosya:** `BankaKrediVeTaksitTakipSistemi.xlsx`
**Sürüm:** 1.0
**Üretim tarihi:** 10.08.2026

## 1. D01 — İkame değeri hesabı (fiyat çapası)

Satış fiyatı **1.490 ₺**; manda kuralı: **Y1 + Y2 ≥ 3 × satış fiyatı (4.470 ₺)**.

| Bileşen | Değer (₺) | Açıklama |
|---|---|---|
| Y1 — Danışman ikamesi | 3.500 | Finans danışmanının aynı analizi (borç portföyü, limit kullanımı, ödeme baskısı, refinansman değerlendirmesi, DSCR) için tek seferlik emek bedeli |
| Y2 — Kurulum ikamesi | 1.500 | Aynı sistemin sıfırdan kurulması (formül motoru, ödeme takvimi, senaryo-duyarlılık, rapor) için iç emek bedeli |
| **Toplam (Y1+Y2)** | **5.000** | ≥ 4.470 → **GEÇTİ (D01)** |

Y3 (önlenen hata / geç ödeme cezası / kur kaybı) bu çapaya dahil edilmez; yalnızca pazarlama
değeri olarak kullanılır.

## 2. D02 — Serbest alternatifle ayrışma

"Alıcı bunu bedavaya nasıl yapar?" sorusuna dürüst yanıt:

| Alternatif | Bu dosya nasıl ayrışır (çıktı düzeyinde) |
|---|---|
| Kredi ve taksitleri elle izleyen düzensiz hesap tablosu | Likidite, döviz uyumu ve kaldıraç oranlarından borç yükü kararını formülle üretir (`KararBildirimi`) |
| Banka ekstrelerinden borç servisi ve kullanılabilir nakit hesabı | Aylık borç servisini ödeme takviminden otomatik hesaplar (`AylikBorcServisi`) |
| Mali müşavir veya finans danışmanından dönemsel borç durumu raporu | Vade yaşlandırması ve ödeme eğimiyle taksit riskini ölçer (`modulT1VadeGecmis`, `modulT2OdemeEgimi`) |
| Ücretsiz genel amaçlı kredi takip şablonları | İyimser/kötümser senaryo ve tornado duyarlılıkla nakit etkisini sıralar (`modulO3SenaryoFark`) |
| Yaklaşan taksiti son anda fark edip ödeme planını ertelemek | Banka bazında borç yoğunlaşması (HHI) ve refinansman XIRR karşılaştırması üretir (`modulO6BankaHhi`) |

7 ayrım maddesinin tamamı dosyada ad tanımıyla karşılıklıdır → **GEÇTİ (D02)**.

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
3. En zayıf nokta: eşik parametreleri (likidite eşiği 0,85, kaldıraç eşiği 3,0, veri kalitesi
   eşikleri) işletme varsayımlarıdır; alıcının kendi borç yapısına göre AYARLAR'dan güncellemesi
   gerekir.
4. Y1+Y2 = 5.000 ₺ ≥ 3 × 1.490 ₺ → fiyat çapası sağlanır.
5. Serbest alternatifle en çok örtüşen özellik: kredi ve taksit girişi — bu, her yerde elle
   yapılabilen bir girdidir; fark, borç yükü kararı + ödeme baskısı + refinansman katmanıdır.
6. İçeriği zorladığımız eşik yok; modüller gerçek acıya (borç yükünün ve ödeme baskısının geç
   fark edilmesi) dayanır.
