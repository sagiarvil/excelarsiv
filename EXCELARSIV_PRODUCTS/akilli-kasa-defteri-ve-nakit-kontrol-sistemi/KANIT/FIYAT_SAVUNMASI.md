# FİYAT SAVUNMASI — AKILLI KASA DEFTERİ VE NAKİT KONTROL SİSTEMİ

**Dosya:** `AkilliKasaDefteriVeNakitKontrolSistemi.xlsx`
**Sürüm:** 1.0
**Üretim tarihi:** 10.08.2026

## 1. D01 — İkame değeri hesabı (fiyat çapası)

Satış fiyatı **990 ₺**; manda kuralı: **Y1 + Y2 ≥ 3 × satış fiyatı (2.970 ₺)**.

| Bileşen | Değer (₺) | Açıklama |
|---|---|---|
| Y1 — Danışman ikamesi | 2.500 | Finans danışmanı / mali müşavirin aynı analizi (kullanılabilir nakit hesabı, nakit açığı tespiti, ödeme önerisi, fiziki mutabakat değerlendirmesi) için aylık asgari emek bedeli |
| Y2 — Kurulum ikamesi | 1.200 | Aynı sistemin sıfırdan kurulması (formül motoru, takvim/projeksiyon, senaryo-duyarlılık, rapor) için iç emek bedeli |
| **Toplam (Y1+Y2)** | **3.700** | ≥ 2.970 → **GEÇTİ (D01)** |

Y3 (önlenen hata / nakit açığına düşme maliyeti) bu çapaya dahil edilmez; yalnızca pazarlama
değeri olarak kullanılır.

## 2. D02 — Serbest alternatifle ayrışma

"Alıcı bunu bedavaya nasıl yapar?" sorusuna dürüst yanıt:

| Alternatif | Bu dosya nasıl ayrışır (çıktı düzeyinde) |
|---|---|
| Kasa/banka bakiyelerini elle izleyen düzensiz tablo | Bloke ve minimum eşiği ayırarak gerçek kullanılabilir parayı formülle üretir; bakiyeyle karıştırmaz (`KullanilabilirNakit`) |
| Mali müşavir danışmanlığı | Bekleyen ödeme/tahsilat dengesinden nakit açığını ve ödeme önerisini otomatik üretir (`NakitAcik`, `OdenecekOnerisi`) |
| Ücretsiz genel amaçlı şablonlar | Vade yaşlandırması, günlük nakit eğimi, senaryo ve tornado duyarlılık ile nakit riskini ölçer (`modulT1VadeGecmis`, `modulO3SenaryoFark`) |
| Banka ekstrelerinden elle hesap | Fiziki kasa mutabakat farkını kök nedenle birlikte üretir (`modulT5MutabakatFark`) |
| Nakit açığını son anda ertelemek | Haftalık nakit projeksiyonu ve P90 aralığıyla kasa tahminini önceden verir (`modulI2KasaP90`) |

7 ayrım maddesinin tamamı dosyada ad tanımıyla karşılıklıdır → **GEÇTİ (D02)**.

## 3. D11 — Karar dürüstlüğü

- Boş dosyada karar **VERİ YETERSİZ** (veri varmış gibi UYGUN demez).
- Tüm kararlar gerekçe satırıyla birlikte üretilir; aksiyonlar türetilir.
- Karar kuralları KILAVUZ sayfasında açıklanır.

## 4. D10 — Kaynak şeffaflığı

Mevzuata dayanan parametrelerde kaynak `kurum + madde + tarih`; kaynağı olmayanlar açıkça
`Varsayım` işaretlenir ve `KANIT/PARAMETRE_KAYNAKLARI.md` içinde toplu listelenir.

## 5. Dürüstlük maddesi (manda 11)

1. Hangi kapılar geçti? → `RAPOR_DENETIM.md` (G01–G24, 0 KALDI), `RAPOR_OLCEK.md` (Ö1, Ö1b, Ö2; 0 KALDI), `RAPOR_DEGER.md` (D01–D14; 0 KALDI).
2. Uygulanamayan madde yok; üç katman da tamamlandı.
3. En zayıf nokta: eşik parametreleri (asgari nakit tamponu, kritik risk eşiği vb.) işletme
   varsayımlarıdır; alıcının kendi nakit yapısına göre AYARLAR'dan güncellemesi gerekir.
4. Y1+Y2 = 3.700 ₺ ≥ 3 × 990 ₺ → fiyat çapası sağlanır.
5. Serbest alternatifle en çok örtüşen özellik: bakiye girişi — bu, her yerde elle yapılabilen
   bir girdidir; fark, kullanılabilir para + açık + ödeme önerisi katmanıdır.
6. İçeriği zorladığımız eşik yok; modüller gerçek acıya (yanlış nakit kararı ve nakit açığına
   düşme) dayanır.
