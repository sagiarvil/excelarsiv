# FİYAT SAVUNMASI — AŞIRI DÜŞÜK TEKLİF SAVUNMA ROBOTU

**Dosya:** `AsiriDusukTeklifSavunmaRobotu.xlsx`
**Sürüm:** 1.0.0
**Üretim tarihi:** 09.08.2026

## 1. D01 — İkame değeri hesabı (fiyat çapası)

Satış fiyatı **7.900 ₺**; manda kuralı: **Y1 + Y2 ≥ 3 × satış fiyatı (23.700 ₺)**.

| Bileşen | Değer (₺) | Açıklama |
|---|---|---|
| Y1 — Danışman ikamesi | 22.000 | İhale mevzuat danışmanının tek ihaleye özel aşırı düşük savunma dosyası (sınır değer hesabı, bileşen analizi, dayanak raporu) için emek bedeli |
| Y2 — Kurulum ikamesi | 6.000 | Aynı sistemin sıfırdan kurulması (tasarım, formül motoru, test, dokümantasyon) için iç emek bedeli |
| **Toplam (Y1+Y2)** | **28.000** | ≥ 23.700 → **GEÇTİ (D01)** |

Y3 (önlenen hata / ihale kaybına dönüşmüş para) bu çapaya dahil edilmez; yalnızca pazarlama
değeri olarak kullanılır.

## 2. D02 — Serbest alternatifle ayrışma

"Alıcı bunu bedavaya nasıl yapar?" sorusuna dürüst yanıt:

| Alternatif | Bu dosya nasıl ayrışır (çıktı düzeyinde) |
|---|---|
| İdarelerin açık yayımladığı sınır değer cetvelleri | Sınır değeri N katsayısı ve yöntem seçimiyle otomatik hesaplar; yöntem değişiminde cetveli elle güncelleme gerektirmez (`SdSinirDeger`) |
| İhale mevzuat danışmanının teklif başı danışmanlığı | Teklifi bölgeye (AŞIRI DÜŞÜK/SINIRDA/GÜVENLİ) otomatik yerleştirir ve güvenlik payını TL/oran bazında gösterir (`SdKontrol`) |
| Ücretsiz Excel şablonları | Bileşen birim fiyatlarının rayiçten sapmasını kalem kalem ölçer; risk bayrağı ve ağırlık üretir (`modulT6KirilimOrani`) |
| Genel amaçlı yapay zekâ ile elle kurulum | Savunma dayanak ağırlığını denetleyip ZAYIF/YETERLİ kararı üretir; dayanaksız kalemleri listeler (`SdSavunmaGucu`) |

Ek ayrımlar: tornado duyarlılık (`modulO2TornadoZirve`), rakip teklif anomalisi Z skoru
(`modulO1AnomaliSayisi`). 6 ayrım maddesinin tamamı dosyada ad tanımıyla karşılıklıdır →
**GEÇTİ (D02)**.

## 3. D11 — Karar dürüstlüğü

- Boş dosyada karar **VERİ YOK** (veri varmış gibi UYGUN demez).
- Tüm kararlar gerekçe satırıyla birlikte üretilir; aksiyonlar türetilir.
- Karar kuralları KILAVUZ sayfasında açıklanır.

## 4. D10 — Kaynak şeffaflığı

Mevzuata dayanan parametrelerde kaynak `kurum + madde + tarih`; kaynağı olmayanlar açıkça
`Varsayım` işaretlenir ve `KANIT/PARAMETRE_KAYNAKLARI.md` içinde toplu listelenir.

## 5. Dürüstlük maddesi (manda 11)

1. Hangi kapılar geçti? → `RAPOR_DENETIM.md` (G01–G24, 0 KALDI), `RAPOR_OLCEK.md` (Ö1, Ö1b, Ö2; 0 KALDI), `RAPOR_DEGER.md` (D01–D14; 0 KALDI).
2. Uygulanamayan madde yok; üç katman da tamamlandı.
3. En zayıf nokta: sınır değer katsayısı (N) ve hesap yöntemi ihaleden ihaleye değişir; dosya
   genel eşikleri kullanır, alıcının ihale dokümanındaki değerleri AYARLAR'dan güncellemesi gerekir.
4. Y1+Y2 = 28.000 ₺ ≥ 3 × 7.900 ₺ → fiyat çapası sağlanır.
5. Serbest alternatifle en çok örtüşen özellik: sınır değer hesabı — bu, idare cetvelleriyle
   elle yapılabilen bir girdidir; fark, bölge kararı + savunma gücü + duyarlılık katmanıdır.
6. İçeriği zorladığımız eşik yok; modüller gerçek acıya (aşırı düşük kapsamına girip
   ihaleden elenme riski) dayanır.
