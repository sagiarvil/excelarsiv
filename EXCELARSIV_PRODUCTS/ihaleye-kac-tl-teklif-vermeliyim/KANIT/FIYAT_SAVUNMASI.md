# FİYAT SAVUNMASI — İHALEYE KAÇ TL TEKLİF VERMELİYİM? (SINIR DEĞER HESAPLAYICI)

**Dosya:** `IhaleyeKacTlTeklifVermeliyim.xlsx`
**Sürüm:** 1.0.0
**Üretim tarihi:** 09.08.2026

## 1. D01 — İkame değeri hesabı (fiyat çapası)

Satış fiyatı **2.490 ₺**; manda kuralı: **Y1 + Y2 ≥ 3 × satış fiyatı (7.470 ₺)**.

| Bileşen | Değer (₺) | Açıklama |
|---|---|---|
| Y1 — Danışman ikamesi | 8.500 | İhale mevzuat danışmanının teklif başına sınır değer hesabı + teklif konumlandırma ve marj kontrolü emek bedeli |
| Y2 — Kurulum ikamesi | 3.000 | Aynı sistemin sıfırdan kurulması (tasarım, formül motoru, test, dokümantasyon) için iç emek bedeli |
| **Toplam (Y1+Y2)** | **11.500** | ≥ 7.470 → **GEÇTİ (D01)** |

Y3 (önlenen hata / ihale kaybına dönüşmüş para) bu çapaya dahil edilmez; yalnızca pazarlama
değeri olarak kullanılır.

## 2. D02 — Serbest alternatifle ayrışma

"Alıcı bunu bedavaya nasıl yapar?" sorusuna dürüst yanıt:

| Alternatif | Bu dosya nasıl ayrışır (çıktı düzeyinde) |
|---|---|
| İdarelerin açık yayımladığı sınır değer hesaplama cetvelleri | Sınır değeri N katsayısı ve hesap yöntemiyle otomatik hesaplar; yöntem değişiminde cetveli elle yeniden kurmayı gerektirmez (`SinirDeger`) |
| İhale mevzuat danışmanının teklif başı danışmanlığı | Öneri teklifini hedef kâr marjına göre TL ve sınır değer oranı bazında üretir; risk bölgesini otomatik işaretler (`OneriTeklif`) |
| Ücretsiz Excel şablonları | Rakip teklif anomalilerini Z skoruyla tespit eder; geçersiz teklifleri ortalama ve sınır değer hesabından ayırır (`modulO1AnomaliSayisi`) |
| Genel amaçlı yapay zekâ ile elle kurulum | Maliyet kalemleri ile toplam maliyet mutabakatını tolerans içinde denetler; sapmayı TL ve oran bazında gösterir (`MutabakatFark`) |

Ek ayrımlar: geçmiş ihalelerden kazanma oranı ve sınır değer iskontosu tahmini + tahmin alt/üst
bandı (`modulI1TahminOrta`), tornado duyarlılıkla en etkili değişken sıralaması
(`modulO2TornadoZirve`). 6 ayrım maddesinin tamamı dosyada ad tanımıyla karşılıklıdır →
**GEÇTİ (D02)**.

## 3. D11 — Karar dürüstlüğü

- Boş dosyada karar **VERİ YOK** (veri varmış gibi UYGUN demez).
- Uç durumlar: negatif/aşırı büyük tutar → **VERİ YOK / VERİ YETERSİZ**; tüm kararlar gerekçe
  satırıyla üretilir; aksiyonlar türetilir.
- Karar kuralları KILAVUZ sayfasında açıklanır.

## 4. D10 — Kaynak şeffaflığı

Mevzuata dayanan parametrelerde kaynak `kurum + madde + tarih`; kaynağı olmayanlar açıkça
`Varsayım` işaretlenir ve `KANIT/PARAMETRE_KAYNAKLARI.md` içinde toplu listelenir.

## 5. Dürüstlük maddesi (manda 11)

1. Hangi kapılar geçti? → `RAPOR_DENETIM.md` (G01–G24, 0 KALDI), `RAPOR_OLCEK.md` (Ö1, Ö1b, Ö2; 0 KALDI), `RAPOR_DEGER.md` (D01–D14; 0 KALDI).
2. Uygulanamayan madde yok; üç katman da tamamlandı.
3. En zayıf nokta: sınır değer katsayısı (N) ve hesap yöntemi ihaleden ihaleye değişir; dosya
   genel eşikleri kullanır, alıcının ihale dokümanındaki değerleri AYARLAR'dan güncellemesi gerekir.
4. Y1+Y2 = 11.500 ₺ ≥ 3 × 2.490 ₺ → fiyat çapası sağlanır.
5. Serbest alternatifle en çok örtüşen özellik: sınır değer hesabı — bu, idare cetvelleriyle
   elle yapılabilen bir girdidir; fark, marj bazlı teklif önerisi + anomali + mutabakat +
   tahmin + duyarlılık katmanıdır.
6. İçeriği zorladığımız eşik yok; modüller gerçek acıya (yanlış teklifle ihaleyi kaybetme veya
   marjın altında iş alma riski) dayanır.
