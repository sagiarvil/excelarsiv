# FİYAT SAVUNMASI — KAÇIRILAN SGK TEŞVİKLERİ & GERÇEK İŞÇİLİK MALİYETİ ANALİZİ

**Dosya:** `KacirilanSgkTesvikleriVeGercekIscilikMaliyetiAnalizi.xlsx`
**Sürüm:** 1.1.0
**Üretim tarihi:** 09.08.2026

## 1. D01 — İkame değeri hesabı (fiyat çapası)

Satış fiyatı **7.900 ₺**; manda kuralı: **Y1 + Y2 ≥ 3 × satış fiyatı (23.700 ₺)**.

| Bileşen | Değer (₺) | Açıklama |
|---|---|---|
| Y1 — Danışman ikamesi | 18.000 | Mali müşavir / SGK danışmanının aynı analizi (gerçek işçilik maliyeti + kaçırılan teşvik tespiti + rapor) için asgari 3 aylık emek bedeli |
| Y2 — Kurulum ikamesi | 9.000 | Aynı sistemin sıfırdan kurulması (tasarım, formül motoru, test, dokümantasyon) için iç emek bedeli |
| **Toplam (Y1+Y2)** | **27.000** | ≥ 23.700 → **GEÇTİ (D01)** |

Y3 (önlenen hata / kaçırılan teşviğin paraya dönüşmüş hâli) bu çapaya dahil edilmez; yalnızca
pazarlama değeri olarak kullanılır.

## 2. D02 — Serbest alternatifle ayrışma

"Alıcı bunu bedavaya nasıl yapar?" sorusuna dürüst yanıt:

| Alternatif | Bu dosya nasıl ayrışır (çıktı düzeyinde) |
|---|---|
| SGK e-Sigorta teşvik ekranları | Kaçırılan teşvik tutarını yıllık TL olarak ölçer; portal ekranı bunu göstermez (`KacirilanTesvik`) |
| Mali müşavirin aylık danışmanlığı | Tüm teşvik türlerini tek ekranda simüle eder ve yıllık tutarlarla karşılaştırır (`ToplamYillikTesvik`) |
| Ücretsiz Excel şablonları | İşveren yükünü SGK+işsizlik+kıdem+yan fayda ile gerçek maliyete çevirir (`AylikGercekMaliyet`) |
| Genel amaçlı yapay zekâ ile elle kurulum | Veriye dayalı UYGUN/İNCELE/DURDUR kararı ve makine gerekçesi üretir (`KararSonuc`) |

Ek ayrımlar: tornado duyarlılık (`modulO2TornadoZirve`), sonraki ay maliyet tahmini ve güven
aralığı (`modulI1TahminOrta`). 6 ayrım maddesinin tamamı dosyada ad tanımıyla karşılıklıdır →
**GEÇTİ (D02)**.

## 3. D11 — Karar dürüstlüğü

- Boş dosyada karar **VERİ YOK** (veri varmış gibi UYGUN demez).
- Tüm kararlar gerekçe satırıyla birlikte üretilir; aksiyonlar türetilir.
- Karar kuralları KILAVUZ sayfasında açıklanır.

## 4. D10 — Kaynak şeffaflığı

Mevzuata dayanan parametrelerde kaynak `kurum + madde/tebliğ + tarih`; kaynağı olmayanlar
açıkça `Varsayım` işaretlenir ve `KANIT/PARAMETRE_KAYNAKLARI.md` içinde toplu listelenir.

## 5. Dürüstlük maddesi (manda 11)

1. Hangi kapılar geçti? → `RAPOR_DENETIM.md` (G01–G24, 0 KALDI), `RAPOR_OLCEK.md` (Ö1, Ö1b, Ö2; 0 KALDI), `RAPOR_DEGER.md` (D01–D14; 0 KALDI).
2. Uygulanamayan madde yok; üç katman da tamamlandı.
3. En zayıf nokta: teşvik oranları mevzuatla birlikte değişir; dosya üretim tarihi itibarıyla
   günceldir ancak alıcının oranları periyodik doğrulaması gerekir.
4. Y1+Y2 = 27.000 ₺ ≥ 3 × 7.900 ₺ → fiyat çapası sağlanır.
5. Serbest alternatifle en çok örtüşen özellik: tek çalışanın teşvik türü seçimi — bu, portalda
   elle yapılabilen bir girdidir; fark, kaçırılan tutarın TL ölçümü ve karar katmanıdır.
6. İçeriği zorladığımız eşik yok; modüller gerçek acıya (kaçırılan teşvik parası) dayanır.
