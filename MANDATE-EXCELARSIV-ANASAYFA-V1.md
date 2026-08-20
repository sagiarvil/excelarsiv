# MANDATE-EXCELARSIV-ANASAYFA-V1

**Konu:** excelarsiv.com ana sayfa tamamlama faaliyeti — veri bütünlüğü, ödeme eşleşmesi, fiyat mimarisi ve vitrin yeniden yazımı
**Sürüm:** V1
**Tarih:** 2026-08-20
**Statü:** BAĞLAYICI
**Sahip:** Barış Bağırlar
**Uygulayıcı:** Cursor / geliştirici ajan
**Hedef repo:** excelarsiv (Astro v7.1.6)

---

## 0. BAĞLAYICILIK VE TANIMLAR

Bu doküman bir öneri listesi değildir. Aşağıdaki kurallar uygulayıcı için bağlayıcıdır.

| Terim | Anlam |
|---|---|
| **GATE** | Sıralı kapı. Bir GATE'in tüm invariantları YEŞİL olmadan sonraki GATE'e geçilemez. |
| **INV** | Makine ile doğrulanabilir invariant. Öznel değerlendirme içermez. |
| **KANIT** | Invariantın sağlandığını gösteren, repoya işlenen dosya (log, JSON, ekran görüntüsü, script çıktısı). |
| **ROLLBACK** | Invariant kırıldığında uygulanacak, tek komutla geri dönülebilir işlem. |
| **YEŞİL** | `verify_anasayfa.py` ilgili GATE için `exit 0` döndürür ve KANIT dosyası repoda mevcuttur. |

**Uygulayıcı yetkisi:** Uygulayıcı bu dokümandaki hiçbir invariantı yorumlayarak gevşetemez, "daha iyi olur" gerekçesiyle kapsam ekleyemez, bir GATE'i atlayamaz. Invariant uygulanamıyorsa iş durdurulur ve sahibe INV kodu ile bildirilir.

**Kapsam dışı (bu mandate'te YAPILMAYACAK):**
- Tekil ürün fiyatlarının değiştirilmesi
- Yeni Excel ürünü üretimi
- Ürün detay sayfalarının yeniden tasarımı
- Blog / rehber içerik üretimi
- Ödeme sağlayıcısının değiştirilmesi

---

## 1. FAALİYETİN GEREKÇESİ (canlı denetim bulguları)

2026-08-20 tarihinde excelarsiv.com ve /sablonlar canlı HTML üzerinden denetlendi.

| Kod | Bulgu | Şiddet |
|---|---|---|
| **B1** | Ana sayfada aynı ürün (*13 Haftalık Nakit Akışı*) için iki farklı fiyat: vitrin kartında 999 TL, "Gerçek ürün ekranları" bölümünde 2.490 TL. Katalogda 999 TL. | BLOKE EDİCİ |
| **B2** | 51 ürün yalnızca 4 Shopier kaydına yönleniyor (bant başına bir kayıt: 49652321 / 49652403 / 49653399 / 49653437). Alıcı ödeme ekranında satın aldığı sistemin adını göremiyor. | BLOKE EDİCİ |
| **B3** | Katalogda 51 sistem / 7 alan var; ana sayfada 5 ürün gösteriliyor, "51" sayısı hiç geçmiyor. Ölçek kanıtı kullanılmıyor. | YÜKSEK |
| **B4** | Ana sayfa H1'i konumu "şablon" kelimesine sabitliyor. | YÜKSEK |
| **B5** | `/kurumsal-lisans`, `/ortaklik-mali-musavir`, `/basari-hikayeleri`, `/neden-excel-arsiv` sayfaları mevcut; ana sayfa gövdesinden erişilemiyor. | YÜKSEK |
| **B6** | Footer tutarsız: ana sayfa footer'ında "Kurumsal Lisans" ve "Mali Müşavir Ortaklık" var, /sablonlar footer'ında yok. | ORTA |
| **B7** | Ana sayfada aynı işi yapan 4 ayrı keşif bölümü var (problem seçici, kategori keşfi, premium vitrin, gerçek ürün ekranları). | ORTA |
| **B8** | Ürün kartlarında rastgele renk kodlaması (mavi/turuncu/mor/yeşil), monospace font buton ve nav'da kullanımı, hero'da stok/AI görsel. | ORTA |
| **B9** | Hero görselinin `src` özniteliği boş dönüyor. LCP ve indekslenebilirlik doğrulanmalı. | ORTA |

**Faaliyetin tek cümlelik amacı:** Ana sayfayı, 51 denetlenmiş sistemin ölçeğini ve karar değerini gösteren, tek fiyat kaynağından beslenen, ürün bazlı ödeme eşleşmesine sahip bir satış vitrinine dönüştürmek.

---

## 2. GLOBAL INVARIANTLAR (her GATE'te geçerli)

| Kod | Invariant | Doğrulama |
|---|---|---|
| **INV-G01** | Sitede görünen hiçbir fiyat, ürün veri kaynağındaki değerden farklı olamaz. | `verify_anasayfa.py --check prices` |
| **INV-G02** | Hiçbir Shopier ürün ID'si birden fazla ürüne bağlanamaz. | `--check shopier` |
| **INV-G03** | Tüm sayfalarda footer DOM'u birebir aynı olmalı. | `--check footer` |
| **INV-G04** | Her sayfada tam olarak bir `<h1>` bulunmalı. | `--check headings` |
| **INV-G05** | Yayına giden hiçbir sayfada `lorem`, `TODO`, `placeholder`, `örnek metin` geçemez. | `--check placeholders` |
| **INV-G06** | Uydurma sosyal kanıt yasaktır: doğrulanamayan müşteri sayısı, isimsiz referans, yıldız puanı, "binlerce işletme" türü ifade kullanılamaz. | Manuel + `--check claims` |
| **INV-G07** | Hiçbir sayfada stok fotoğraf, 3D render, generic gradient/glassmorphism görsel kullanılamaz. Yalnız gerçek ürün ekranı ve gerçek portre kullanılır. | Manuel + görsel envanteri KANIT |
| **INV-G08** | Her değişiklik tek branch'te (`feat/anasayfa-v1`) yapılır; her GATE tek commit ile kapanır ve commit mesajı GATE kodunu içerir. | `git log` KANIT |

---

## 3. VERİ SÖZLEŞMESİ

Bu faaliyetin temeli tek doğruluk kaynağıdır. Aşağıdaki şema zorunludur.

### 3.1 Ürün kaydı (tek kaynak)

```yaml
slug: 13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi
ad: "13 Haftalık Nakit Akışı ve Ödeme Planlama Sistemi"
kategori: "Nakit Akışı"
sayfa_sayisi: 18
fiyat_tl: 999            # tek kaynak — hiçbir sayfada elle yazılmaz
kdv_dahil: true
shopier_id: "49653399"   # ÜRÜNE ÖZEL, banda değil (GATE-2 sonrası benzersiz)
shopier_urun_adi: "13 Haftalık Nakit Akışı ve Ödeme Planlama Sistemi"
kapak_gorsel: /images/kapak/13-haftalik-nakit-akisi-....webp
demo_var: true
```

**Kural:** Şablonlarda fiyat yalnızca `{urun.fiyat_tl}` üzerinden basılır. Kaynak kodda ` TL` ile biten sabit sayı bulunması INV-G01 ihlalidir.

### 3.2 Paket kaydı

```yaml
slug: mali-musavir-vergi-paketi
ad: "Mali Müşavir Vergi Paketi"
fiyat_tl: 3490
icerik_slugs: [kdv-iadesi-azami-alacak-hesabi-dosya-hazirlayici,
               kdv-iade-listesi-robotu-gib7,
               kdv-tevkifat-mahsup-iade-listesi,
               cari-ba-bs-toplu-mutabakat,
               ymm-tasdik-kontrol-robotu]
shopier_id: "<yeni>"
```

**INV-P01:** Her paketin fiyatı, içerdiği ürünlerin tekil fiyat toplamından **kesinlikle düşük** olmalıdır. Aksi halde paket yayınlanamaz.
**INV-P02:** Paket sayfasında dosya listesi, dosya adedi, toplam sayfa sayısı, tekil toplam ve tasarruf tutarı **hesaplanarak** gösterilir; elle yazılmaz.

---

## 4. GATE'LER

### GATE-0 — ÖLÇÜM TEMELİ

**Amaç:** Değişikliğin etkisinin ölçülebilir olması. Bu GATE atlanırsa faaliyetin başarısı hakkında hiçbir karar verilemez.

| Kod | Invariant |
|---|---|
| INV-0.1 | Şu event'ler kurulu ve veri üretiyor: `hero_cta_primary`, `hero_cta_demo`, `paket_view`, `paket_click`, `urun_card_click`, `kurumsal_form_submit`, `scroll_50`, `scroll_75`. |
| INV-0.2 | 7 günlük baseline toplanmış; ana sayfa oturum sayısı, `scroll_75` oranı, ürün sayfasına geçiş oranı, AOV, tekil satış adedi kayıt altına alınmış. |
| INV-0.3 | Baseline verisi `kanit/GATE-0-baseline.json` olarak repoya işlenmiş. |

**KANIT:** `kanit/GATE-0-baseline.json`
**ROLLBACK:** Yok (yalnız ölçüm katmanı eklenir).
**KABUL:** 8 event de en az 1 kayıt üretmiş; baseline JSON'da 5 metrik dolu.

---

### GATE-1 — FİYAT TEK KAYNAK (B1)

**Amaç:** Aynı ürünün sitede iki farklı fiyatla görünmesini yapısal olarak imkânsız kılmak.

| Kod | Invariant |
|---|---|
| INV-1.1 | Ana sayfadaki "Gerçek ürün ekranları" bölümündeki sabit kodlu fiyat alanı kaldırılmış; değer ürün kaydından okunuyor. |
| INV-1.2 | Kaynak kodda fiyat içeren sabit metin yok: `grep -rEn "[0-9]{3,4}(\.[0-9]{3})? ?TL" src/` yalnızca veri dosyalarını döndürür, şablon/bileşen dosyalarını değil. |
| INV-1.3 | Ana sayfa, `/sablonlar` ve ürün detay sayfasında aynı slug için basılan fiyat birebir aynı. |
| INV-1.4 | JSON-LD `Product.offers.price` değeri, sayfada görünen fiyatla aynı ve `priceCurrency: TRY`. |
| INV-1.5 | KDV ifadesi tüm yüzeylerde tek biçim: `KDV dahil · tek ödeme`. |

**KANIT:** `kanit/GATE-1-fiyat-tutarlilik.json` (slug × yüzey × fiyat matrisi, fark sayısı = 0)
**DOĞRULAMA:** `python verify_anasayfa.py --check prices --base <preview-url>`
**ROLLBACK:** Tek commit revert.
**KABUL:** Fark sayısı **0**. 2.490 TL ibaresi sitede hiçbir yerde geçmiyor (paket fiyatı olarak geçmesi hariç).

---

### GATE-2 — SHOPIER ÜRÜN EŞLEŞMESİ (B2)

**Amaç:** Alıcının ödeme ekranında satın aldığı sistemin adını görmesi; iade/uyuşmazlıkta ödeme kaydının ürünü tanımlaması.

| Kod | Invariant |
|---|---|
| INV-2.1 | Katalogdaki her ürün için ayrı Shopier kaydı açılmış; `shopier_id` alanı 51 ürün için **benzersiz**. |
| INV-2.2 | Shopier'deki ürün adı, sitedeki ürün adıyla birebir aynı (`shopier_urun_adi == ad`). |
| INV-2.3 | Shopier'deki fiyat, ürün kaydındaki `fiyat_tl` ile aynı. |
| INV-2.4 | `veri/shopier-eslesme.csv` dosyası repoda: `slug,ad,fiyat_tl,shopier_id`. |
| INV-2.5 | CI kontrolü: aynı `shopier_id` iki satırda geçemez. |
| INV-2.6 | Teslimat akışı, ödeme sonrası doğru dosyayı ayırt ediyor: 3 farklı banttan yapılan test satın almasında doğru dosya indiriliyor. |

**KANIT:** `veri/shopier-eslesme.csv` + `kanit/GATE-2-satinalma-testi.md` (3 test satın alması, ekran görüntüleriyle)
**DOĞRULAMA:** `python verify_anasayfa.py --check shopier`
**ROLLBACK:** Eski bant ID'leri `veri/shopier-eslesme.bak.csv` içinde saklanır; geri dönüş tek dosya değişimi.
**KABUL:** 51 benzersiz ID; 3 test satın almasının 3'ünde de doğru ürün adı ve doğru dosya.

> **Not:** Bu GATE tamamlanmadan GATE-5'e (tasarım) geçilmesi yasaktır. Yeni tasarım hatalı eşleşmeyi miras alırsa düzeltme maliyeti iki katına çıkar.

---

### GATE-3 — FOOTER VE KURUMSAL YÜZEYLER (B5, B6)

| Kod | Invariant |
|---|---|
| INV-3.1 | Tek footer bileşeni; tüm sayfalarda aynı DOM. |
| INV-3.2 | Footer'da şu linkler her sayfada mevcut: Kurumsal Lisans, Mali Müşavir Ortaklık, Başarı Hikâyeleri, Neden Excel Arşiv. |
| INV-3.3 | `/kurumsal-lisans` ve `/ortaklik-mali-musavir` sayfaları ana sayfa **gövdesinden** (footer değil) en az birer kez linklenmiş. |
| INV-3.4 | Bu iki sayfada teklif formu çalışıyor; zorunlu alanlar: firma unvanı, VKN (10 hane doğrulaması), kullanıcı sayısı, ihtiyaç. |
| INV-3.5 | Form gönderimi rate-limit'li ve e-posta olarak ulaşıyor. |

**KANIT:** `kanit/GATE-3-form-testi.md` (3 test gönderimi)
**DOĞRULAMA:** `--check footer --check corporate-links`
**ROLLBACK:** Tek commit revert.
**KABUL:** Footer farkı 0; iki sayfa gövdeden linkli; 3 test formunun 3'ü de ulaştı.

---

### GATE-4 — PAKET KATMANI

**Amaç:** Ana sayfada fiyat merdiveni kurmak. Yeni ürün üretilmez; mevcut dosyalar paketlenir.

**Yayınlanacak paketler:**

| Paket | İçerik (slug) | Fiyat | Tekil toplam |
|---|---|---|---|
| Mali Müşavir Vergi Paketi | kdv-iadesi-azami-alacak (1.499) + kdv-iade-listesi-robotu-gib7 (799) + kdv-tevkifat-mahsup (799) + cari-ba-bs-toplu-mutabakat (499) + ymm-tasdik-kontrol-robotu (999) | **3.490 TL** | 4.595 TL |
| Patron Karar Paketi | aylik-patron-finans-paneli (999) + sube-karlilik (999) + proje-ve-is-bazinda-karlilik (999) + ttk-376-sermaye-tamamlama (1.499) | **3.490 TL** | 4.496 TL |
| Nakit Kontrol Paketi | akilli-kasa-defteri (499) + 13-haftalik-nakit-akisi (999) + cari-hesap-tahsilat-risk (799) + cek-senet-vade-risk (799) | **2.490 TL** | 3.096 TL |

| Kod | Invariant |
|---|---|
| INV-4.1 | Her paket fiyatı, içerik toplamından düşük (INV-P01). |
| INV-4.2 | Her paketin ayrı Shopier kaydı ve ayrı ürün sayfası var. |
| INV-4.3 | Paket sayfasında dosya listesi, dosya adedi, toplam sayfa sayısı, tekil toplam ve tasarruf **hesaplanarak** basılıyor. |
| INV-4.4 | Paket ZIP'i, içerik slug'larındaki güncel dosyalardan üretiliyor; içerik listesiyle ZIP içeriği birebir eşleşiyor. |
| INV-4.5 | Paket sayfasında güncelleme politikası yazılı: kapsam, süre, ücretsiz güncelleme sınırı. |
| INV-4.6 | Lisans metni paket satın alımını kapsıyor. |

**KANIT:** `kanit/GATE-4-paket-dogrulama.json` (paket × içerik × fiyat, ZIP manifest karşılaştırması)
**ROLLBACK:** Paket sayfaları `draft: true` ile kapatılır; ana sayfa merdiveni tekil vitrine döner.
**KABUL:** 3 paket yayında; ZIP manifest farkı 0; her paket sayfasında güncelleme politikası mevcut.

---

### GATE-5 — ANA SAYFA YENİDEN YAZIMI (B3, B4, B7, B8, B9)

#### 5.1 Bölüm mimarisi — ZORUNLU SIRA VE SAYI

Ana sayfa **tam olarak 7 bölümden** oluşur. Fazlası yasak, sırası bağlayıcı.

| # | Bölüm | Tek işi |
|---|---|---|
| 1 | Otorite hero | Ölçek + denetim iddiası, iki CTA |
| 2 | Tasarlayan + ölçek şeridi | Kim yaptı, kaç sistem, hangi alanlar |
| 3 | Farklılaştırıcı (tek koyu blok) | "ChatGPT tablo üretir. Biz denetlenmiş karar sistemi satarız." |
| 4 | Fiyat merdiveni | Kurumsal Lisans → 3 paket → "51 tekil sistem" linki |
| 5 | Problem seçici | 6 problem kartı (diğer keşif bölümleri silinir) |
| 6 | Kurumsal lisans şeridi | Çok kullanıcılı lisans + uyarlama + proforma teklif |
| 7 | Güven + SSS + kapanış CTA | Ödeme/teslimat tek satır, 5 SSS, son CTA |

**Silinecek bölümler:** kategori keşfi ("İşiniz hangi tarafta sıkışıyorsa") → `/sablonlar`a taşınır; "Gerçek ürün ekranları" sekmeli bölümü → ürün detay sayfalarına; "Dosyanın içine bakın" (CF-01..04) → `/neden-excel-arsiv` sayfasına; Shopier 3 adımlı kahraman bölümü → tek satır rozete iner.

#### 5.2 Hero içerik sözleşmesi

```
Eyebrow: 51 DENETLENMİŞ FİNANSAL KARAR SİSTEMİ

H1: Mevzuata ve sahaya göre denetlenmiş Excel karar sistemleri.

Alt metin: KDV iadesinden TTK 376'ya, konkordato ön projesinden Logo cari
yaşlandırmasına kadar 51 sistem. Her biri açık formül, girdi disiplini ve
yöneticiye sunulabilir karar çıktısıyla.

CTA-1 (birincil): 51 sistemi inceleyin  → /sablonlar
CTA-2 (ikincil):  Ücretsiz demo         → /demo

Kanıt şeridi: 51 sistem · 7 işletme alanı · 13–24 sayfa · açık formül · KDV dahil tek ödeme

Sağ kolon: 13 Haftalık Nakit Akışı sisteminin SONUÇ ekranı (NET NAKİT / GİRİŞ /
ÇIKIŞ / MİN. HAFTA + kümülatif eğri). Formül metni gösteren görsel kullanılamaz.
```

| Kod | Invariant |
|---|---|
| INV-5.1 | Ana sayfada `<section>` sayısı **≤ 7**. |
| INV-5.2 | H1'de "şablon" kelimesi geçmez. `/sablonlar` sayfasının H1'inde geçmeye devam eder. |
| INV-5.3 | Hero'da arama kutusu ve "Popüler" çip listesi yok; her ikisi `/sablonlar`a taşınmış. |
| INV-5.4 | Hero'da stok/3D render görsel yok; `src` boş `<img>` yok, tüm görsellerin `src` ve `alt` değeri dolu. |
| INV-5.5 | Ana sayfada 2.490 TL'nin altında hiçbir fiyat basılmıyor. Tekil fiyatlar yalnız `499 TL'den başlar` ifadesiyle, link metninde geçer. |
| INV-5.6 | Sayfada **tek** koyu zeminli bölüm var. |
| INV-5.7 | Ana sayfadaki iç link sayısı **≥ 41** (2026-08-20 ölçümü). |
| INV-5.8 | `/kurumsal-lisans`, `/ortaklik-mali-musavir`, `/basari-hikayeleri`, `/neden-excel-arsiv`, `/sablonlar`, `/demo` ana sayfa gövdesinden linklenmiş. |

#### 5.3 Tasarım invariantları (AI izlenimini kesen kurallar)

| Kod | Invariant | Doğrulama |
|---|---|---|
| INV-D1 | Tek aksan rengi ailesi (koyu yeşil) + nötr gri skala. Mavi/turuncu/mor aksan yok. | CSS hex taraması |
| INV-D2 | Renk yalnız **durum** taşır (pozitif/negatif/uyarı). Kategori veya dekorasyon amaçlı renk yok. | Manuel + kart bileşeni tek dosyadan |
| INV-D3 | Ürün/paket kartları tek bileşenden üretilir; kart başına özel renk/buton stili yok. | Bileşen sayımı |
| INV-D4 | Monospace font yalnız sayısal veri, kod ve eyebrow'da. Buton, nav ve gövde metninde yasak. | CSS sınıf denetimi |
| INV-D5 | `border-radius: 0`; gölge yok, `1px` hairline sınır. | CSS taraması |
| INV-D6 | Tipografi: tek sans + tek mono. H1 `clamp(2.5rem, 5vw, 4rem)`, gövde satır uzunluğu 62–72ch. | CSS taraması |
| INV-D7 | Eyebrow (mono üst etiket) en fazla 3 bölümde kullanılır. | DOM sayımı |
| INV-D8 | En fazla 2 CTA türü: birincil (sistemleri incele) ve ikincil (demo). Üçüncü bir buton stili yasak. | Bileşen sayımı |

**KANIT:** `kanit/GATE-5-anasayfa-denetim.json` + 320/768/1440/1920px ekran görüntüleri
**ROLLBACK:** `feat/anasayfa-v1` branch'i; eski ana sayfa `src/pages/index.legacy.astro` olarak 30 gün saklanır.
**KABUL:** INV-5.x ve INV-D1..D8'in tamamı YEŞİL.

---

### GATE-6 — YAYIN ÖNCESİ DENETİM

| Kod | Invariant |
|---|---|
| INV-6.1 | Lighthouse mobil: Performans ≥ 90, Erişilebilirlik ≥ 95, SEO ≥ 95, Best Practices ≥ 95. |
| INV-6.2 | LCP < 2.0 s (mobil, 4G kısıtı). |
| INV-6.3 | Mevcut SEO invariant script'i (MANDATE-EXCELARSIV-SEO-V1) **0 ihlal** döndürür. |
| INV-6.4 | 320px–1920px arasında yatay taşma ve kırılma yok. |
| INV-6.5 | Klavye ile tüm CTA'lara erişilebiliyor; odak halkası görünür; kontrast oranı ≥ 4.5:1. |
| INV-6.6 | JSON-LD doğrulaması hatasız (Organization, WebSite, Product/ItemList). |
| INV-6.7 | `verify_anasayfa.py --check all` → `exit 0`. |

**KANIT:** `kanit/GATE-6-lighthouse.json`, `kanit/GATE-6-verify-all.log`
**ROLLBACK:** Preview'dan yayına geçilmez.
**KABUL:** 7 invariantın tamamı YEŞİL.

---

### GATE-7 — YAYIN VE ÖLÇÜM PENCERESİ

| Kod | Invariant |
|---|---|
| INV-7.1 | Yayın sonrası **14 gün** boyunca ana sayfada başka değişiklik yapılmaz (ölçüm bütünlüğü). |
| INV-7.2 | 14. ve 30. günde metrikler `kanit/GATE-7-olcum-<gun>.json` olarak kaydedilir. |
| INV-7.3 | Search Console'da ana sayfa organik oturumları günlük izlenir. |

---

## 5. KARAR KAPILARI (30 gün sonunda)

### DEVAM koşulları — hepsi sağlanmalı
- AOV ≥ **1.800 TL** (GATE-0 baseline'ına göre yükselmiş)
- En az **2 paket satışı** veya **1 kurumsal/mali müşavir teklifi**
- Organik ana sayfa oturumu ≥ baseline
- `scroll_75` ≥ **%45**

### DÜZELTME koşulları
| Gözlem | Aksiyon |
|---|---|
| `paket_view` yüksek, `paket_click` düşük | Paket kartı değer ifadesi yeniden yazılır; fiyat sabit kalır |
| `paket_click` yüksek, satış yok | Paket sayfasına "hangi problem zincirini çözer" akış şeması eklenir |
| Mali müşavir paketi satıyor, patron paketi satmıyor | Ana sayfa dili mali müşavir / YMM segmentine kaydırılır |
| Kurumsal form 30 günde 0–1 talep | Bölüm 6 tek satır rozete indirilir |

### DURDURMA koşulları — biri yeterli
| Koşul | Aksiyon |
|---|---|
| Organik ana sayfa oturumu 30 günde **−%15**'ten fazla düştü | GATE-5.2'deki H1 değişimi geri alınır (yalnız hero) |
| Tekil satış adedi 30 günde **−%25**'ten fazla düştü | Paket katmanı ana sayfadan indirilir, tekil vitrin geri gelir |
| Lighthouse mobil performans kalıcı **< 85** | GATE-5 geri alınır |
| 60 günde 0 paket + 0 kurumsal teklif | Enterprise hedefi bu fiyat yapısıyla test edilmiş ve **reddedilmiş** sayılır; konum "51 sistemlik profesyonel finans kütüphanesi" olarak sadeleştirilir |

---

## 6. KARAR ETİKETİ VE GERİ DÖNÜŞ ÖZETİ

**Karar: [Geri döndürülebilir]**

| GATE | Geri dönüş maliyeti |
|---|---|
| GATE-1 | Tek commit revert |
| GATE-2 | `veri/shopier-eslesme.bak.csv` geri yüklenir (yeni Shopier kayıtları pasife alınır) |
| GATE-3 | Tek commit revert |
| GATE-4 | Paket sayfaları `draft: true` |
| GATE-5 | `index.legacy.astro` geri alınır (30 gün saklanır) |

**Geri döndürülemez tek unsur:** Yeni Shopier ürün kayıtları oluşturulması. Bu kayıtlar zarar vermez, pasife alınabilir — ancak silinmeleri geçmiş sipariş kayıtlarını etkileyebileceğinden pasife alma tercih edilir.

---

## 7. UYGULAYICIYA KAPALI ALANLAR

Aşağıdakiler bu mandate kapsamında **yapılmayacaktır**. Yapılırsa mandate ihlalidir:

1. Tekil ürün fiyatını değiştirmek
2. Yeni Excel ürünü veya yeni kategori üretmek
3. Ana sayfaya 7'den fazla bölüm eklemek
4. "Binlerce müşteri", "Türkiye'nin 1 numarası" gibi doğrulanamayan iddia yazmak
5. Sahte referans, sahte logo, sahte yıldız puanı eklemek
6. Koyu tema uygulamak (tek koyu blok dışında)
7. Stok görsel veya AI üretimi illüstrasyon kullanmak
8. Ödeme sağlayıcısını değiştirmek veya kart verisi toplayan bir akış kurmak
9. GATE sırasını değiştirmek veya bir GATE'i "sonra yaparız" diyerek atlamak

---

## 8. TESLİM PAKETİ

Faaliyet tamamlandığında repoda bulunması zorunlu dosyalar:

```
kanit/GATE-0-baseline.json
kanit/GATE-1-fiyat-tutarlilik.json
kanit/GATE-2-satinalma-testi.md
kanit/GATE-3-form-testi.md
kanit/GATE-4-paket-dogrulama.json
kanit/GATE-5-anasayfa-denetim.json
kanit/GATE-5-ekran-goruntuleri/ (320, 768, 1440, 1920)
kanit/GATE-6-lighthouse.json
kanit/GATE-6-verify-all.log
kanit/GATE-7-olcum-14.json
kanit/GATE-7-olcum-30.json
veri/shopier-eslesme.csv
veri/shopier-eslesme.bak.csv
verify_anasayfa.py
src/pages/index.legacy.astro   (30 gün saklanır)
```

---

**Bu mandate, KANIT dosyalarının tamamı repoda mevcut ve `verify_anasayfa.py --check all` çıktısı `exit 0` olduğunda tamamlanmış sayılır.**

---

## EK A — ÖLÇÜLEN BAŞLANGIÇ DURUMU (2026-08-20, canlı site)

`python verify_anasayfa.py --check homepage --base https://excelarsiv.com` çıktısı:

```
[KIRMIZI] INV-5.1: <section> sayisi 10 (ust sinir 7)
[YESIL]   INV-G04: <h1> sayisi 1
[KIRMIZI] INV-5.2: H1 "işletmeler için hazır excel ŞABLONLARI ve karar sistemleri"
[KIRMIZI] INV-5.3: Ana sayfada arama input sayisi 1 (olmasi gereken 0)
[YESIL]   INV-5.4: src/alt bos gorsel 0
[KIRMIZI] INV-5.5: 2490 TL altinda fiyat: [499, 799, 999]
[KIRMIZI] INV-5.8: Eksik zorunlu link: /kurumsal-lisans, /ortaklik-mali-musavir
[ATLA]    INV-5.7: mevcut ic link sayisi 41  → baseline olarak sabitlendi
[YESIL]   INV-G05: Placeholder ifade yok
[YESIL]   INV-G06: Dogrulanamaz iddia yok

TOPLAM: 9 invariant | YESIL 4 | KIRMIZI 5 → GATE KAPALI
```

**Başlangıç skoru: 9 invariantın 4'ü yeşil.** Faaliyet, bu 5 kırmızının kapatılması ve GATE-1/2/3/4'ün eklenmesiyle tamamlanır.

Bu çıktı `kanit/GATE-5-baseline-run.txt` olarak repoya işlenmelidir; sonraki koşularla karşılaştırma temeli budur.
