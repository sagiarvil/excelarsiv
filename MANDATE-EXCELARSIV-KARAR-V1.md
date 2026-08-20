# MANDATE-EXCELARSIV-KARAR-V1

**Konu:** Satın alma niyetli doğal dil sorguları için karar (recommendation) sayfa kümesi + ana sayfa karar motoru bölümü
**Sürüm:** V1 — MANDATE-EXCELARSIV-ANASAYFA-V1'in **EK'idir**, onun yerine geçmez
**Tarih:** 2026-08-20
**Statü:** BAĞLAYICI
**Bağımlılık:** GATE-1 ve GATE-2 (fiyat tek kaynak + Shopier ürün eşleşmesi) YEŞİL olmadan bu mandate başlatılamaz

---

## 0. YAPISAL KARAR — ÇAKIŞMANIN ÇÖZÜMÜ

ANASAYFA-V1, ana sayfayı **7 bölümle** sınırlar (INV-5.1). Yeni bir "karar motoru" bölümü eklenirse bu invariant kırılır.

**Karar:** Karar motoru, **Bölüm 5'in (problem seçici) yerine geçer** — yeni bölüm eklenmez.
Gerekçe: ikisi de aynı işi yapar (ziyaretçiyi ihtiyacından ürüne götürmek). Karar motoru bunu daha güçlü yapar: 6 genel problem kartı yerine, insanların gerçekten yazdığı 18 satın alma sorgusuna doğrudan cevap verir.

**Sonuç:** Ana sayfa yine 7 bölüm. INV-5.1 korunur.

---

## 1. FAALİYETİN İKİ AYAĞI

| Ayak | Ne | Nerede |
|---|---|---|
| **A** | 18 karar sayfası (`/karar/<slug>`) — her biri bir satın alma sorgu kümesine cevap verir | Yeni sayfa kümesi |
| **B** | Ana sayfa karar motoru bölümü — 18 sayfanın en güçlü 8'ine giriş | Ana sayfa Bölüm 5 |

**Amaç:** Hem Google'da "…Excel'i" arayan alıcıyı, hem ChatGPT/Gemini/Perplexity'e "KOBİ için nakit akışı Excel'i önerir misin" diye soran alıcıyı yakalamak.

**Dürüst sınır [Varsayım]:** Üretken motorların hangi kaynağı alıntıladığı garanti edilemez ve algoritmaları açıklanmaz. Aşağıdaki yapı, bilinen alıntılanabilirlik sinyallerine (net soru-cevap yapısı, yapılandırılmış veri, tablo, dürüst sınır ifadeleri, sabit URL) dayanır — **garanti değil, olasılık artırıcıdır.** Ölçüm zorunludur (GATE-K5).

---

## 2. SORGU KÜMESİ HARİTASI (18 sayfa)

Her sayfanın **tek birincil ürünü** vardır. Aynı ürün iki sayfada birincil olamaz (yamyamlık önlemi, INV-K3.4).

| # | URL `/karar/…` | Hedef sorgu ifadesi | Birincil ürün (slug) | Fiyat | Alternatif (en fazla 2) |
|---|---|---|---|---|---|
| 01 | `hangi-excel-sistemini-almaliyim` | "hangi Excel'i almalıyım", "bana hangi Excel şablonu lazım" | *yönlendirici sayfa* | — | Tüm küme + `/urun-bulucu` |
| 02 | `kobi-nakit-akisi-excel` | "KOBİ için nakit akışı Excel'i", "13 haftalık nakit akış tablosu" | `13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi` | 999 | Akıllı Kasa Defteri, KOBİ Finans Yönetim Paketi |
| 03 | `kasa-defteri-excel` | "günlük kasa defteri Excel", "kasa sayım farkı takibi" | `akilli-kasa-defteri-ve-nakit-kontrol-sistemi` | 499 | 13 Haftalık Nakit |
| 04 | `mali-musavir-cari-takip-excel` | "muhasebeciler için cari takip Excel'i", "cari hesap takip tablosu" | `cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi` | 799 | Cari Ba-Bs Mutabakat, Çek-Senet Vade Risk |
| 05 | `pos-komisyon-kontrol-excel` | "POS komisyon kontrol Excel'i", "POS net tahsilat hesaplama" | `pos-komisyon-ve-net-tahsilat-kontrol-sistemi` | 499 | Günlük Gelir-Gider Kârlılık |
| 06 | `trendyol-pazaryeri-net-kar-excel` | "Trendyol komisyon sonrası net kâr Excel", "pazaryeri kârlılık tablosu" | `trendyol-komisyon-sonrasi-net-kar` | 499 | Pazaryeri Net Kâr & Eksik Hakediş |
| 07 | `kdv-iade-dosyasi-excel` | "KDV iade listesi hazırlama Excel", "GİB 7 liste" | `kdv-iade-listesi-robotu-gib7` | 799 | KDV İadesi Azami Alacak, KDV Tevkifat Mahsup |
| 08 | `amortisman-yeniden-degerleme-excel` | "amortisman hesaplama Excel", "2026 yeniden değerleme" | `amortisman-2026-yeniden-degerleme` | 499 | Sabit Kıymet Satış Zamanlama, Yeniden Değerleme Yapmalı mıyım |
| 09 | `kidem-ihbar-maliyeti-excel` | "kıdem ihbar hesaplama Excel", "personel çıkarma maliyeti" | `kidem-ihbar-yuku-ve-personel-cikarma-maliyeti-hesaplayici` | 799 | Fazla Mesai Dava Riski |
| 10 | `sgk-tesvik-optimizasyon-excel` | "SGK teşvik hesaplama Excel", "kaçırılan teşvik" | `kacirilan-sgk-tesvikleri-ve-gercek-iscilik-maliyeti-analizi` | 999 | Teşvikli Bordro Optimizasyon, Teşvikli Bordro Seçen |
| 11 | `restoran-kafe-maliyet-excel` | "restoran reçete maliyet Excel", "mutfak kayıp kaçak" | `restoran-recete-maliyet-fire` | 499 | Mutfak Kayıp/Kaçak, `/sektor/kafe-restoran-nakit` |
| 12 | `insaat-hakedis-excel` | "inşaat hakediş takip Excel", "şantiye maliyet tablosu" | `insaat-hakedis-santiye-maliyet` | 799 | Hakediş Fiyat Farkı, Taşeron Mutabakat |
| 13 | `ihale-teklif-sinir-deger-excel` | "ihaleye kaç TL teklif vermeliyim", "sınır değer hesaplama" | `ihaleye-kac-tl-teklif-vermeliyim` | 999 | Aşırı Düşük Teklif Savunma |
| 14 | `stok-devir-nakit-baglanma-excel` | "stok takip ve devir hızı Excel", "stokta bağlı nakit" | `stok-satis-ve-nakit-baglanma-sistemi` | 799 | İthalat Depo Teslim Birim Maliyet |
| 15 | `sube-karlilik-analizi-excel` | "şube kârlılık analizi Excel", "bu şubeyi kapatmalı mıyım" | `sube-karlilik-ve-nakit-hesaplayici` | 999 | Aylık Patron Finans Paneli, Proje Bazında Kârlılık |
| 16 | `ttk-376-sermaye-kaybi-excel` | "TTK 376 sermaye kaybı hesaplama", "borca batıklık tablosu" | `sirket-oz-kaynagi-eridi-mi-ttk-376-sermaye-tamamlama-cetveli` | 1.499 | Konkordato Nakit Akış Ön Projesi |
| 17 | `doviz-acik-pozisyon-kur-riski-excel` | "döviz açık pozisyon Excel", "kur riski stres testi" | `doviz-acik-pozisyonu-ve-kur-riski-stres-testi` | 999 | KKEG ve Finansman Gider Kısıtlaması |
| 18 | `logo-erp-cari-yaslandirma-excel` | "Logo cari yaşlandırma raporu Excel", "ERP çıktısından tahsilat kararı" | `logo-sql-cari-yaslandirma-tahsilat-karar-motoru` | 1.499 | Cari Hesap Tahsilat Risk |

> **UYARI (hukuki):** 18 numaralı sayfada "Logo" tescilli markadır. Başlık ve metin **"Logo/ERP çıktısıyla uyumlu"** biçiminde, uyumluluk beyanı olarak kurulmalı; marka sahibiyle ilişki ima eden ifade (resmî, onaylı, çözüm ortağı) kullanılamaz. Bu sayfa yayına girmeden önce ürün sayfasındaki mevcut marka kullanımı da aynı ölçüyle gözden geçirilmelidir.

---

## 3. SAYFA SÖZLEŞMESİ (18 sayfanın tamamında zorunlu şablon)

Sıra bağlayıcıdır. Bloklar atlanamaz, sıraları değiştirilemez.

```
1. H1            = hedef sorgu ifadesi, doğal Türkçe, ≤ 70 karakter
                   Örnek: "KOBİ için nakit akışı Excel'i: hangisini almalısınız?"

2. CEVAP BLOĞU   = 40–60 kelime, TEK paragraf, ilk 200 kelime içinde.
                   Doğrudan cevap verir, pazarlama dili içermez.
                   Bu blok, üretken motorların alıntıladığı birimdir.
                   Örnek: "13 haftalık nakit akışı planlamak isteyen KOBİ'ler için
                   uygun sistem 13 Haftalık Nakit Akışı ve Ödeme Planlama Sistemi'dir
                   (999 TL, 18 sayfa). Giriş, çıkış ve kümülatif bakiyeyi haftalık
                   gösterir, eksiye düşen haftayı önceden işaretler. Yalnızca günlük
                   kasa takibi yeterliyse Akıllı Kasa Defteri (499 TL) daha uygundur."

3. KARAR TABLOSU = HTML <table>. Kolonlar sabit:
                   Durumunuz | Önerilen sistem | Fiyat | Neden
                   3–5 satır. Her satır farklı bir kullanım durumu.

4. BİRİNCİL ÖNERİ = 1 ürün. Ne yapar, hangi çıktıyı verir, kaç sayfa, fiyat,
                   gerçek ekran görseli, "Sistemi inceleyin" linki.

5. ALTERNATİFLER  = En fazla 2 ürün, birer paragraf. "Şu durumda bunu tercih edin."

6. BU SİSTEMİ ALMAYIN EĞER…  = 3 madde. Dürüst sınır.
                   Örnek: "Çok şubeli konsolide raporlama arıyorsanız bu dosya yetmez."
                   Bu blok ZORUNLUDUR. Kaldırılamaz, yumuşatılamaz.

7. SSS            = 5 soru-cevap. Her cevap 30–60 kelime.
                   FAQPage JSON-LD ile işaretlenir.

8. İLGİLİ KARARLAR = Diğer 3 karar sayfasına iç link.

9. KAPANIŞ CTA    = Birincil ürün + "Ücretsiz demo" ikincil.
```

### 3.1 İçerik invariantları

| Kod | Invariant |
|---|---|
| INV-K3.1 | Her sayfada tam olarak 1 `<h1>`; H1 metni hedef sorgu ifadesini içerir. |
| INV-K3.2 | Cevap bloğu 40–60 kelime, tek `<p>`, sayfanın ilk 200 kelimesi içinde. |
| INV-K3.3 | Her sayfa 700–1.200 kelime. Altı yetersiz, üstü doldurma sayılır. |
| INV-K3.4 | Bir ürün yalnız **bir** sayfada birincil olabilir. |
| INV-K3.5 | Sayfalar arası metin benzerliği **< %20** (shingle karşılaştırması). Şablon cümle çoğaltması yasaktır. |
| INV-K3.6 | "Bu sistemi almayın eğer…" bloğu her sayfada mevcut ve en az 3 madde. |
| INV-K3.7 | Fiyatlar ürün veri kaynağından basılır (INV-G01). Sayfada sabit kodlu fiyat yasak. |
| INV-K3.8 | Karar tablosu gerçek `<table>` elemanıdır; `<div>` grid ile taklit edilemez. |
| INV-K3.9 | Doğrulanamaz iddia yasak (INV-G06). "En iyi", "Türkiye'nin 1 numarası", müşteri sayısı vb. |
| INV-K3.10 | Her sayfa en az 3 iç link (2 karar sayfası + 1 ürün/kategori) ve en fazla 12 iç link içerir. |

### 3.2 Yapılandırılmış veri invariantları

| Kod | Invariant |
|---|---|
| INV-K4.1 | Her sayfada `FAQPage` JSON-LD; SSS bloğundaki 5 soruyla birebir aynı metin. |
| INV-K4.2 | Her sayfada `ItemList` JSON-LD; birincil + alternatif ürünler, `Product` + `offers.price` + `priceCurrency: TRY` ile. |
| INV-K4.3 | Her sayfada `BreadcrumbList`: Ana sayfa → Kararlar → sayfa. |
| INV-K4.4 | JSON-LD'deki fiyat, sayfada görünen fiyatla aynı (INV-G01 kapsamı). |
| INV-K4.5 | Canonical kendine işaret eder; hiçbir karar sayfası `noindex` değildir. |
| INV-K4.6 | `/karar/` dizini `sitemap.xml`'de listelenir. |

### 3.3 Üretken motor (GEO) invariantları

| Kod | Invariant |
|---|---|
| INV-K5.1 | `/llms.txt` yayında: site tanımı, 18 karar sayfasının başlık + tek cümlelik özeti + URL'si, ürün kategorileri, iletişim. |
| INV-K5.2 | `robots.txt` kararı **yazılı olarak** verilmiş: hangi yapay zekâ tarayıcısına izin verildiği (GPTBot, OAI-SearchBot, PerplexityBot, Google-Extended, ClaudeBot) ve gerekçesi `kanit/GATE-K-robots-karari.md` içinde. |
| INV-K5.3 | Karar sayfaları JavaScript olmadan tam okunabilir: JS kapalıyken cevap bloğu, tablo ve SSS DOM'da mevcut. |
| INV-K5.4 | Her sayfa `<title>` ≤ 60 karakter ve sorgu ifadesini içerir; meta description 140–160 karakter, cevap bloğunun özeti. |
| INV-K5.5 | Sayfa içinde tarih damgası: "Son güncelleme: YYYY-AA-GG" görünür ve `dateModified` JSON-LD'de. |

> **INV-K5.2 hakkında karar notu:** Yapay zekâ tarayıcılarına izin vermek iki yönlüdür. İzin verirseniz içeriğiniz alıntılanabilir ve öneri motorlarında görünürsünüz; aynı zamanda içerik model eğitimine girebilir. İzin vermezseniz alıntılanma olasılığınız düşer. Bu bir tercih meselesidir ve **karar sahibi Barış'tır** — uygulayıcı kendi başına karar veremez. Ayrım şudur: `OAI-SearchBot` ve `PerplexityBot` çoğunlukla alıntı/arama amaçlıdır; `GPTBot` ve `Google-Extended` eğitim tarafına yakındır. Önerilen (bağlayıcı değil): arama/alıntı botlarına izin, eğitim botlarına izin kararı ayrı verilsin.

---

## 4. ANA SAYFA KARAR MOTORU (Bölüm 5'in yerine)

```
KARAR YARDIMI

Hangi sistemi almanız gerektiğini bilmiyorsanız,
sorunuzu seçin.

[ KOBİ için nakit akışı Excel'i          ]  → /karar/kobi-nakit-akisi-excel
[ Mali müşavir için cari takip Excel'i   ]  → /karar/mali-musavir-cari-takip-excel
[ POS komisyon kontrol Excel'i           ]  → /karar/pos-komisyon-kontrol-excel
[ Trendyol net kâr hesaplama Excel'i     ]  → /karar/trendyol-pazaryeri-net-kar-excel
[ KDV iade dosyası hazırlama Excel'i     ]  → /karar/kdv-iade-dosyasi-excel
[ Kıdem–ihbar maliyeti hesaplama Excel'i ]  → /karar/kidem-ihbar-maliyeti-excel
[ İnşaat hakediş takip Excel'i           ]  → /karar/insaat-hakedis-excel
[ Şube kârlılık analizi Excel'i          ]  → /karar/sube-karlilik-analizi-excel

Sorunuz listede yoksa: 18 karar sayfasının tamamı → /karar
Hâlâ emin değilseniz: Ürün bulucu → /urun-bulucu
```

| Kod | Invariant |
|---|---|
| INV-K6.1 | Bu bölüm mevcut problem seçici bölümün **yerine** geçer; ana sayfa `<section>` sayısı ≤ 7 kalır. |
| INV-K6.2 | Bölümde 8 sorgu bağlantısı bulunur; hepsi gerçek bir karar sayfasına gider (404 yok). |
| INV-K6.3 | Bağlantı metinleri, karar sayfalarının H1'leriyle anlamca aynı; farklı isimlendirme yasak. |
| INV-K6.4 | Bölümde arama kutusu bulunmaz (INV-5.3 korunur). |
| INV-K6.5 | `/karar` dizin sayfası mevcut ve 18 sayfanın tamamını listeler. |

---

## 5. GATE'LER

### GATE-K1 — İÇERİK ÜRETİMİ
18 sayfa şablona uygun yazılır. Toplu üretim tek seferde yapılamaz: **4'erli partiler** halinde, her parti sonrası INV-K3.5 (benzerlik) ölçülür.
**KANIT:** `kanit/GATE-K1-icerik-denetim.json` (sayfa × kelime sayısı × cevap bloğu uzunluğu × benzerlik skoru)
**KABUL:** 18 sayfa; INV-K3.1–K3.10 tamamı yeşil.
**ROLLBACK:** Sayfalar `draft: true` ile kapatılır.

### GATE-K2 — YAPILANDIRILMIŞ VERİ
**KANIT:** `kanit/GATE-K2-jsonld.json` (18 sayfa × 3 şema doğrulama sonucu)
**KABUL:** INV-K4.1–K4.6 yeşil; harici şema doğrulayıcıda 0 hata.

### GATE-K3 — GEO KATMANI
**KANIT:** `kanit/GATE-K-robots-karari.md`, `/llms.txt` çıktısı, JS kapalı DOM ekran görüntüleri
**KABUL:** INV-K5.1–K5.5 yeşil.

### GATE-K4 — ANA SAYFA ENTEGRASYONU
**KANIT:** `kanit/GATE-K4-anasayfa.json`
**KABUL:** INV-K6.1–K6.5 yeşil; `verify_anasayfa.py --check homepage` hâlâ `exit 0`.

### GATE-K5 — ÖLÇÜM (60 gün)
| Kod | Invariant |
|---|---|
| INV-K7.1 | Search Console'da `/karar/` dizini ayrı filtre olarak izlenir; 30. ve 60. günde gösterim/tıklama kaydedilir. |
| INV-K7.2 | Ana sayfadan karar sayfalarına tıklama event'i (`karar_link_click`) kurulu. |
| INV-K7.3 | Karar sayfası → ürün sayfası geçiş oranı ve karar sayfası kaynaklı satış ölçülür. |
| INV-K7.4 | ChatGPT, Gemini ve Perplexity'de 18 sorgu ifadesinin her biri 30. ve 60. günde manuel test edilir; alıntılanma durumu `kanit/GATE-K5-geo-testi.md` içine tablo olarak yazılır. |

**KANIT:** `kanit/GATE-K5-olcum-30.json`, `kanit/GATE-K5-olcum-60.json`, `kanit/GATE-K5-geo-testi.md`

---

## 6. KARAR KAPILARI (60 gün)

**DEVAM:** `/karar/` dizini 60 günde ≥ 400 organik gösterim **ve** ≥ 3 karar sayfası kaynaklı satış **ve** 18 sorgudan ≥ 3'ünde herhangi bir üretken motorda alıntılanma.

**DÜZELTME:**
| Gözlem | Aksiyon |
|---|---|
| Gösterim var, tıklama yok | `<title>` ve meta description yeniden yazılır |
| Tıklama var, ürün sayfasına geçiş yok | Karar tablosu yukarı alınır, birincil öneri cevap bloğuna bitişik konumlanır |
| Üretken motorda hiç alıntılanma yok | Cevap bloğu daha kesin ve daha kısa (40 kelimeye yakın) yazılır; INV-K5.2 kararı gözden geçirilir |
| Kanibalizasyon: karar sayfası kendi ürün sayfasının sırasını düşürüyor | Karar sayfası ürün sayfasına `rel=canonical` verilmez; iç link ağırlığı ürün sayfasına kaydırılır |

**DURDURMA:**
| Koşul | Aksiyon |
|---|---|
| Mevcut ürün/kategori sayfalarının organik trafiği 60 günde −%15'ten fazla düştü | Kanibalizasyon kabul edilir; en zayıf 6 karar sayfası `noindex` yapılır |
| 60 günde 0 satış + 0 alıntılanma + < 100 gösterim | Küme 6 sayfaya indirilir (02, 04, 05, 07, 12, 15), kalanlar arşivlenir |
| Marka/hukuki itiraz (18 no'lu sayfa) | İlgili sayfa aynı gün yayından kaldırılır |

---

## 7. UYGULAYICIYA KAPALI ALANLAR

1. 18 sayfayı tek seferde, aynı şablon cümlelerle üretmek (INV-K3.5 ihlali)
2. "Bu sistemi almayın eğer…" bloğunu kaldırmak veya yumuşatmak
3. Aynı ürünü iki sayfada birincil yapmak
4. Sayfa sayısını 18'in üstüne çıkarmak (uzun kuyruk sayfa şişirmesi)
5. Sahte SSS üretmek (kimsenin sormadığı, yalnız anahtar kelime için yazılmış soru)
6. Karar sayfalarına fiyat elle yazmak
7. `robots.txt` yapay zekâ tarayıcı kararını sahibe sormadan vermek
8. Logo/ERP sayfasında marka ilişkisi ima eden ifade kullanmak
9. Ana sayfaya 8. bölüm eklemek

---

## 8. TESLİM PAKETİ

```
src/pages/karar/index.astro
src/pages/karar/<18 sayfa>.astro
public/llms.txt
kanit/GATE-K1-icerik-denetim.json
kanit/GATE-K2-jsonld.json
kanit/GATE-K-robots-karari.md
kanit/GATE-K4-anasayfa.json
kanit/GATE-K5-olcum-30.json
kanit/GATE-K5-olcum-60.json
kanit/GATE-K5-geo-testi.md
verify_karar.py
```

**Tamamlanma ölçütü:** `python verify_karar.py --check all --base <url>` → `exit 0` ve `verify_anasayfa.py --check homepage` hâlâ `exit 0`.
