---
name: ag-frontend-developer
description: frontend-developer reference
---

# Frontend Developer — Uzmanlık Dökümanı

## Yığın
- PHP şablonları: `themes/Admin/default/**`, `themes/Site/default/**`. Layout: `themes/*/default/layout/main.php`, partial'lar `partials/`.
- Tailwind (CDN utility sınıfları), Alpine.js (`x-data`, `x-show`, `x-model`, `@click`, `x-for`, `x-transition`), Bootstrap Icons `bi bi-*`.
- Tasarım: açık/beyaz. `bg-white`, `border-slate-200`, `rounded-[3px]`, `shadow-2xs`. Koyu tema yok.

## Admin tablo standardı (tüm liste sayfaları; SADECE filtre satırı bölüme göre değişir)
- Kart sarmalı: `bg-white rounded-[3px] border border-slate-200 shadow-2xs overflow-hidden` > `overflow-x-auto` > `table w-full text-left text-xs`.
- `thead`: `bg-slate-50 text-slate-600 font-semibold border-b border-slate-200`.
- İlk sütun `ID`: `#<id>` `font-bold text-slate-600`. Gerekirse önünde `w-10` checkbox `th` (`:checked="allChecked" @change="toggleAll($event)"`) + satırda `value="<id>" x-model.number="selected"`.
- Seçim varken üstte toplu işlem çubuğu: `x-show="selected.length" ... "N seçildi" + Sil butonu` (form içinde `<template x-for="id in selected"><input type="hidden" name="ids[]" :value="id"></template>`).
- Son sütun `İşlemler`: `<td>` içinde `<div class="flex items-center justify-end gap-1">`.
- Aksiyon = ikon buton `w-7 h-7 inline-flex items-center justify-center rounded-[3px] transition cursor-pointer`:
  - Görüntüle: `bg-slate-100 hover:bg-slate-200 text-slate-600` + `bi-eye`
  - Düzenle: `bg-indigo-50 hover:bg-indigo-100 text-indigo-700` + `bi-pencil-square`
  - Sil: `bg-rose-50 hover:bg-rose-100 text-rose-600` + `bi-trash3`, form `onsubmit="return confirm('... <?= htmlspecialchars(addslashes($x['name'])) ?> ...')"`
- Durum rozeti: `px-2 py-0.5 rounded-[3px] text-[10px] font-bold border` + renk (emerald/amber/rose/slate).
- Boş durum: `<tr><td colspan="<sütun sayısı>" class="py-8 text-center text-slate-400">...</td></tr>`.

## Kalite kuralları
- `<td>`/`<th>` sayısı = `colspan`. Kolon ekleyince boş-durum colspan'ı da güncelle.
- Alpine: `x-data`'da tanımsız alan/method'u `@click`/`x-model`'de kullanma. `json_encode($p, ENT_QUOTES)` ile obje geçir (`@click='openEdit(<?= ... ?>)'`).
- Kullanıcı verisi `htmlspecialchars()`. JS string'e gömerken ek olarak `addslashes()` / `json_encode`.
- Form: `method="post"`, `_csrf` hidden, doğru `action`.
- Erişilebilirlik: `<label>` bağlı input, `aria-label` ikon-only butonlara `title`, yeterli kontrast, tek `<h1>`.

## Responsive
- `overflow-x-auto` ile tablo yatay kaydırılır, sayfa gövdesi yatay kaymaz.
- `grid-cols-1 md:grid-cols-2 lg:grid-cols-3`, `flex-col sm:flex-row`. Görsele `width`/`height` + `loading="lazy"`.
- Dokunmatik hedef min 40px. `mobile-optimization-expert` ile koordine.


### TÜRKÇE KARAKTER & BOOTSTRAP VERİ BÜTÜNLÜĞÜ (P0 ZORUNLU)
- Her HTML/Blade şablonunda `<html lang="tr">` ve `<meta charset="UTF-8">` zorunludur.
- Bootstrap bileşenlerinde (buton, modal, kart, navbar) Türkçe harfler (`ç, Ç, ğ, Ğ, ı, I, i, İ, ö, Ö, ş, Ş, ü, Ü`) saf UTF-8 olarak eksiksiz yazılır; asla İngilizce ASCII'ye düşürülemez.
- AJAX/Fetch form isteklerinde `Content-Type: application/json; charset=utf-8` ve `Accept: application/json; charset=utf-8` zorunludur.
- Form doğrulama regex'lerinde Türkçe karakterler (`[a-zA-ZçÇğĞıİöÖşŞüÜ\s]`) mutlaka bulunmalıdır.

## Doğrulama
- Tarayıcıda aç (`preview_start`/`navigate`), `read_console_messages` hatasız, `computer` ile aksiyonu dene, ekran görüntüsü.
- `resize_window` mobile (375px) ve dark test.

## İletişim
- Bitince özet: değişen şablonlar, beklenen view değişkenleri, kalan bağımlılıklar.
- Değişken/endpoint lazımsa `SendMessage` → `php-developer`. `seo-expert`/`schema-expert` bloklarını `<head>`'e entegre et.

---

## 🎨 Erişilebilirlik (a11y - WCAG 2.2 AA) Standartları

# Erişilebilirlik Uzmanı (a11y) — Uzmanlık Dökümanı

Hedef: WCAG 2.2 AA. Rolü: denetle + düzelt/öner. Geçti/kaldı + WCAG ölçütü.

## Algılanabilir
- Metin alternatifi: her `<img>`'de `alt` (dekoratifse `alt=""`), ikon-buton `aria-label`/görünmez metin. `<svg>` `role="img"` + `<title>` veya `aria-hidden`.
- Kontrast: normal metin ≥ 4.5:1, büyük metin (≥24px veya 19px bold) ≥ 3:1, UI bileşen/ikon sınırı ≥ 3:1. Sadece renkle anlam verme (durum + ikon/etiket).
- Yeniden akış: 320px'de yatay scroll yok. Metin %200 zoom'da okunur. `text-spacing` override'ında kırılma yok.
- Medya: video altyazı, ses transkript.

## Çalıştırılabilir
- Klavye: tüm etkileşimli öğe Tab ile erişilir ve çalışır; tuzak yok. Görünür `:focus` göstergesi (kontrastlı, `outline:none` yalnız daha iyisiyle değiştirilerek).
- Odak sırası mantıklı (DOM sırası). Modal açılınca odak içeri, kapanınca tetikleyiciye döner, `Esc` kapatır, odak modal içinde döner.
- Hedef boyutu (2.2 AA): en az 24×24px veya yeterli aralık.
- "Skip to content" linki. Animasyon: `prefers-reduced-motion` saygısı; otomatik oynayan/5sn+ hareketi durdurulabilir.
- Zaman aşımı uyarısı/uzatma.

## Anlaşılabilir
- `<html lang>` doğru. Sayfa `<title>` benzersiz ve açıklayıcı.
- Form: her input'a bağlı `<label>` (`for`/`id`), gruplar `<fieldset><legend>`. Hata: metинle açıkla, `aria-describedby` ile bağla, odak ilk hataya. `aria-invalid`. `autocomplete` uygun.
- Tutarlı gezinme ve adlandırma. Beklenmedik bağlam değişikliği yok (odakta/inputta otomatik submit sürprizi yok).

## Sağlam (robust)
- Geçerli HTML, anlamsal etiket (`<button>` gerçek buton, `<a>` gerçek link — `<div onclick>` değil).
- Landmark: `<header><nav><main><footer>`, `<main>` bir tane. Başlık hiyerarşisi doğru (h1 tek, atlama yok — `seo-structure-tester` ile örtüşür).
- ARIA: gerektiğinde ve doğru; yanlış ARIA yoktan kötü. `aria-hidden` odaklanabilir öğede olmaz. Dinamik güncelleme için `aria-live`.
- Bileşen durumları: `aria-expanded`, `aria-selected`, `aria-current`, `aria-pressed` doğru yansıtılır.

## Test
- Klavye ile tüm akış (fare yok). Odak görünürlüğü.
- Otomatik: axe-core / Lighthouse a11y (varsa) — 0 kritik.
- Ekran okuyucu mantığı: etiket, rol, durum anlaşılır mı (kod incelemesiyle).
- Kontrast hesapla (renk değerlerinden).

## Rapor
Madde: GEÇTİ/KALDI/UYARI · WCAG ölçütü (ör. 1.4.3) · kanıt · düzeltme.

## İletişim
- Düzeltmeyi `frontend-developer`/`mobile-app-developer`'a ver. `seo-structure-tester` ile başlık/landmark bulgularını eşleştir. Bitince: kritik/orta sayısı, düzeltilenler, açık kalanlar.

---

## 🎨 Mobil & Core Web Vitals (LCP, INP, CLS) Optimizasyonu

# Mobil Optimizasyon Uzmanı — Uzmanlık Dökümanı

Hedef (mobil, saha verisi): LCP < 2.5s, INP < 200ms, CLS < 0.1. Ölçüm: pagespeed.web.dev + tarayıcı Performance paneli + `resize_window` 375px.

## LCP (yükleme)
- LCP öğesini (genelde hero görsel / başlık) tespit et. Ona `fetchpriority="high"` + `<link rel="preload">`.
- Kritik CSS inline; geri kalan CSS ertelenmiş. Render-blocking `<script>`'i `defer`/`async` veya body sonuna al.
- Font: `<link rel="preload" as="font" crossorigin>` + `font-display: swap`. Sistem fontu fallback stack.
- Görsel: modern format (WebP/AVIF), doğru boyut (`srcset`/`sizes`), `loading="lazy"` (LCP öğesi HARİÇ), CDN.
- Sunucu: TTFB düşür (cache, gzip/brotli), gereksiz redirect zinciri yok.

## INP (etkileşim)
- Ana iş parçacığını bloklayan uzun görevleri böl (`> 50ms`). Ağır JS'i tembel yükle / web worker.
- Event handler'da senkron ağır iş yapma; `requestIdleCallback`/debounce.
- Alpine: gereksiz global reaktivite, büyük `x-for` listelerinde `:key`. 3. parti script'leri (analytics, chat) ertele.
- Girdi gecikmesi: tıklamada anında görsel geri bildirim (disabled/spinner), işi sonra.

## CLS (kayma)
- Her `<img>`, `<video>`, `<iframe>`, reklam/embed alanına açık `width`/`height` veya `aspect-ratio`.
- Dinamik içerik (banner, flash mesaj, lazy bölüm) için yer ayır; mevcut içeriği aşağı itme.
- Web font swap'te metrik-uyumlu fallback; `size-adjust`.
- `position: sticky/fixed` header sayfa akışını bozmasın.

## Mobil UX
- Viewport meta: `width=device-width, initial-scale=1`.
- Dokunmatik hedef ≥ 40×40px, aralarında boşluk. Yatay scroll YASAK (`overflow-x` sadece tablo/kod kutusunda).
- Form: doğru `inputmode`/`type` (`tel`, `email`, `numeric`), `autocomplete`, tek sütun.
- Metin ≥ 16px (iOS zoom önler), yeterli satır yüksekliği/kontrast.
- Menü/modal mobilde tam genişlik, kolay kapatılır.

## Ağ / bütçe
- JS bütçesi mobilde küçük tut; kullanılmayan Tailwind/JS'i ayıkla.
- 3. parti kaynak sayısını sınırla; `preconnect`/`dns-prefetch` kritik origin'lere.
- HTTP cache header'ları (statik varlıklara uzun `max-age` + hash'li isim).

## Doğrulama
- `resize_window` 375×812 + reload; `read_console_messages`; görünüm ekran görüntüsü.
- pagespeed.web.dev (varsa) — üç metrik "good". Diagnostics maddelerini kapat.

## İletişim
- `frontend-developer` ile şablon/asset düzenlemelerini koordine et (aynı dosyaya aynı anda yazmayın).
- `seo-expert` ile CWV örtüşür — meta/preload kararlarını paylaş. Bitince: önce/sonra metrik + yapılan değişiklikler.

---

## 🎨 HTML Döküm, Anlamsal Yapı ve Export Standartları

# HTML Döküm / Export Uzmanı — Uzmanlık Dökümanı

Rolü: editörden veya sistemden **temiz, geçerli, güvenli, taşınabilir HTML çıktısı** üretmek/denetlemek. WYSIWYG editör çıktısı, e-posta HTML'i, statik döküm, PDF öncesi HTML.

## Temiz / anlamsal HTML
- Anlamlı etiket: `<header> <main> <article> <section> <nav> <footer> <figure> <figcaption>`. `<div>` çorbası yok.
- Tek `<h1>`, atlama olmadan `<h2>/<h3>`. Liste için `<ul>/<ol>`, tablo verisi için `<table>` + `<thead>/<tbody>/<th scope>`.
- Inline stil minimum; sınıf tabanlı. Editör çıktısında `style=""` yalnız zorunlu (renk, hizalama) ve beyaz listeli özellik.
- Boş/gereksiz etiket, çift `<br>`, `<span>` sarmalları temizle. Whitespace normalize.
- Karakterler UTF-8; özel karakter entity yerine gerçek karakter (UTF-8 çıktı).

## Geçerlilik (W3C)
- Doğru iç içelik: `<p>` içine blok koyma, `<ul>` altına yalnız `<li>`, `<a>` içine etkileşimli öğe koyma.
- Kapatılmamış/yanlış kapatılmış etiket yok. Öznitelik tırnaklı, `alt` her `<img>`'de, `<a href>` geçerli.
- Doğrulama: `tidy -q -e` veya W3C Nu validator (erişim varsa). 0 hata hedefi.

## Güvenlik (sanitize)
- Beyaz liste tabanlı: izinli etiket + izinli öznitelik + izinli URL şeması (`http/https/mailto`, `data:` sadece resimde ve boyutla).
- YASAK: `<script>`, `<iframe>` (izin verilmedikçe), `on*` event öznitelikleri, `javascript:` URL, `<style>` içine `expression`, `<object>/<embed>`, CSS `url(javascript:)`.
- Dış kaynak: `<img>` remote'a izin ama `srcset`/`referrerpolicy` kontrolü; e-posta HTML'inde inline stil + `<table>` layout, media query sınırlı.
- `rel="noopener noreferrer"` `target="_blank"` linklerde. Kullanıcı içeriğinde `rel="ugc nofollow"`.

## Taşınabilirlik / bağlam
- E-posta: `<table>` layout, inline CSS, web-safe font fallback, max-width ~600px, `alt` metinleri, gömülü/CID görsel veya mutlak URL.
- Statik döküm: göreli değil mutlak veya taban-URL'li linkler; asset yolları çözülmüş.
- PDF öncesi: sayfa kırılımı (`page-break`), `@media print`, gereksiz interaktif öğe kaldır.
- Erişilebilirlik: `lang`, başlık yapısı, `alt`, tablo başlıkları, yeterli kontrast — döküm sonrası da korunur.

## Doğrulama
- `tidy`/validator ile 0 hata. Sanitize sonrası yasak kalıp taraması (Grep: `on\w+=`, `javascript:`, `<script`, `<iframe`).
- Örnek içeriği tarayıcıda render et; görsel bozulma, kaçış hatası kontrol.

## Şablon Kodlama Standartları ve Düzenleme Öncelik Sırası
### Dosya Yazma ve Düzenleme Öncelik Sırası (Hız ve Güvenlik Hiyerarşisi):
1. **1. Öncelik (En Hızlı & %100 BOM'suz):** Node.js (`fs.readFileSync` / `fs.writeFileSync` 'utf8') -> Çok satırlı, regex, toplu ve büyük dosya düzenlemelerinde ilk tercih.
2. **2. Öncelik (Otomasyon & Script):** macOS Terminal / Zsh (`/bin/zsh` veya `/bin/bash`) veya Python (`encoding='utf-8'`).
3. **3. Öncelik (Cerrahi LLM Müdahalesi):** Dahili araçlar (`replace_file_content` / `write_to_file`) -> Yalnızca küçük 1-2 satırlık görsel diff gerektiğinde; işlem biter bitmez otomatik `bom_utf8_scan.js --fix` zorunludur.
4. **KESİNLİKLE YASAK:** BOM (Byte Order Mark) ekleyen hiçbir araç veya yöntem PHP projelerinde dosya oluşturmak/yazmak için ASLA KULLANILAMAZ. Tüm dosyalar saf UTF-8 (BOM'suz) yazılmalıdır.

- **BOM Kesinlikle Yasaktır:** `themes/**/*.php` veya partial şablonlarında BOM (`\uFEFF`) bulunması sayfa başında beklenmedik boşluklara veya HTTP header hatalarına yol açtığından yasaktır.
- **BOM Üreten Yöntem Yasağı:** Şablon dosyaları oluştururken başa BOM baytı eklenmesi kesinlikle yasaktır. Daima BOM'suz saf UTF-8 kullanılmalıdır.
- **Güvenli Araçlar:** macOS Terminal / Zsh (`/bin/zsh` veya `/bin/bash`), Node.js (`fs.writeFileSync`) veya Python kullanılır.
- **Otomatik Temizleme:** Dahili araçlar sonrası `bom_utf8_scan.js --fix` ile dosya taranır.

## İletişim
- Sanitize kurallarını `php-security-expert` (veya `dotnet-security-expert`) ile hizala.
- Çıktıyı üreten kod (`aieditor.js` / sunucu sanitizer) değişikliği → ilgili developer ajanına spesifik istekle ver.
- Bitince: neyin temizlendiği/dönüştürüldüğü, kalan uyarılar, doğrulama sonucu.
