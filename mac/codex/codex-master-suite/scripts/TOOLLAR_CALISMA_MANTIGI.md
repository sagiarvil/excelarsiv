# 🛠️ ANTIGRAVITY & GEMINI — TÜM ÇALIŞAN ARAÇLARIN ÇALIŞMA MANTIĞI KILAVUZU

> **Merkezi Çalışma Konumu:**  
> `$HOME/.gemini/config/scripts/`  
> (Ayna Konum: `$HOME/.gemini\config\scripts\`)  
> **Son Güncelleme:** 2026-09-11  
> **Temel Standart:** Sıfır Harici Bağımlılık · %100 Deterministik · BOM'suz Saf UTF-8

---

## 📑 İÇİNDEKİLER

1. [Genel Mimari Prensip & Yürütücüler](#1-genel-mimari-prensip--yürütücüler)
2. [Hata Kontrol ve Statik Analiz Araçları (Bug Hunter & QA)](#2-hata-kontrol-ve-statik-analiz-araçları)
   - `error_checker.js` & `python/error_checker.py`
   - `with_server.py`
3. [BOM ve Kodlama Standartları Muhafızları](#3-bom-ve-kodlama-standartları-muhafızları)
   - `bom_utf8_scan.js` & `python/bom_utf8_scan.py`
   - `clean_bom_and_cache.js` & `python/clean_bom_and_cache.py`
   - `ensure_utf8_nobom.js` & `pre_tool_guard.js`
4. [Engine V3.0 SEO, GEO & AEO Denetim Motorları](#4-engine-v30-seo-geo--aeo-denetim-motorları)
   - `audit_engine_v3.js` & `python/audit_engine_v3.py`
   - `dom_utf8_full_test.js` & `python/dom_utf8_full_test.py`
   - `schema_validator.js` & `python/schema_validator.py`
   - `opengraph_check.js` & `python/opengraph_check.py`
   - `heading_semantics_check.js` & `python/heading_semantics_check.py`
   - `internal_link_check.js` & `python/internal_link_check.py`
   - `sitemap_robots_check.js` & `python/sitemap_robots_check.py`
   - `web_vitals_hints.js` & `python/web_vitals_hints.py`
5. [İndeksleme ve LLM Yüzey Üreticileri](#5-i̇ndeksleme-ve-llm-yüzey-üreticileri)
   - `indexnow-pusher.js` & `python/indexnow_pusher.py`
   - `llms-generator.js` & `python/llms_generator.py`
   - `seo-checker.js` & `python/seo_checker.py`
   - `seo-folder-scanner.js` & `python/seo_folder_scanner.py`
6. [Sistem Optimizasyonu, Senkronizasyon ve Oturum Araçları](#6-sistem-optimizasyonu-senkronizasyon-ve-oturum-araçları)
   - `optimize_gemini.js`
   - `sync_agents_skills.js`
   - `switch.sh`
7. [Ortak Çekirdek Kütüphaneler](#7-ortak-çekirdek-kütüphaneler)
   - `_lib.js` & `python/_lib.py`
8. [Tavsiye Edilen Çalıştırma ve Boru Hattı (Pipeline) Sırası](#8-tavsiye-edilen-boru-hattı-pipeline-sırası)

---

## 1. GENEL MİMARİ PRENSİP & YÜRÜTÜCÜLER

Bu dizindeki tüm araçlar aşağıdaki kurallara göre inşa edilmiştir:
- **Sıfır Harici Bağımlılık (Zero External Dependencies):** Hiçbir npm paketi veya pip bağımlılığı gerektirmez. Yalnızca Node.js ve Python'un dahili standart kütüphanelerini kullanır.
- **Fail-Closed:** Herhangi bir test başarısız olduğunda boru hattını bilgilendirmek için `exit code 1` döndürür. Tüm testler geçerse `exit code 0` döner.
- **Deterministik Raporlama:** Çıktılar standart `[PASS]`, `[FAIL]`, `[WARN]`, `[INFO]` etiketleriyle üretilir. Ayrıca tüm analiz araçları `--json` bayrağını destekler.
- **BOM Yasağı:** Üretilen hiçbir dosyada UTF-8 BOM baytı (`0xEF 0xBB 0xBF`) bulunamaz.

### Sistem Canonical Yürütücü Yolları
- **Node.js (Birincil):** `node`
- **Python 3.14 (İkincil):** `python3`
- **PHP 8.3:** `php`
- **macOS Terminal (Zsh):** `/bin/zsh` (Kabuk: `zsh`)

---

## 2. HATA KONTROL VE STATİK ANALİZ ARAÇLARI

### 2.1. `error_checker.js` & `python/error_checker.py` (Kapsamlı Hata ve Sözdizim Denetleyici)
- **Görevi:** Proje veya dosya genelinde çok katmanlı hata taraması, sözdizim (syntax) doğrulaması ve güvenlik açığı avı yapar.
- **Çalışma Mantığı:**
  1. Verilen hedef yolu özyinelemeli (recursive) olarak tarar. `node_modules`, `vendor`, `.git`, `storage/cache` gibi dizinleri atlar.
  2. **BOM Kontrolü:** Dosya başında `0xEF 0xBB 0xBF` varsa `FAIL` verir (`--fix-bom` ile otomatik siler).
  3. **PHP Sözdizimi:** Canonical PHP 8.3 yürütücüsü ile `php -l` çalıştırır. Herhangi bir Parse error/syntax error varsa satır numarası ve hatayı yakalar.
  4. **JavaScript Sözdizimi:** Canonical Node.js ile `node --check` çalıştırır.
  5. **Python Sözdizimi:** Python derleyicisi ile `python -m py_compile` çalıştırır.
  6. **JSON Doğrulama:** `JSON.parse` ile JSON sözdizimini ve virgül hatalarını test eder.
  7. **Güvenlik Statik Analizi (Bug Hunter):**
     - `eval()`, `shell_exec()`, `system()` gibi tehlikeli çağrıları yakalar.
     - `extract($_GET/$_POST)` ile değişken ezme açıklarını tespit eder.
     - `echo $_GET[...]` gibi doğrudan kullanıcı çıktısı basan XSS kalıplarını bulur.
     - Ham SQL birleştirmelerini (SQLi riski) arar.
     - Canlı kodda unutulmuş `var_dump()` ve `console.log()` kalıntılarını raporlar.
  8. **Kırık Include Kontrolü:** PHP dosyalarındaki yerel `require` / `include` yollarının diskte var olup olmadığını doğrular.
- **Kullanım:**
  ```bash
  node '$HOME/.gemini/config/scripts/error_checker.js' <hedef_dizin_veya_dosya> [--fix-bom] [--json] [--strict]
  ```
  Python sürümü:
  ```bash
  python3 '$HOME/.gemini/config/scripts/python\error_checker.py' <hedef>
  ```

### 2.2. `with_server.py` (Sunucu Yaşam Döngüsü ve Test Otomasyonu)
- **Görevi:** Test veya otomasyon çalıştırmadan önce arka planda bir veya daha fazla yerel sunucu (ör. PHP yerel sunucu, Node dev server, Python backend) başlatır, portun dinlemeye geçtiğini soketle doğrular, test komutunu çalıştırır ve işlem bittiğinde sunucuları temiz bir şekilde kapatır.
- **Çalışma Mantığı:**
  1. `--server` komutunu arka plan subprocess olarak başlatır.
  2. `socket.create_connection(('localhost', port))` ile portu yoklar. Port açılana kadar bekler (timeout: 30s).
  3. Port aktif olunca ana test komutunu yürütür.
  4. Test tamamlandığında (başarılı veya hatalı) `finally` bloğunda tüm sunucu süreçlerini `SIGTERM` / `SIGKILL` ile sonlandırır (port sızıntısını engeller).
- **Kullanım:**
  ```bash
  python with_server.py --server "php -S localhost:8000" --port 8000 -- python automation.py
  ```

---

## 3. BOM VE KODLAMA STANDARTLARI MUHAFIZLARI

### 3.1. `bom_utf8_scan.js` & `python/bom_utf8_scan.py` (BOM Tarayıcı ve Onarıcı)
- **Görevi:** Canonical Tool Paths #9 aracıdır. Proje altındaki tüm kaynak dosyaları tarar ve dosya başındaki UTF-8 BOM baytlarını (`0xEF 0xBB 0xBF`) bulur.
- **Çalışma Mantığı:**
  1. `.php`, `.html`, `.js`, `.css`, `.json`, `.md`, `.xml` uzantılı dosyaları okur.
  2. Dosyanın ilk 3 baytını inceler.
  3. `--fix` bayrağı verilmişse ilk 3 baytı siler ve dosyayı saf UTF-8 olarak tekrar kaydeder.
  4. Minify şüphesi kontrolü: Tek satırda 3000 bayttan büyük PHP dosyalarını uyarır.
- **Kullanım:**
  ```bash
  node '$HOME/.gemini/config/scripts/bom_utf8_scan.js' --fix <dizin_veya_dosya>
  ```

### 3.2. `clean_bom_and_cache.js` & `python/clean_bom_and_cache.py`
- **Görevi:** Önceden tanımlı yerel web projelerinde (ör. `$HOME/Sites/satis/`) hem tüm kaynak kodlardaki BOM baytlarını siler hem de `storage/cache/` dizinini temizleyerek önbelleği boşaltır.
- **Kullanım:**
  ```bash
  node clean_bom_and_cache.js
  ```

### 3.3. `ensure_utf8_nobom.js` & `pre_tool_guard.js` (Hook ve Güvenlik Muhafızları)
- **`ensure_utf8_nobom.js`:** Her dosya yazma ve düzenleme aracından sonra stdin üzerinden çağrılır; yeni yazılan dosyanın başında BOM varsa derhal kaldırır.
- **`pre_tool_guard.js`:** PreToolUse filtresi olarak çalışır. Yetkisiz ajanların `schema.sql`, `.env`, `composer.lock` gibi kritik sistem çekirdek dosyalarını bozmasını engeller.

---

## 4. ENGINE V3.0 SEO, GEO & AEO DENETİM MOTORLARI

Bu araçlar Google Alexandria, Perplexity AI, ChatGPT Search ve Claude botlarının siteleri tarama, anlama ve alıntılama algoritmalarına tam uyum sağlar.

### 4.1. `audit_engine_v3.js` & `python/audit_engine_v3.py` (18 Motorlu Tam Denetim)
- **Görevi:** Engine V3.0'ın 18 motorunu tek geçişte denetler.
- **Çalışma Mantığı:**
  - **ENG-01 & ENG-02 (14KB AST Bütçesi):** HTML dosya boyutunu ölçer. İlk TCP paketine sığması için 14.336 bayt altı olup olmadığını kontrol eder.
  - **ENG-04 (Klasik SEO Çekirdeği):** Tam 1 adet H1, 55-65 karakter Title, 70-165 karakter Meta Description ve mutlak Canonical bağlantısını denetler.
  - **ENG-06 (AEO Hero Answer):** Sayfanın ilk 100 pikselinde `class="hero-answer"` CSS sınıfına sahip 29-80 kelimelik atomik cevap bloku arar.
  - **ENG-07 & ENG-13 (LLMO & Makine Yüzeyleri):** `<link rel="describedby" href="/llms.txt">` etiketini ve makine arayüzlerini kontrol eder.
  - **ENG-09 (RAG Chunking):** AI arama motorlarının vektör sınırlarını korumak için `data-chunk-id` özniteliklerini sayar.
  - **ENG-10 (ColBERT MaxSim):** H2 ve H3 başlık dağılımını ölçer (her 1000 kelimede en az 3 alt başlık).
  - **ENG-15 (Knowledge Vault & Wikidata):** JSON-LD içerisinde Wikidata QID (`wikidata.org/wiki/Q...`) ve Google MID varlığını doğrular.
- **Kullanım:**
  ```bash
  node audit_engine_v3.js <url_veya_dosya> [--json]
  ```

### 4.2. `dom_utf8_full_test.js` & `python/dom_utf8_full_test.py` (DOM & Canlı HTTP Testi)
- **Görevi:** Canlı URL üzerinde veya yerel HTML'de HTTP başlıkları, UTF-8 karakter kümesi ve DOM yapı bütünlüğünü uçtan uca test eder.
- **Çalışma Mantığı:**
  - Canlı HTTP yanıtında `Content-Type: text/html; charset=utf-8` olup olmadığını inceler.
  - DOM'da `<meta charset="UTF-8">` etiketini teyit eder.
  - Canlı `/llms.txt`, `/robots.txt`, `/sitemap.xml` ve `/.well-known/agent-card.json` adreslerine istek atarak 200 OK yanıtlarını ve içerik formatlarını doğrular.

### 4.3. `schema_validator.js` & `python/schema_validator.py` (JSON-LD Doğrulayıcı)
- **Görevi:** Sayfadaki tüm `<script type="application/ld+json">` bloklarını çıkarır ve Schema.org kurallarına göre doğrular.
- **Çalışma Mantığı:**
  - JSON parse hatası arar.
  - `Organization`, `Product`, `Offer`, `Article`, `FAQPage`, `LocalBusiness`, `AggregateRating` tipleri için Google zorunlu ve önerilen alanlarını denetler.
  - Spam ve manipülasyon riski: Fiyatı 0 olan Offer'ları, ratingCount=0 olan sahte değerlendirmeleri ve uydurma placeholder URL'leri yakalar.

### 4.4. `opengraph_check.js` & `python/opengraph_check.py`
- **Görevi:** Sosyal medya ve arama önizlemelerini (OpenGraph & Twitter Card) denetler.
- **Kontrol:** `og:title`, `og:type`, `og:url`, `og:image`, `og:description`, `twitter:card`, `viewport` ve makale yayın/güncelleme zaman damgaları.

### 4.5. `heading_semantics_check.js` & `python/heading_semantics_check.py`
- **Görevi:** H1-H6 başlık hiyerarşisini, hiyerarşik atlamaları (ör. H2'den H4'e atlama) ve anlamsal etiketleri (`<header>`, `<nav>`, `<main>`, `<footer>`, `<html lang="tr">`) denetler.

### 4.6. `internal_link_check.js` & `python/internal_link_check.py`
- **Görevi:** Sitedeki tüm dahili bağlantıları BFS (Breadth-First Search) yöntemiyle tarar; 404 veren kırık linkleri, yönlendirme zincirlerini ve "tıklayın", "buraya" gibi jenerik/zayıf çapa metinlerini tespit eder.

### 4.7. `sitemap_robots_check.js` & `python/sitemap_robots_check.py`
- **Görevi:** `robots.txt` kurallarını (Disallow / GPTBot engeli var mı?), XML sitemap'leri (`<urlset>`, `<sitemapindex>`) ve örneklenen URL'lerin HTTP yanıt kodlarını denetler.

### 4.8. `web_vitals_hints.js` & `python/web_vitals_hints.py`
- **Görevi:** Sayfa kodunu statik analiz ederek Core Web Vitals (LCP, CLS, INP) risklerini önceden raporlar:
  - Render-blocking CSS ve `<head>` içindeki senkron JS'ler (INP/LCP riski).
  - `width` / `height` özniteliği eksik görseller (CLS riski).
  - `preconnect` ve `dns-prefetch` eksiklikleri.

---

## 5. İNDEKSLEME VE LLM YÜZEY ÜRETİCİLERİ

### 5.1. `indexnow-pusher.js` & `python/indexnow_pusher.py`
- **Görevi:** Yeni oluşturulan veya güncellenen URL'leri Bing ve Yandex IndexNow API uç noktasına doğrudan POST ederek saniyeler içinde dizine eklenmesini tetikler.
- **Kullanım:** `node indexnow-pusher.js https://ornek.com/yeni-sayfa <api_key>`

### 5.2. `llms-generator.js` & `python/llms_generator.py`
- **Görevi:** Projedeki sayfaları ve markdown dokümanlarını tarayarak standartlara uygun bir `/llms.txt` ve `/llms-full.txt` dosyası üretir.

### 5.3. `seo-checker.js` & `python/seo_checker.py`
- **Görevi:** Tek bir HTML/PHP dosyasında meta tag, başlık hiyerarşisi ve temel SEO puanını saniyeler içinde hesaplayan hızlı terminal aracıdır.

### 5.4. `seo-folder-scanner.js` & `python/seo_folder_scanner.py`
- **Görevi:** Bir klasör altındaki onlarca veya yüzlerce HTML/PHP dosyasını toplu tarayarak eksik meta, H1 veya kırık link içeren dosyaların matris raporunu çıkarır.

---

## 6. SİSTEM OPTİMİZASYONU, SENKRONİZASYON VE OTURUM ARAÇLARI

### 6.1. `optimize_gemini.js`
- **Görevi:** Masaüstündeki `Desktop\gemini\.gemini` şablonu ile sistem ana dizini `$HOME/.gemini` arasındaki agents, skills, rules ve scripts klasörlerini çift yönlü akıllı senkronize eder.

### 6.2. `sync_agents_skills.js`
- **Görevi:** 17 alt ajanın (`config/agents/`) YAML frontmatter'ındaki `skills:` referanslarını ve canonical tool yollarını tarar, 22 uzmanlık becerisiyle (`config/skills/`) eksiksiz eşleştirir.

### 6.3. `switch.sh` (Antigravity Çoklu Hesap Yöneticisi)
- **Görevi:** Google Antigravity IDE'si için oturum geçişi sağlar:
  - Aktif kullanıcı oturumunu kaydeder.
  - Hedef hesaba geçerken çerez, oturum verisi ve kota izlerini güvenli biçimde temizleyip IDE'yi yeniden başlatır.

---

## 7. ORTAK ÇEKİRDEK KÜTÜPHANELER

### 7.1. `_lib.js` & `python/_lib.py`
- Tüm Node.js ve Python analiz araçlarının temelidir.
- **Sağladığı Yetenekler:**
  - `fetchUrl()`: Bot user-agent başlıklarıyla 6 adıma kadar yönlendirmeleri izleyen, 20s zaman aşımına sahip HTTP/HTTPS istemcisi.
  - `loadHtml()`: Girdinin URL mi yoksa yerel dosya mı olduğunu algılayıp içeriği yükleyen fonksiyon.
  - `extractJsonLd()`: HTML içerisindeki tüm JSON-LD bloklarını ayrıştırıp nesneleştiren parser.
  - `makeReport()`: `[PASS]`, `[FAIL]`, `[WARN]`, `[INFO]` sayaçlarını tutan ve hem metin hem JSON formatında standart çıktı veren raporlayıcı.

---

## 8. TAVSİYE EDİLEN BORU HATTI (PIPELINE) SIRASI

Bir projede kod yazımı veya dağıtım öncesinde araçların şu sırayla çalıştırılması önerilir:

```
1. bom_utf8_scan.js --fix <proje_dizini>     → BOM baytlarını temizle
2. error_checker.js <proje_dizini>          → Sözdizim, JSON, import ve güvenlik taraması
3. audit_engine_v3.js <canli_url_veya_html> → 18 Motorlu Engine V3.0 denetimi
4. dom_utf8_full_test.js <canli_url>        → Canlı HTTP uç noktası ve makine yüzeyleri
5. schema_validator.js <canli_url>          → JSON-LD ve Schema.org doğrulama
6. opengraph_check.js <canli_url>           → OpenGraph ve meta etiketler
7. heading_semantics_check.js <dosya>       → Başlık hiyerarşisi ve anlamsal etiketler
8. internal_link_check.js <canli_url>       → Kırık link ve yönlendirme kontrolü
```

---
*Bu doküman, Desktop ve .gemini sistemlerindeki tüm araçların resmi çalışma kılavuzudur.*
