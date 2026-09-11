# Antigravity & Gemini — Çalışan Araçlar ve Scriptler Kataloğu

Bu dizin (`.gemini/config/scripts/`), Antigravity yapay zeka ajanları, SEO motoru, veri ve sistem yönetimi için geliştirilmiş tüm çalışan araçları barındırır.
Tüm araçlar sıfır harici bağımlılıkla çalışır ve hem Node.js hem de Python ile tam uyumludur.

---

## 🚀 Hızlı Başlangıç & Çalıştırma Örnekleri

Tüm Node.js araçları sistemdeki canonical Node yolu ile çalıştırılır:
```bash
node <script_yolu> [parametreler]
```

Python araçları için:
```bash
python3 <script_yolu> [parametreler]
```

---

## 🗂️ Araç Listesi ve Kategorileri

### 1. SEO, GEO & AEO Denetim Araçları (Engine V3.0 Suite)

| Araç | Açıklama | Örnek Komut |
|---|---|---|
| `audit_engine_v3.js` | 18 Motorlu Engine V3.0 kapsamlı SEO, GEO, AEO analizi | `node audit_engine_v3.js <url_veya_dosya>` |
| `dom_utf8_full_test.js` | DOM yapısı, UTF-8 charset ve canlı HTTP uç nokta testi | `node dom_utf8_full_test.js <url>` |
| `schema_validator.js` | Schema.org JSON-LD ve Google Rich Results doğrulayıcı | `node schema_validator.js <url_veya_dosya>` |
| `opengraph_check.js` | Title, Meta Description, Open Graph & Twitter Card kontrolü | `node opengraph_check.js <url_veya_dosya>` |
| `heading_semantics_check.js` | H1-H6 başlık hiyerarşisi ve anlamsal landmark denetimi | `node heading_semantics_check.js <dosya>` |
| `internal_link_check.js` | Kırık iç linkler, yönlendirme zincirleri ve anchor text denetimi | `node internal_link_check.js <url>` |
| `sitemap_robots_check.js` | robots.txt, sitemap.xml ve LLM yüzeyleri (/llms.txt) kontrolü | `node sitemap_robots_check.js <url>` |
| `web_vitals_hints.js` | Core Web Vitals (LCP, CLS, INP) statik kaynak analizörü | `node web_vitals_hints.js <url_veya_dosya>` |
| `_lib.js` | Tüm JS araçları için ortak HTTP, DOM ve raporlama kütüphanesi | (İç kütüphane) |

### 2. BOM & Kodlama Standardı Araçları

| Araç | Açıklama | Örnek Komut |
|---|---|---|
| `bom_utf8_scan.js` | Proje genelinde UTF-8 BOM taraması yapar ve `--fix` ile temizler | `node bom_utf8_scan.js --fix <dizin>` |
| `clean_bom_and_cache.js` | Belirli proje dizinindeki BOM'ları temizler ve cache sıfırlar | `node clean_bom_and_cache.js` |
| `ensure_utf8_nobom.js` | Dosya yazma/düzenleme sonrası olası BOM baytlarını otomatik siler | Hook / Pre/Post Tool |
| `ensure_utf8_bom.js` | Dosya başında BOM kontrolü sağlayan yardımcı script | Hook / Tool |

### 3. İndeksleme ve LLM Yüzey Üreticileri

| Araç | Açıklama | Örnek Komut |
|---|---|---|
| `indexnow-pusher.js` | Değişen URL'leri Bing ve Yandex'e anında push eder | `node indexnow-pusher.js <url>` |
| `llms-generator.js` | Web sitesi için standart `llms.txt` dosyasını üretir | `node llms-generator.js <dizin_veya_url>` |
| `seo-checker.js` | Hızlı terminal tabanlı genel SEO kontrolü | `node seo-checker.js <dosya>` |
| `seo-folder-scanner.js` | Bir klasördeki tüm HTML/PHP sayfalarını tarayarak SEO raporu çıkarır | `node seo-folder-scanner.js <klasor>` |

### 4. Sistem Yapılandırma, Eşleştirme ve Hesap Yönetimi

| Araç | Açıklama | Örnek Komut |
|---|---|---|
| `optimize_gemini.js` | Desktop ↔ .gemini yapılandırma senkronizasyonu ve optimizasyon | `node optimize_gemini.js` |
| `sync_agents_skills.js` | 17 ajan ile 22 skill arasındaki bağlantı ve araç yollarını eşitler | `node sync_agents_skills.js` |
| `switch.sh` | Antigravity çoklu hesap geçişi, çerez/oturum ve token yöneticisi | `./switch.sh` |
| `pre_tool_guard.js` | Kritik sistem dosyalarının yanlışlıkla silinmesini/yazılmasını önleyen güvenlik muhafızı | Hook / PreToolUse |

### 5. Python Araçları (`python/` Klasörü)

`python/` alt klasöründe aynı motorların bağımsız Python 3.14 sürümleri yer alır:
- `python/_lib.py`: Ortak kütüphane
- `python/audit_engine_v3.py`: 18 Motorlu denetim
- `python/bom_utf8_scan.py`: BOM tarama ve onarım
- `python/clean_bom_and_cache.py`: Cache ve BOM temizleyici
- `python/dom_utf8_full_test.py`: DOM ve HTTP testi
- `python/heading_semantics_check.py`: Başlık semantiği
- `python/internal_link_check.py`: İç link kontrolü
- `python/opengraph_check.py`: OpenGraph kontrolü
- `python/schema_validator.py`: JSON-LD doğrulayıcı
- `python/sitemap_robots_check.py`: Sitemap kontrolü
- `python/web_vitals_hints.py`: Web vitals analizi
- `python/indexnow_pusher.py`: IndexNow bildirici
- `python/llms_generator.py`: LLMS generator
- `python/seo_checker.py`: SEO denetleyici
- `python/seo_folder_scanner.py`: Klasör tarayıcı

---

## 📌 Sıralı CI/CD ve Test Boru Hattı

Önerilen çalışma akışı:
1. `bom_utf8_scan.js --fix <hedef>`
2. `audit_engine_v3.js <hedef>`
3. `dom_utf8_full_test.js <hedef>`
4. `schema_validator.js <hedef>`
5. `opengraph_check.js <hedef>`
6. `heading_semantics_check.js <hedef>`
7. `web_vitals_hints.js <hedef>`
