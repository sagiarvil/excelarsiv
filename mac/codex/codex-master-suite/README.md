# 🔌 GEMINI MASTER SUITE — AI MASTER EKLENTİSİ

> **Paket Türü:** Antigravity / Gemini Resmi Eklentisi (Plugin)  
> **Sürüm:** 1.0.0  
> **Taşınabilirlik:** %100 Dinamik Kullanıcı Yolu (os.homedir(), $HOME, $HOME)  
> **Model Uyumluluğu:** Tüm Modeller (Gemini 3.8 Flash, 3.7 Pro, Claude 3.5/3.7 Sonnet, GPT, Cursor)  
> **Yetenek Havuzu:** 15 Uzman Ajan · 21 Çekirdek Beceri (Slash Komutu) · 39 Deterministik CLI Aracı · 12 Kural

---

## 📑 İÇİNDEKİLER

1. [Komut Kullanım Mimarisi](#-1-komut-kullanim-mimarisi)
2. [Sohbet İçi Slash Komutları (Chat Commands)](#-2-sohbet-ici-slash-komutlari-chat-commands)
3. [15 Uzman Ajanı Tetikleme Komutları (Agents)](#-3-15-uzman-ajani-tetikleme-komutlari-agents)
4. [Terminal / CLI Araç Komutları (Scripts Hub)](#-4-terminal--cli-arac-komutlari-scripts-hub)
5. [Tek Komutla Kurulum ve Tescil (CLI)](#-5-tek-komutla-kurulum-ve-tescil-cli)
6. [Otomatik Kancalar (Hooks) ve Güvenlik Muhafızları](#-6-otomatik-kancalar-hooks-ve-güvenlik-muhafizlari)
7. [Otonom Güncelleme ve Kendi Kendini Geliştirme (Self-Evolution)](#-7-otonom-güncelleme-ve-kendi-kendini-geliştirme-self-evolution)
8. [Yeni Yazılım Dili Ekleme ve Az Token Mimarisi (Language Synthesizer)](#-8-yeni-yazilim-dili-ekleme-ve-az-token-mimarisi-language-synthesizer)
9. [macOS Geliştirici Araçları Kurulum Rehberi ve Resmi Linkler](#-9-macos-geliştirici-araçlari-kurulum-rehberi-ve-resmi-linkler)




---

## 🌟 1. KOMUT KULLANIM MİMARİSİ

Bu eklenti paketindeki tum yetenek ve araclar 2 temel kanaldan çalıştırılır:

1. **Sohbet / Chat Alanı (Slash Komutları):**
   - Antigravity arayüzünde sohbet kutusundana `/<komut>` yazılarak doğrudan ilgili uzman beceri aktive edilir.
   - Örnek: `/suite`, `/ag-seo-expert`, `/ag-php-developer`, `/taste-skill`
2. **Terminal / Konsol Alanı (Deterministik CLI Yurutme):**
   - macOS Terminal (Zsh veya Bash) terminalinden sifir harici paket bağımlılığıyla çalıştırilan Node.js ve Python araçlarıdır.
   - Örnek: `node scripts/suite.js`, `node scripts/error_checker.js <hedef>`

---

## ⚡ 2. SOHBET ICI SLASH KOMUTLARI (CHAT COMMANDS)

Sohbet alaninda `/` tuşuna basıldığında veya doğrudan komut girildiğinde ilgili beceri otonom olarak devreye girer:

| Slash Komutu | Açıklama ve Görevi | Doğal Dil Tetikleyicileri | Çalışma Mekanizmasi |
|---|---|---|---|
| **`/suite`** | **Tüm Suite Kataloğunu Listeleme:** 15 Ajanı, 21 Beceriyi, 39 Aracı ve talimatlarını eksiksiz listeler. | "suite", "ajanlari listele", "becerileri göster" | `suite.js` betiğini çalıştırır, tum modellerde %100 Türkçe rapor sunar. |
| **`/ag-seo-expert`** | **18 Motorlu Engine V3.0 SEO & GEO Analizi:** Canlı URL tarar, 14KB AST bütçesini denetler, drop-in onarım hazırlar. | "seo analizi yap", "siteyi tara", "geo analizi", "18 motorlu analiz" | `audit_engine_v3.js` ve `curl` ile canli siteyi inceler, P0-P3 reçeteleri üretir. |
| **`/ag-php-developer`** | **Modern Full-Stack PHP Geliştirme:** PHP 8.3/8.4 strict types, mini-MVC mimarisi, PDO prepared statements. | "php kodu yaz", "controller ekle", "rest api yap" | `php -l` ile sözdizimi doğrular, OWASP güvenlik standartlarını uygular. |
| **`/ag-test-engineer`** | **Test Otomasyonu & QA:** Doğrulama Demir Kanunu (Iron Law), Pest, PHPUnit ve Playwright E2E testleri. | "test yaz", "testleri çalıştır", "qa kontrolü" | Test koşucularını tetikler, görev tamamlanmadan once terminal test kanıtı sunar. |
| **`/ag-database-expert`** | **Veritabanı Mimarisi & DBA:** 3NF normalizasyon, çift adımlı migrasyon, InnoDB/Postgres indeksleme. | "veritabani semasi", "sql optimize et", "migrasyon oluştur" | Veri modelleme kurallarını uygular, güvenli ACID işlem SQLleri üretir. |
| **`/ag-bug-hunter`** | **Cerrahi Kod Avcısı:** Mantık hataları, bellek sızıntıları, sınır durumlar ve güvenlik açığı taraması. | "hata bul", "bug hunt", "kod denetimi", "açık ara" | `error_checker.js` ile tum projeyi statik ve dinamik denetimden geçirir. |
| **`/ag-accessibility-expert`**| **Erişilebilirlik (a11y) Denetimi:** WCAG 2.2 AA uyumu, ARIA rolleri, klavye navigasyonu ve formlar. | "erisilebilirlik", "a11y denetimi", "wcag testi" | DOM ağacındaki landmark, kontrast ve aria etiketlerini inceler. |
| **`/ag-mobile-optimization-expert`** | **Mobil CWV ve Performans:** Core Web Vitals (LCP, CLS, INP) optimizasyonu, dokunma hedefleri. | "mobil hız", "lcp iyileştir", "cwv optimizasyonu" | `web_vitals_hints.js` çalıştırarak kritik render yollarını optimize eder. |
| **`/ag-devops-engineer`** | **DevOps & Altyapı:** Docker, CI/CD, Nginx/Apache konfigurasyonu, SSL/HTTPS ve güvenli dagitim. | "dockerfile hazırla", "nginx ayarla", "ci/cd kur" | Sunucu vhost ve konteyner yapılandırmalarını güvenli parametrelerle oluşturur. |
| **`/ag-project-manager`** | **Proje Yönetimi & Cerrahi Görev Kırılımı:** Az token cok is, Karpathy disiplini, 4/4 PASS kabul kapisi. | "görev planla", "sprint plani", "adım adım kır" | `plan_manager.js` ile `PLAN.md` üretir ve alt ajanlari koordine eder. |
| **`/ag-schema-expert`** | **Yapısal Veri (Schema.org):** JSON-LD (Product, BreadcrumbList, Organization, FAQPage). | "schema ekle", "json-ld doğrula", "yapısal veri" | `schema_validator.js` ile zengin sonuç standartlarını doğrular. |
| **`/ag-seo-structure-tester`** | **Başlık ve Semantik Yapı Denetimi:** H1-H6 hiyerarşisi, DOM etiket bütünlüğü. | "başlıkları denetle", "h1 h2 kontrolü" | `heading_semantics_check.js` ile baslik sırasını analiz eder. |
| **`/ag-content-writer`** | **SEO & E-E-A-T İçerik Yazarlığı:** Ters piramit editoryal metin, arama niyetine uygun Türkçe icerik. | "icerik yaz", "seo makalesi", "blog yazısı" | Semantik varliklar içeren, anahtar kelime doldurmasından uzak metin üretir. |
| **`/ag-technical-writer`** | **Teknik Dokümantasyon:** docs/ altında README, OpenAPI şartnameleri ve Changelog. | "dokumantasyon hazırla", "api kılavuzu yaz" | Markdown ve OpenAPI formatinda teknik şartnameler kaleme alır. |
| **`/ag-html-export-expert`** | **Temiz HTML/Bootstrap Dışa Aktarımı:** W3C geçerli, harici bağımlılıklardan arındırılmış şablonlar. | "html export", "bootstrap temizle" | DOM agacini sanitize eder ve CSS/JS varlıklarını optimize eder. |
| **`/taste-skill`** | **Anti-Slop Modern UI:** Kurumsal, özgün, şablon gibi durmayan minimalist landing page tasarımı. | "modern landing page", "şık arayüz", "özgün tasarım" | Yüksek tasarım zevkine sahip Tailwind / modern CSS arayüzleri kurar. |
| **`/systematic-debugging`** | **Sistematik Hata Ayıklama:** Kok neden odakli 4 adımlı debugging protokolü. | "neden calismiyor", "hata kaynağını bul" | Varsayım yapmadan kanıta dayalı kok neden araştırması yapar. |
| **`/doubt-driven-development`**| **Karşıt Doğrulama (Adversarial Review):** Kritik kararlarda savunma amaçlı çapraz denetim. | "karari sorgula", "risk analizi yap" | Onerilen mimarinin zayif noktalarini test ederek sağlamlaştırır. |
| **`/code-simplification`** | **Kod Sadeleştirme:** Davranışı değiştirmeden bilişsel yuku ve karmaşıkligi azaltma. | "kodu sadelestir", "refactor yap" | Gereksiz soyutlamalari ve ölü kodları cerrahi olarak temizler. |
| **`/antigravity-account-switch`** | **Çoklu Hesap Yönetimi:** Antigravity oturum çerezlerini guvenle değiştirme. | "hesap degistir", "diğer hesaba geç" | `switch.sh` motoruyla oturumu kapatip hedef hesabi baslatir. |
| **`/ag-standards`** | **Ortak Standartlar:** Karpathy cerrahi kurallari, sifir dolgu ve BOM-suz UTF-8 denetimi. | "standartlari uygula", "cerrahi kural" | Az token - cok is prensibini ve dosya bütünlüğünu güvenceye alır. |

---

## 🤖 3. 15 UZMAN AJANI TETİKLEME KOMUTLARI (AGENTS)

Ajanlar, karmaşık görevlerde `invoke_subagent` aracılığıyla arka planda eszamanli çalıştırilabilir:

```javascript
// Örnek: SEO Expert alt ajanini göreve başlatma
invoke_subagent({
  Subagents: [
    {
      TypeName: "seo-expert",
      Role: "Teknik SEO Stratejisti",
      Prompt: "https://ornek.com adresini 18 Motorlu Engine V3 ile analiz et ve P0-P3 raporunu sun."
    }
  ]
})
```

### 15 Uzman Ajan Rol Kataloğu:
1. `seo-expert` : 18-Engine SEO/GEO/AEO Stratejisti
2. `php-developer` : Modern Full-Stack PHP Geliştirici
3. `backend-developer` : REST API & Veritabanı Mimarisi Geliştirici
4. `frontend-developer` : UI & Semantik Arayuz Geliştirici
5. `database-expert` : Veritabanı Mimarı & DBA
6. `test-engineer` : Test Otomasyon Mühendisi (QA)
7. `bug-hunter` : Mantık & Güvenlik Hata Avcısı
8. `devops-engineer` : CI/CD & Altyapı Mühendisi
9. `project-manager` : Proje Yöneticisi & Cerrahi Denetçi
10. `accessibility-expert` : Erişilebilirlik (a11y) Uzmani
11. `mobile-optimization-expert` : Mobil Performans ve CWV Uzmani
12. `html-export-expert` : HTML & Bootstrap Export Uzmani
13. `php-security-expert` : PHP Güvenlik & OWASP Uzmani
14. `content-writer` : SEO & E-E-A-T İçerik Yazari
15. `technical-writer` : Teknik Dokümantasyon Yazari

---

## 🛠️ 4. TERMINAL / CLI ARAÇ KOMUTLARI (SCRIPTS HUB)

Eklentinin `scripts/` klasöründeki tum araclar sifir harici paket bağımlılığıyla (`zero-dependency`) doğrudan calisir:

### 1. Suite Kataloğu ve Durum Raporu
```bash
node scripts/suite.js
```

### 2. Kapsamlı Hata ve Sözdizim Kontrölü (Fail-Closed)
```bash
# Temel tarama
node scripts/error_checker.js <hedef_dizin_veya_dosya>

# Kati mod (tum uyarilari hata sayar)
node scripts/error_checker.js <hedef_dizin> --strict

# Otomatik BOM onarımli ve JSON çıktılı tarama
node scripts/error_checker.js <hedef_dizin> --fix-bom --json
```

### 3. BOM Taraması ve Otomatik Temizlik
```bash
# Dizin veya dosyadaki UTF-8 BOM baytlarını (0xEF, 0xBB, 0xBF) aninda temizler
node scripts/bom_utf8_scan.js --fix <hedef_dizin_veya_dosya>
```

### 4. 18-Motorlu Engine V3.0 SEO & GEO Analiz Motoru
```bash
# Canlı web sitesini veya yerel HTML dosyasını tarar
node scripts/audit_engine_v3.js https://ornek.com
node scripts/audit_engine_v3.js ./index.html --json
```

### 5. Schema.org / JSON-LD Doğrulayıcı
```bash
node scripts/schema_validator.js https://ornek.com
node scripts/schema_validator.js ./sayfa.html
```

### 6. Başlık Hiyerarsisi (H1-H6) Semantik Testi
```bash
node scripts/heading_semantics_check.js https://ornek.com
```

### 7. Core Web Vitals ve Performans İpuçları
```bash
node scripts/web_vitals_hints.js https://ornek.com
```

### 8. Dahili Kırık Bağlantı ve Anchor Metni Denetimi
```bash
node scripts/internal_link_check.js https://ornek.com
```

### 9. robots.txt ve sitemap.xml İndekslenebilirlik Kontrölü
```bash
node scripts/sitemap_robots_check.js https://ornek.com
```

### 10. Proje Hafıza ve Kural Eşleme Motoru
```bash
node scripts/connect_projects.js
node scripts/ai_project_resolver.js
```

### 11. Az Token Cok Is Plan Yöneticisi
```bash
node scripts/plan_manager.js create <proje_adi> "Görev Başlığı"
node scripts/plan_manager.js view <proje_adi>
```

### 12. Geçici Test Sunucusuyla Otomatik Test Çalıştırma
```bash
python scripts/with_server.py "php -S localhost:8000" "node scripts/error_checker.js http://localhost:8000"
```

---

## 💻 5. TEK KOMUTLA KURULUM VE TESCİL (CLI)

Eklenti paketini herhangi bir bilgisayarda sisteme kurup tescil etmek için:

```bash
# Kurulum klasörüne gidin ve çalıştırin
node install_plugins.js
```

Kurulumu Antigravity resmi CLI üzerinden teyit etmek için:

```bash
agy plugin list
```

---

## 🛡️ 6. OTOMATIK KANCALAR (HOOKS) VE GÜVENLİK MUHAFIZLARI

Eklenti sisteme kurulduğunda `hooks.json` üzerinden arka planda 2 kanca otonom calisir:
1. **PreToolUse Güvenlik Kapisi (`scripts/pre_tool_guard.js`):** `.env`, `schema.sql`, `composer.lock` gibi kritik dosyaların kazara silinmesini veya bozulmasını engeller.
2. **PostToolUse BOM Muhafizi (`scripts/ensure_utf8_nobom.js`):** Yapılan her dosya yazma işleminden hemen sonra devreye girerek UTF-8 BOM baytlarını otomatik siler.

---

## 🔄 7. OTONOM GÜNCELLEME VE KENDİ KENDİNİ GELİŞTİRME (SELF-EVOLUTION)

Bu eklenti mimarisi, statik bir paket değildir; **canlı projeler sırasında eksiklerini tamamlayan, yeni yetenekler kazanan ve kendi kendini geliştiren (Self-Evolving)** otonom bir yapıya sahiptir.

### 🎯 7.1. Proje Sırasında Eklenti Geliştirme İlkesi
Geliştirici veya çalışan Ajan, bir proje (örneğin e-ticaret, SEO analizi, API geliştirme) yürütürken:
1. **Eksik Bir Araç veya Kural Fark Ettiğinde:** Eklentinin `scripts/`, `rules/`, `skills/` veya `agents/` klasörüne yeni aracı yazar veya mevcut olanı onarır.
2. **Sürüm Değişimi Zorunluluğu:** Eklentiye yapılan her müdahale, düzeltme veya yeni özellik eklendiğinde sürüm numarası (`version`) otomatik olarak değişir.
3. **Senkronizasyon:** Proje ortamındaki veya Masaüstündeki değişiklikler sisteme anında tescil edilir.

### 🛠️ 7.2. Güncelleme ve Versiyon Motoru (`scripts/self_updater.js`)

Eklenti dizininde veya proje terminalinde tek bir komutla çalıştırılır:

#### A) Hata Düzeltmesi veya Kural İyileştirmesi (Patch: `1.0.0` ➔ `1.0.1`):
```bash
node scripts/self_updater.js --patch "Kural dosyasındaki SQL injection denetimi güçlendirildi"
```

#### B) Yeni Yetenek, Araç veya Beceri Ekleme (Minor: `1.0.0` ➔ `1.1.0`):
```bash
node scripts/self_updater.js --minor "Yeni Webhook test aracı ve Pest test becerisi eklendi"
```

#### C) Kapsamlı Mimari Değişim (Major: `1.0.0` ➔ `2.0.0`):
```bash
node scripts/self_updater.js --major "Çoklu model orkestrasyon motoru baştan yapılandırıldı"
```

#### D) Çift Yönlü Canlı Senkronizasyon (Sürüm Değiştirmeden Eşitleme):
```bash
node scripts/self_updater.js --sync
```

#### E) Mevcut Durum ve Sağlık Kontrolü:
```bash
node scripts/self_updater.js --status
```

### 📋 7.3. Güncelleme Sırasında Otomatik Yürütülen İşlemler
`self_updater.js` komutu tetiklendiğinde arka planda deterministik olarak şu adımlar gerçekleşir:
1. **SemVer Versiyon Artışı:** `plugin.json` dosyasındaki sürüm (patch/minor/major) artırılır.
2. **Yetenek Sayımı:** Ajan, beceri, kural ve betik sayıları taranarak manifestteki `capabilities` bloğu güncellenir.
3. **Versiyon Mührü:** `installed_version.json` güncel sürüm ve zaman damgasıyla yenilenir.
4. **Tarihçeli Değişiklik Günlüğü (`CHANGELOG.md`):** Tarih, saat, yeni sürüm, yapılan değişiklik ve yazar bilgisi otomatik olarak günlüğe eklenir.
5. **Çift Yönlü Dosya Senkronizasyonu:** Masaüstündeki kaynak klasör (`Desktop/plugins/...`) ile Antigravity'nin kurulu dizini (`~/.gemini/config/plugins/...`) anında eşitlenir.
6. **BOM Koruması:** `bom_utf8_scan.js --fix` çalıştırılarak tüm dosyaların BOM-suz saf UTF-8 olduğu doğrulanır.
7. **Antigravity CLI Tescili:** `agy plugin install` otomatik çalıştırılarak Antigravity runtime'ına yeni eklenti sürümü bildirilir.

### 🤖 7.4. Yapay Zeka Ajanları İçin Canlı Güncelleme Protokolü
Ajan proje esnasında eklentide bir eksiklik giderdiğinde şu adımları izler:
1. İlgili dosyayı (`scripts/`, `skills/` vb.) yazar veya düzeltir.
2. `node scripts/self_updater.js --patch "<düzeltme_özeti>"` veya `--minor` komutunu çalıştırır.
3. `agy plugin list` ile yeni sürümün tescil edildiğini teyit eder ve projesine devam eder.

---

## 🧬 8. YENİ YAZILIM DİLİ EKLEME VE AZ TOKEN MİMARİSİ (LANGUAGE SYNTHESIZER)

Bu Master Suite eklentisinin temel varlık amacı; **hangi yazılım dili eklenirse eklensin, en az token tüketimiyle en stabil, en güvenli ve en sağlıklı kod mimarisini otonom olarak inşa etmektir.**

### 🎯 8.1. Çift Dilli Token Tasarrufu Prensibi (Bilingual Token-Efficiency)
- **Neden Ajana İngilizce Anlatılıyor?**  
  Yapay zeka modellerinin (Gemini, Claude, GPT) muhakeme yeteneği, teknik İngilizce yönergelerde matematiksel olarak en yüksek başarıyı gösterir. İngilizce teknik direktifler, Türkçe promptlara kıyasla **%40-%60 oranında daha az token harcar**, modelin bağlam penceresini (context) şişirmez ve derleyici/linter kurallarına sıfır sapmayla uymasını sağlar.
- **Kullanıcıya %100 Türkçe İletişim Güvencesi:**  
  Kullanıcı Türkçe talimat verse ve eklentiden Türkçe yanıt beklese dahi, ajanın beyni (rules, skills, agents) öz ve yoğun İngilizce ile çalışır; kullanıcıya sunulan tüm chat yanıtları, açıklamalar, analizler ve araç başlıkları (`toolAction`, `toolSummary`) **istisnasız %100 Türkçe** üretilir.

### 🛠️ 8.2. Tek Komutla Yeni Dil Ekleme (`scripts/language_synthesizer.js`)

Terminalden tek bir komut verilerek eklentiye saniyeler içinde yeni bir dil uzmanlığı kazandırılır:

```bash
# Örnek: Python ekosistemini entegre et
node scripts/language_synthesizer.js python

# Örnek: TypeScript & Modern Node.js ekosistemini entegre et
node scripts/language_synthesizer.js typescript

# Örnek: Go / Golang sistem mimarisini entegre et
node scripts/language_synthesizer.js golang

# Örnek: Rust sistem mühendisliğini entegre et
node scripts/language_synthesizer.js rust

# Örnek: Bilinmeyen herhangi bir dili otomatik kurallarla sentezle
node scripts/language_synthesizer.js kotlin
```

### 📋 8.3. Sentezleme Motorunun Otomatik Türettiği 3 Katman
Komut çalıştırıldığında sistem deterministik olarak şu 3 dosyayı üretir:
1. **Katı Kurallar (`rules/<lang>-rules.md`):**  
   Strict typing, bellek güvenliği, linter standartları, OWASP güvenlik açığı engelleri ve 14KB AST bütçesi.
2. **Çekirdek Beceri (`skills/ag-<lang>-developer/SKILL.md`):**  
   Adım adım cerrahi düzenleme, derleyici/yorumlayıcı kontrolleri ve iki aşamalı doğrulama kancaları. En tepede `%100 Türkçe kullanıcı çıktısı zorunluluğu` yer alır.
3. **Uzman Ajan (`agents/<lang>-developer/agent.md`):**  
   Az token harcayan, doğrudan hedefe odaklı (Andrej Karpathy disiplini), ilgili beceriyi yöneten uzman ajan profili.

### 🔄 8.4. Otomatik Sürüm Değişimi ve Canlı Tescil
Dil sentezleme tamamlandığı anda:
1. `scripts/self_updater.js` otomatik tetiklenir.
2. `plugin.json` içerisindeki sürüm minor olarak artırılır (örn: `1.2.0` ➔ `1.3.0`).
3. `installed_version.json` ve `CHANGELOG.md` güncellenir.
4. Masaüstü kaynak repo ile Antigravity kurulu dizini çift yönlü eşitlenir.
5. `agy plugin install` ile yeni ajan ve beceri sisteme anında tescil edilir.

---

## 💻 9. macOS GELİŞTİRİCİ ARAÇLARI KURULUM REHBERİ VE RESMİ LİNKLER

> [!IMPORTANT]
> **MACOS GELİŞTİRİCİLERİ İÇİN:**  
> Antigravity Master Suite, macOS sisteminizde kurulu olan geliştirici araçlarını (Homebrew, Laravel Herd for Mac, DBngin, MAMP veya yerel binary) otomatik keşfeder. Sisteminizde bu araçlar henüz kurulu değilse, **aşağıdaki resmi macOS kaynaklarından kurmanız şiddetle tavsiye edilir.**

### 📋 Resmi macOS İndirme & Terminal Kurulum Tablosu

| Araç | Tavsiye Edilen Sürüm | Resmi macOS İndirme Bağlantısı | Hızlı Terminal (Homebrew) Kurulum Komutu |
|---|---|---|---|
| **Node.js & NPM** | v20+ / v22 LTS | [nodejs.org](https://nodejs.org) | `brew install node` |
| **Python** | Python 3.12+ / 3.14 | [python.org/downloads/macos](https://www.python.org/downloads/macos/) | `brew install python` |
| **Git** | macOS Yerel / Homebrew | [git-scm.com](https://git-scm.com) | `brew install git` veya `xcode-select --install` |
| **PHP** | PHP 8.3 / PHP 8.4 | [herd.laravel.com](https://herd.laravel.com) | `brew install php` |
| **Composer** | Composer 2.x (En Son) | [getcomposer.org](https://getcomposer.org) | `brew install composer` |
| **Veritabanı** | MySQL / Postgres / Redis | [dbngin.com](https://dbngin.com) | `brew install mysql postgresql redis` |

### 🌟 Şiddetle Tavsiye Edilen Hepsi Bir Arada macOS Stack: Laravel Herd & DBngin
Eğer PHP, Composer, MySQL, Nginx ve Redis araçlarını tek tek derlemek veya yönetmekle uğraşmak istemiyorsanız, macOS için geliştirilmiş en modern ikiliyi kullanabilirsiniz:
- **Laravel Herd macOS:** [herd.laravel.com](https://herd.laravel.com) — PHP 8.3/8.4, Composer, Nginx ve otomatik yerel SSL (.test domainler) sıfır konfigürasyonla saniyeler içinde çalışır.
- **DBngin macOS:** [dbngin.com](https://dbngin.com) — MySQL 8.4, PostgreSQL ve Redis servislerini tek tıkla açıp kapatan minimalist ve ultra hızlı macOS uygulaması.

### ⚡ Otomatik Araç Kurulumu ve İndirme Motoru
Eğer bu araçları web sitelerinden tek tek indirmek yerine doğrudan Terminal üzerinden resmi Homebrew kaynaklarından kurmak isterseniz:

```bash
node scripts/tool_installer.js
```

Bu araç:
1. Eksik araçları tespit eder.
2. Homebrew aracılığıyla Node.js, Python, Git, PHP ve Composer'ı macOS üzerine yerel olarak kurar.
3. Çalışma zamanı yollarını `~/.gemini/config/runtime_paths.json` içine işler.

---|---|---|---|
| **Node.js & NPM** | v20+ / v22 LTS / v26 | [nodejs.org/en/download](https://nodejs.org/en/download) | `brew install node` |
| **Python** | Python 3.12+ / 3.14 | [python.org/downloads/macos](https://www.python.org/downloads/macos/) | `brew install python` |
| **Git & Git Bash** | En Son macOS Git | [git-scm.com](https://git-scm.com) | `brew install git` |
| **PHP** | PHP 8.3 / PHP 8.4 (macOS) | [herd.laravel.com](https://herd.laravel.com/) | `brew install php` |
| **Composer** | Composer 2.x (En Son) | [getcomposer.org/brew install composer](https://getcomposer.org/brew install composer) | `https://getcomposer.org/composer.phar` |


### 🌟 Şiddetle Tavsiye Edilen Hepsi Bir Arada macOS Stack: Laravel Herd
Eğer PHP, Composer, MySQL, Nginx ve Redis araçlarını tek tek kurmakla uğraşmak istemiyorsanız, tüm ortamı tek bir setup ile kuran **Laravel Herd**'i indirebilirsiniz:
- **Doğrudan İndirme Linki:** [Laravel-Herd.dmg (https://herd.laravel.com)](https://betasoft.com.tr/downloads/Laravel-Herd.dmg (https://herd.laravel.com))
- **Paket İçeriği:** PHP 8.3/8.4 NTS, Composer 2.x, MySQL 8.4, Apache 2.4, Nginx 1.27, Redis, Mailpit, Mkcert SSL.

### ⚡ Otomatik Araç Kurulumu ve İndirme Motoru
Eğer bu araçları web sitelerinden tek tek indirmek yerine terminalden tek bir komutla resmi Windows kaynaklarından kurmak isterseniz:

```bash
node scripts/tool_installer.js
```

Bu araç:
1. Eksik araçları tespit eder.
2. Windows `brew` veya doğrudan indirme ile Node.js, Python, Git, PHP ve Composer'ı yapılandırır.
3. Çalışma zamanı yollarını `~/.gemini/config/runtime_paths.json` içine işler.

---
*MANDATE-CLEAN-2026 · Taşınabilir Dinamik Yollar · %100 Türkçe · BOM-suz Saf UTF-8*
