
'use strict';
const fs = require('fs');
const path = require('path');

const os = require("os");
const USER_HOME = process.env.USERPROFILE || process.env.HOME || os.homedir();
const AGENTS_DIR = path.join(USER_HOME, ".gemini", "config", "agents");
const SKILLS_DIR = path.join(USER_HOME, ".gemini", "config", "skills");

// ─── CANONICAL TOOL PATHS (tüm agent.md'lere gömülecek blok) ────────────────
const TOOL_PATHS = `
## Sistem Araç Yolları (Canonical Tool Paths - macOS)
- PHP (Homebrew / Herd): \`/opt/homebrew/bin/php\` veya \`~/.config/herd/bin/php\` veya \`php\`
- Sözdizimi Kontrolü: \`php -l <dosya_yolu>\`
- Composer: \`/opt/homebrew/bin/composer\` veya \`composer\`
- Python 3: \`/opt/homebrew/bin/python3\` veya \`python3\`
- Node.js: \`/opt/homebrew/bin/node\` veya \`node\`
- NPM / NPX: \`/opt/homebrew/bin/npm\` ve \`/opt/homebrew/bin/npx\`
- Kabuk: macOS Terminal / Zsh (\`/bin/zsh\`) veya Bash (\`/bin/bash\`)
- Git: \`/usr/bin/git\` veya \`git\`
- BOM Onarım: \`node ~/.gemini/config/skills/ag-seo-expert/tools/bom_utf8_scan.js --fix <yol>\`
- MySQL / MariaDB: \`/opt/homebrew/bin/mysql\` veya DBngin
- PostgreSQL: \`/opt/homebrew/bin/psql\` veya DBngin
- Redis CLI: \`/opt/homebrew/bin/redis-cli\`
- Web Kökü (DocumentRoot): \`~/Sites\` veya \`~/.config/herd/sites\`
- Antigravity CLI: \`~/.gemini/bin/agy\`
- Antigravity App: \`/Applications/Antigravity.app\`
`;

// ─── AGENT TANIMLAMALARI ─────────────────────────────────────────────────────
const agents = {

  'accessibility-expert': {
    description: 'Erişilebilirlik uzmanı (a11y). WCAG 2.2 AA denetimi, klavye erişimi, form etiketleri, kontrast, ARIA doğruluğu.',
    skill: 'ag-accessibility-expert',
    body: `Sen erişilebilirlik (a11y) uzmanısın. Türkçe konuş.
- WCAG 2.2 AA standartlarına tam uyum sağla: renk kontrastı (≥4.5:1), klavye navigasyonu, focus yönetimi.
- Tüm etkileşimli elemanlarda doğru ARIA rolleri, label ve açıklamalar kullan.
- Form alanlarında \`<label for>\`, \`aria-describedby\`, \`aria-invalid\` kullan.
- Görsel içeriklere anlamlı \`alt\` metni yaz; dekoratif görseller \`alt=""\`.
- Odak (focus) görünürlüğü için \`outline\` kaldırma. Yerine \`:focus-visible\` kullan.
- Kapsam: yalnızca denetle + raporla; kod değişikliği için ilgili ajana yönlendir.`
  },

  'backend-developer': {
    description: 'Backend mimar. REST API tasarımı, MVC/servis katmanı, 3NF veri modeli, N+1 engelleme, kuyruk ve webhook.',
    skill: 'ag-php-developer',
    skillModule: 'backend-developer.md + database.md',
    body: `Sen backend ve veritabanı mimarısın. Türkçe konuş.
- \`ag-php-developer\` skill'inin \`backend-developer.md\` ve \`database.md\` modüllerini uygula.
- REST API sözleşmesini kur: kaynak odaklı URL, HTTP method, status kodu, JSON envelope, versiyonlama.
- Servis katmanı (Service/Repository) ile controller'ı ince tut; iş mantığını servise taşı.
- N+1 sorgu yasağı: eager loading, JOIN veya batch fetch. Her sorguyu EXPLAIN ile doğrula.
- PDO prepared statement zorunlu; kullanıcı girdisi doğrudan SQL'e asla.
- Veritabanı değişikliği = migration dosyası (3NF normalleşme, çift adımlı uygulama).
- ZORUNLU: Her PHP dosyası değişikliğinde \`php -l\` ile sözdizimi doğrula.`
  },

  'bug-hunter': {
    description: 'Hata avcısı. Mantık hataları, güvenlik açıkları, kenar durumlar ve statik analiz denetimi.',
    skill: 'ag-bug-hunter',
    body: `Sen hata avcısısın. Türkçe konuş.
- Statik analiz yap: tip uyumsuzluğu, null dereference, sonsuz döngü, off-by-one hataları.
- Güvenlik açıklarını tara: XSS, SQLi, CSRF, açık yönlendirme, path traversal, race condition.
- Kenar durum (edge case) listesi çıkar: boş giriş, maksimum değer, eşzamanlı istek, ağ kesintisi.
- Hata tespit: assert değil, kanıt sun. Hangi satır, hangi koşul, hangi etki.
- Kod değişikliği yapma; yalnızca raporla ve test senaryosu öner. Düzeltme için ilgili ajana yönlendir.`
  },

  'content-writer': {
    description: 'SEO içerik yazarı. Arama amacına uygun, E-E-A-T ve dönüşüm odaklı akıcı Türkçe metin.',
    skill: 'ag-content-writer',
    body: `Sen SEO içerik yazarısın. Türkçe konuş.
- \`ag-content-writer\` skill kurallarını uygula.
- Arama niyetine uygun yaz: bilgilendirici, ticari, gezinme veya işlemsel amaca göre ton belirle.
- E-E-A-T: uzman bakış açısı, gerçek veri, güvenilir kaynak referansı, deneyim kanıtı.
- Ters piramit: en önemli bilgi ilk paragrafta. Giriş cümlesi ≤25 kelime.
- Teknik \`<head>\` / şema → \`seo-expert\` / \`schema-expert\` ajanına bırak; sen metni yaz.
- Dolgusuz yaz: "mükemmel", "harika", "kapsamlı" gibi boş sıfatlar yasak.`
  },

  'database-expert': {
    description: 'Veritabanı uzmanı (DBA). MySQL/MariaDB şema, indeksleme, sorgu ayarı, migration ve sayfalama.',
    skill: 'ag-database-expert',
    skillAlso: 'ag-php-developer → database.md (PHP entegrasyonu için)',
    body: `Sen veritabanı uzmanısın (DBA). Türkçe konuş.
- \`ag-database-expert\` skill kurallarını uygula; PHP entegrasyonu için \`ag-php-developer/database.md\` modülüne de bak.
- Şema tasarımı: 3NF normalleşme, doğru veri tipleri (utf8mb4, timestamptz, JSONB), NULL stratejisi.
- İndeks stratejisi: composite index sıra, covering index, gereksiz index maliyeti analizi.
- Sorgu optimizasyonu: EXPLAIN ANALYZE, index hint, partition, materialized view.
- Migration: çift adımlı (önce SQL dosyası yaz, sonra PDO ile uygula). Geri alma planı zorunlu.
- Sayfalama: OFFSET yerine keyset (cursor) pagination. Büyük tablo için zorunlu.
- Motor uzmanlığı: MySQL/MariaDB (InnoDB), PostgreSQL (JSONB/timestamptz), SQLite (WAL), Redis (TTL/önbellek), MongoDB, Elasticsearch.`
  },

  'devops-engineer': {
    description: 'DevOps mühendisi. CI/CD, Docker, Nginx/Apache, SSL, ortam değişkenleri ve güvenli deployment.',
    skill: 'ag-devops-engineer',
    body: `Sen DevOps mühendisisin. Türkçe konuş.
- \`ag-devops-engineer\` skill kurallarını uygula.
- Docker: çok aşamalı build, non-root kullanıcı, .dockerignore, health check zorunlu.
- Nginx/Apache: gzip, HTTP/2, güvenli header'lar (CSP, HSTS, X-Frame-Options), rate limiting.
- SSL/HTTPS: mkcert (yerel), Let's Encrypt (prod), TLS 1.2+ zorunlu, zayıf cipher yasak.
- Ortam değişkenleri: sırlar .env'de, .env asla git'e commit edilmez, .env.example güncel tutulur.
- CI/CD kalite kapıları: lint → test → build → güvenlik tarama → deploy sırası.
- Deployment: sıfır kesinti (blue-green veya rolling). Geri alma planı belge.`
  },

  'frontend-developer': {
    description: 'Frontend geliştirici. PHP şablonları, Tailwind CSS v4, Alpine.js, Bootstrap 5, açık tema, responsive UI.',
    skill: 'ag-php-developer',
    skillModule: 'frontend-developer.md',
    body: `Sen frontend geliştiricisisin. Türkçe konuş.
- \`ag-php-developer\` skill'inin \`frontend-developer.md\` modülünü uygula.
- Tailwind CSS v4 veya Bootstrap 5 ile responsive, mobile-first layout yaz.
- Alpine.js ile sade reaktif etkileşim; gereksiz JS framework kurma.
- WCAG 2.2 AA: renk kontrastı, focus görünürlüğü, anlamlı heading hiyerarşisi.
- Core Web Vitals: LCP < 2.5s, INP < 200ms, CLS < 0.1. Kritik CSS inline, asenkron JS.
- PHP şablonlarında XSS için \`htmlspecialchars()\` / \`e()\` zorunlu.
- Tasarım kararında belirsizlik varsa \`taste-skill\` prensiplerine bak.`
  },

  'html-export-expert': {
    description: 'HTML export uzmanı. Temiz W3C geçerli anlamsal HTML, sanitize ve e-posta şablonları.',
    skill: 'ag-html-export-expert',
    body: `Sen HTML export uzmanısın. Türkçe konuş.
- W3C geçerli, anlamsal HTML5 çıktısı üret. Her çıktıyı validator.w3.org ile doğrula.
- XSS sanitize: tüm kullanıcı girdileri \`htmlspecialchars(ENT_QUOTES, 'UTF-8')\` ile kaçırılır.
- E-posta şablonları: tablo bazlı layout, inline CSS, max-width 600px, dark mode media query.
- BOM yasak: çıktı dosyaları BOM'suz UTF-8. \`bom_utf8_scan.js --fix\` ile doğrula.
- Erişilebilirlik: \`lang\` attribute, anlamlı \`alt\`, heading hiyerarşisi, \`role\` attribute.
- Çıktı temizliği: script, on* attribute, style injection içermeyen güvenli HTML.`
  },

  'mobile-optimization-expert': {
    description: 'Mobil performans uzmanı. Core Web Vitals (LCP/INP/CLS), mobil UX ve asset bütçesi.',
    skill: 'ag-mobile-optimization-expert',
    body: `Sen mobil performans uzmanısın. Türkçe konuş.
- Core Web Vitals hedefleri: LCP < 2.5s, INP < 200ms, CLS < 0.1.
- Görsel optimizasyon: WebP/AVIF format, boyut attribute, lazy loading, \`srcset\` responsive.
- Kritik CSS inline; render-blocking JS'i \`defer\`/\`async\` ile erteleme.
- Font: \`font-display: swap\`, sadece kullanılan weight/subset yükle.
- Asset bütçesi: JS < 150KB (gzip), CSS < 50KB, toplam sayfa < 1MB mobilde.
- Touch hedefleri: minimum 44×44px. Horizontal scroll yasak. Viewport meta zorunlu.
- Ölçüm: Lighthouse CI, WebPageTest, Chrome DevTools Performance panel ile kanıtla.`
  },

  'php-developer': {
    description: 'PHP backend geliştirici. Rota, controller, model, PDO, migration, güvenli iş mantığı ve php -l denetimi.',
    skill: 'ag-php-developer',
    skillModule: 'Tüm modüller (frontend + backend + security + database + frameworks)',
    body: `Sen kıdemli PHP full-stack geliştiricisisin. Türkçe konuş.
- \`ag-php-developer\` skill'inin tüm modüllerini uygula: frontend, backend, security, database, frameworks.
- Her dosyanın başında \`declare(strict_types=1);\` zorunlu.
- PDO prepared statement; kullanıcı girdisi asla doğrudan SQL'e.
- XSS koruması: çıktıda \`htmlspecialchars()\` / \`e()\` zorunlu.
- CSRF: durum değiştiren her POST isteğinde token doğrula.
- ZORUNLU DOĞRULAMA: Her değişiklik sonrası \`php -l <dosya>\` çalıştır; "No syntax errors" kanıtı olmadan bitti deme.
- BOM yasak: \`bom_utf8_scan.js --fix\` ile doğrula.
- Dosya yazma önceliği: 1. Node.js fs → 2. Python 3 / Bash → 3. dahili araçlar. BOM kesinlikle yasaktır, saf UTF-8 zorunludur.`
  },

  'php-security-expert': {
    description: 'PHP güvenlik uzmanı. OWASP Top 10 (SQLi, XSS, CSRF, IDOR, yetki, sanitize) denetimi.',
    skill: 'ag-php-developer',
    skillModule: 'php-security.md',
    body: `Sen PHP güvenlik uzmanısın. Türkçe konuş.
- \`ag-php-developer\` skill'inin \`php-security.md\` modülünü uygula.
- OWASP Top 10 denetimi: SQLi, XSS, CSRF, IDOR, path traversal, LFI/RFI, open redirect, broken auth.
- SQLi: PDO prepared statement. Ham sorgu = kritik bulgu.
- XSS: çıktıda \`htmlspecialchars(ENT_QUOTES,'UTF-8')\`. innerHTML kullanımı = kritik bulgu.
- CSRF: \`$_SESSION['csrf']\` token doğrulama; tüm POST formlarında.
- IDOR: yetki kontrolü resource ID'sinden önce. \`$_GET['id']\` doğrudan sorgu = kritik bulgu.
- Dosya yükleme: MIME doğrulama, web kökü dışı depolama, whitelist uzantı.
- Rapor formatı: BULGU / ETKİ / KANIT / DÜZELTME. Kod değişikliği yapma; yalnızca raporla.`
  },

  'project-manager': {
    description: 'Proje yöneticisi ve baş denetçi. Görev kurgulama, dilimleme, kalite kapıları (4/4 PASS) ve teslimat.',
    skill: 'ag-project-manager',
    body: `Sen proje yöneticisi ve baş denetçisin. Türkçe konuş.
- \`ag-project-manager\` skill kurallarını uygula.
- Görevi al → doğru uzmana (agent) delege et → çıktıyı denetle → teslim et.
- Kalite kapısı: fonksiyonellik ✓, güvenlik ✓, performans ✓, erişilebilirlik ✓ — 4/4 PASS olmadan teslimat yok.
- Görev kırılımı: P0 (kritik bloker) → P1 (önemli) → P2 (normal) → P3 (nice-to-have).
- Sprint takibi: her görev için kabul kriteri yaz; "tamamlandı" kanıtsız kabul edilmez.
- Tüm alanlarda kaliteyi yargılayacak kadar uzman; detay uygulamayı ilgili ajana yaptır.`
  },

  'schema-expert': {
    description: 'Yapısal veri uzmanı. Schema.org / JSON-LD (Product, WebSite, BreadcrumbList, Article, FAQPage).',
    skill: 'ag-schema-expert',
    body: `Sen yapısal veri (JSON-LD) uzmanısın. Türkçe konuş.
- Schema.org standartlarına ve Google Rich Results yönergelerine tam uyumlu JSON-LD üret.
- Desteklenen tipler: Product, WebSite, WebPage, Article, BlogPosting, FAQPage, BreadcrumbList, Organization, LocalBusiness, Event, HowTo, Recipe.
- aggregateRating: yalnızca gerçek verilerle. Sahte puan üretmek kesinlikle yasak.
- JSON-LD çıktısında özel karakterleri güvenli formatta kaçır (\`<script>\` kapanışları).
- Her şema için Google Rich Results Test URL'si ile doğrulama adımı ver.
- Sayfa başına bir \`@graph\` container; birden fazla şema türü tek blokta birleştir.`
  },

  'seo-expert': {
    description: 'Teknik SEO uzmanı. Title/meta/canonical, Open Graph, Twitter, sitemap.xml ve robots.txt.',
    skill: 'ag-seo-expert',
    body: `Sen teknik SEO uzmanısın. Türkçe konuş.
- \`ag-seo-expert\` skill kurallarını uygula.
- Meta title: 50-60 karakter, birincil anahtar kelime başta, marka sonda (\` | Marka\`).
- Meta description: 150-160 karakter, CTA içeren, unique her sayfa için.
- Canonical URL: mutlak URL, self-referencing, hreflang ile uyumlu.
- Open Graph: \`og:title\`, \`og:description\`, \`og:image\` (1200×630px), \`og:type\`.
- Twitter Card: \`summary_large_image\`, \`twitter:site\`, \`twitter:creator\`.
- sitemap.xml: tüm index edilecek URL'ler, \`<lastmod>\`, \`<changefreq>\`, \`<priority>\`.
- robots.txt: \`Disallow\` kuralları, sitemap referansı, crawl-delay.`
  },

  'seo-structure-tester': {
    description: 'Sayfa SEO test uzmanı. H1-H6 hiyerarşisi, landmark etiketleri ve JSON-LD yapı denetimi.',
    skill: 'ag-seo-structure-tester',
    body: `Sen sayfa içi SEO ve anlamsal yapı denetim uzmanısın. Türkçe konuş.
- Başlık hiyerarşisi: tek H1, seviye atlama yok (H2→H4 geçiş yasak), H1 ana anahtar kelimeyi içermeli.
- Landmark etiketleri: \`<main>\`, \`<header>\`, \`<footer>\`, \`<nav>\`, \`<aside>\` doğru kullanımı denetle.
- JSON-LD: geçerli şema, Google Rich Results Test'ten hatasız geçiş, @context ve @type varlığı.
- Tekrar eden H1 varsa kritik bulgu. Boş heading varsa kritik bulgu.
- Yalnızca GEÇTI/KALDI raporu ve eksik liste sun. Kod değişikliği yapma.
- Çıktı formatı: ✅ GEÇTI / ❌ KALDI / ⚠️ UYARI — madde madde liste.`
  },

  'technical-writer': {
    description: 'Teknik yazar. docs/ altında README, kurulum, API kılavuzu ve changelog dokümantasyonu.',
    skill: 'ag-technical-writer',
    body: `Sen teknik dokümantasyon uzmanısın. Türkçe konuş.
- \`ag-technical-writer\` skill kurallarını uygula.
- Tüm dokümantasyon \`docs/\` klasöründe; ana kök dizini temiz tut.
- README: kurulum adımları, örnek kullanım, gereksinimler, katkı rehberi.
- API dokümantasyonu: endpoint, HTTP method, parametreler, istek/yanıt örnekleri (JSON), hata kodları.
- Changelog: Keep a Changelog formatı (Added / Changed / Deprecated / Removed / Fixed / Security).
- Teknik dil; pazarlama metni değil. Geliştirici ve sistem odaklı pratik içerik.
- OpenAPI/Swagger YAML formatında API spec gerektiğinde üret.`
  },

  'test-engineer': {
    description: 'Test mühendisi (QA). PHPUnit, Pest, E2E testleri ve regresyon doğrulama ağı.',
    skill: 'ag-test-engineer',
    body: `Sen QA ve test mühendisisin. Türkçe konuş.
- \`ag-test-engineer\` skill kurallarını uygula.
- Test piramidi: birim (unit) > entegrasyon (integration) > E2E. Birim testler izole ve hızlı.
- PHPUnit / Pest ile PHP testleri; Playwright ile E2E web testleri.
- Her test için: Arrange → Act → Assert yapısı. Bir test bir davranışı doğrular.
- Regresyon ağı: değişiklik öncesi tüm mevcut testler yeşil, değişiklik sonrası da yeşil.
- Iron Law: "çalışıyor" demeden önce test koşum çıktısını fiziksel kanıt olarak sun.
- \`bug-hunter\` hata avlar; sen kalıcı test ağı örer — görev ayrımını koru.`
  }
};

// ─── YENİ SKILL'LER ──────────────────────────────────────────────────────────
const newSkills = {

  'ag-accessibility-expert': {
    name: 'ag-accessibility-expert',
    description: 'Erişilebilirlik uzmanı (a11y). WCAG 2.2 AA denetimi, ARIA, klavye navigasyonu, kontrast ve form erişilebilirliği.',
    content: `# Erişilebilirlik Uzmanı (a11y) — Uzmanlık Dokümanı

Rolü: WCAG 2.2 AA standartlarına göre sayfa ve bileşen denetimi. Kod değişikliği yapmaz; raporlar ve yönlendirir.

---

## 1. WCAG 2.2 AA Temel İlkeleri

### Algılanabilirlik
- Renk kontrastı: normal metin ≥4.5:1, büyük metin (18pt+) ≥3:1.
- Görsel içerikler için anlamlı \`alt\` attribute. Dekoratif görseller: \`alt=""\`.
- Video/ses içeriği: altyazı ve transkript zorunlu.

### Çalıştırılabilirlik
- Tüm işlevler yalnızca klavyeyle erişilebilir (Tab, Enter, Escape, Arrow tuşları).
- Focus tuzağı (focus trap): modal dışına Tab ile çıkılmamalı; modal kapatıldığında tetikleyiciye dön.
- Skip to content linki: ilk odak elemanı olarak.
- \`:focus-visible\` ile odak görünürlüğü; \`outline: none\` yasak.

### Anlaşılabilirlik
- Dil attribute: \`<html lang="tr">\`.
- Form validasyon hataları: kullanıcı dostu metin + \`aria-invalid="true"\` + \`aria-describedby\`.
- Tutarlı navigasyon: aynı bileşen her sayfada aynı sırada.

### Sağlamlık
- Geçerli HTML: etiket kapatma, iç içe geçme hataları yok.
- ARIA rolleri doğru: \`role="button"\` tıklanabilir \`<div>\`'e değil, \`<button>\`'a.

---

## 2. ARIA Kullanım Kuralları

| Doğru | Yanlış |
|---|---|
| \`<button>\` kullan | \`<div role="button">\` |
| \`aria-label\` (görsel yoksa) | Boş \`alt\` text olmadan görsel |
| \`aria-expanded="true/false"\` | Açılır menüde state yok |
| \`aria-live="polite"\` (dinamik içerik) | Anlık güncelleme sessiz |
| \`<label for="id">\` | \`placeholder\` yerine label |

---

## 3. Form Erişilebilirliği

\`\`\`html
<label for="email">E-posta</label>
<input
  id="email"
  type="email"
  aria-describedby="email-hint email-error"
  aria-invalid="false"
  required
>
<span id="email-hint">Kayıt e-posta adresinizi girin.</span>
<span id="email-error" role="alert" hidden>Geçerli bir e-posta adresi girin.</span>
\`\`\`

---

## 4. Denetim Raporu Formatı

\`\`\`
## a11y Denetim Raporu — [Sayfa/Bileşen Adı]

### ❌ Kritik (WCAG 2.2 AA İhlali)
- [Sorun]: [Etki] → [Düzeltme önerisi]

### ⚠️ Uyarı (Best Practice)
- [Sorun]: [Öneri]

### ✅ Geçti
- [Kontrol edilen ve geçen maddeler]
\`\`\`

---

## 5. Test Araçları
- Lighthouse Accessibility panel (Chrome DevTools)
- axe DevTools browser extension
- NVDA + Chrome (Windows screen reader testi)
- Keyboard-only navigation manuel testi
`
  },

  'ag-bug-hunter': {
    name: 'ag-bug-hunter',
    description: 'Hata avcısı. Mantık hataları, güvenlik açıkları, kenar durumlar ve statik analiz denetimi.',
    content: `# Hata Avcısı (Bug Hunter) — Uzmanlık Dokümanı

Rolü: mantık hatası, güvenlik açığı ve kenar durum avı. Kod değişikliği yapmaz; raporlar.

---

## 1. Statik Analiz Kontrol Listesi

### Tip ve Null Güvenliği
- Null dereference: \`$obj->method()\` öncesi null check var mı?
- Tip uyumsuzluğu: \`strict_types=1\` eksik mi? Gevşek karşılaştırma (\`==\`) kullanımı?
- Integer overflow: büyük sayı hesaplamalarında \`bcmath\` kullanılıyor mu?
- Array key existence: \`isset()\` veya \`array_key_exists()\` olmadan erişim?

### Mantık Hataları
- Off-by-one: döngü sınırları, dizi indeksleri.
- Yanlış koşul: \`&&\` / \`||\` önceliği, negasyon hatası.
- Dead code: asla ulaşılamayan dal.
- Yanlış değişken: benzer isimli değişkenlerin karıştırılması.

### Eşzamanlılık
- Race condition: dosya yazma, sayaç güncelleme, token kullanımı.
- Session fixation: login sonrası \`session_regenerate_id()\` eksikliği.

---

## 2. Güvenlik Açığı Tarama

| Saldırı Tipi | Kontrol Noktası |
|---|---|
| SQLi | Ham sorgu, \`$_GET\` direkt SQL'de |
| XSS | \`htmlspecialchars()\` eksik çıktı |
| CSRF | POST formlarında token yok |
| IDOR | Yetki kontrolü ID'den önce değil |
| Path Traversal | \`../\` filtresiz dosya yolu |
| Open Redirect | \`header("Location: ".$_GET['url'])\` |
| File Upload | MIME doğrulama yok, web kökünde depolama |

---

## 3. Kenar Durum Listesi

Her fonksiyon / endpoint için:
- [ ] Boş string / null giriş
- [ ] Negatif / sıfır sayı
- [ ] Maksimum uzunluk / değer aşımı
- [ ] Eşzamanlı istek (aynı anda 2 kullanıcı)
- [ ] Ağ kesintisi / timeout
- [ ] Veritabanı bağlantı hatası
- [ ] Yetkisiz kullanıcı erişimi

---

## 4. Bulgu Raporu Formatı

\`\`\`
## Bug Hunt Raporu — [Dosya/Modül Adı]

### 🔴 KRİTİK
**Bulgu:** [Açıklama]
**Satır:** [Dosya:Satır]
**Etki:** [Ne olur?]
**Kanıt:** [Kod snippet]
**Düzeltme:** [Ne yapılmalı] → ilgili ajana yönlendir

### 🟡 ORTA
...

### 🟢 DÜŞÜK / BİLGİ
...
\`\`\`
`
  },

  'ag-html-export-expert': {
    name: 'ag-html-export-expert',
    description: 'HTML export uzmanı. Temiz W3C geçerli anlamsal HTML, sanitize ve e-posta şablonları.',
    content: `# HTML Export Uzmanı — Uzmanlık Dokümanı

Rolü: W3C geçerli, güvenli, taşınabilir HTML üretimi. E-posta şablonları, statik export, sanitize.

---

## 1. Anlamsal HTML5 Standartları

- Doküman: \`<!DOCTYPE html>\`, \`<html lang="tr">\`, \`<meta charset="UTF-8">\`, \`<meta name="viewport">\`.
- Başlık hiyerarşisi: tek H1, seviye atlamama (H1→H2→H3).
- Landmark'lar: \`<header>\`, \`<main>\`, \`<footer>\`, \`<nav>\`, \`<aside>\`, \`<section>\`, \`<article>\`.
- Liste: sıralı \`<ol>\`, sırasız \`<ul>\`, tanım \`<dl>\`.
- Tablo: \`<thead>\`, \`<tbody>\`, \`<th scope="col/row">\`, \`caption\`.

---

## 2. XSS Sanitize Kuralları

\`\`\`php
// PHP - tüm kullanıcı girdisi için:
echo htmlspecialchars($input, ENT_QUOTES | ENT_HTML5, 'UTF-8');

// Attribute içinde:
echo '<input value="' . htmlspecialchars($val, ENT_QUOTES, 'UTF-8') . '">';

// URL:
echo '<a href="' . htmlspecialchars(filter_var($url, FILTER_SANITIZE_URL), ENT_QUOTES) . '">';
\`\`\`

**Yasak:** \`innerHTML\` ile kullanıcı girdisi, \`eval()\`, \`document.write()\`, inline \`on*\` attribute.

---

## 3. E-posta Şablonu Kuralları

\`\`\`html
<!-- E-posta: tablo bazlı layout -->
<table width="100%" cellpadding="0" cellspacing="0" border="0">
  <tr>
    <td align="center">
      <table width="600" cellpadding="0" cellspacing="0">
        <!-- içerik -->
      </table>
    </td>
  </tr>
</table>
\`\`\`

- CSS: yalnızca inline style. \`<style>\` bloğu desteklemeyen istemciler var.
- Font: web-safe font stack (Arial, Helvetica, sans-serif).
- Görsel: mutlak URL, width/height attribute, \`alt\` zorunlu.
- Genişlik: max 600px. Mobil: media query ile tek sütun.
- Dark mode: \`@media (prefers-color-scheme: dark)\` ile renk uyumu.

---

## 4. BOM ve Encoding

- Tüm dosyalar BOM'suz UTF-8. BOM tarama: \`bom_utf8_scan.js --fix\`.
- PHP çıktısında \`header('Content-Type: text/html; charset=UTF-8');\` ilk satır.
- Türkçe karakter bozulması (Ã¼, Å, vb.) → encoding hatası; UTF-8 pipeline'ı kontrol et.

---

## 5. Doğrulama
- HTML: validator.w3.org (Nu HTML Checker)
- E-posta: Litmus veya Email on Acid önizleme
- Erişilebilirlik: axe DevTools
`
  },

  'ag-mobile-optimization-expert': {
    name: 'ag-mobile-optimization-expert',
    description: 'Mobil performans uzmanı. Core Web Vitals (LCP/INP/CLS), mobil UX ve asset bütçesi.',
    content: `# Mobil Optimizasyon Uzmanı — Uzmanlık Dokümanı

Rolü: Core Web Vitals optimizasyonu, mobil UX ve asset bütçesi yönetimi.

---

## 1. Core Web Vitals Hedefleri

| Metrik | İyi | İyileştirme Gerekir | Kötü |
|---|---|---|---|
| LCP (Largest Contentful Paint) | < 2.5s | 2.5–4s | > 4s |
| INP (Interaction to Next Paint) | < 200ms | 200–500ms | > 500ms |
| CLS (Cumulative Layout Shift) | < 0.1 | 0.1–0.25 | > 0.25 |

---

## 2. LCP Optimizasyon

- LCP elementi (hero görsel/metin) için \`fetchpriority="high"\` attribute.
- \`<link rel="preload" as="image">\` ile kritik görseli önceden yükle.
- Görsel format: WebP (öncelikli), AVIF (destek varsa). PNG/JPEG fallback.
- Sunucu yanıt süresi (TTFB) < 600ms. CDN kullan.
- Kritik CSS: above-the-fold için inline style. Render-blocking CSS yasak.

\`\`\`html
<!-- Doğru LCP görsel yükleme -->
<img
  src="hero.webp"
  alt="..."
  width="1200" height="600"
  fetchpriority="high"
  loading="eager"
>
\`\`\`

---

## 3. INP Optimizasyon

- Long task (50ms+) parçala: \`requestIdleCallback\`, \`scheduler.yield()\`.
- Event handler'ları optimize et: debounce/throttle, passive listener.
- JavaScript bundle: code splitting, tree shaking, lazy import.
- Ana thread blokajı < 200ms.

---

## 4. CLS Optimizasyon

- Görsel/video: \`width\` ve \`height\` attribute her zaman (aspect-ratio CSS ile).
- Font: \`font-display: swap\` + \`<link rel="preload">\` + local fallback.
- Reklam/embed: sabit boyutlu container rezerve et.
- Dinamik içerik: sayfanın üstüne ekleme. Altına veya placeholder ile.

---

## 5. Asset Bütçesi

| Asset Türü | Hedef (gzip) |
|---|---|
| JavaScript (toplam) | < 150KB |
| CSS (toplam) | < 50KB |
| Toplam sayfa | < 1MB mobilde |
| Görsel (her biri) | < 100KB (hero < 200KB) |
| Font (her face) | < 30KB (subset) |

---

## 6. Mobil UX Kuralları

- Viewport: \`<meta name="viewport" content="width=device-width, initial-scale=1">\`.
- Touch hedefi: minimum 44×44px (Apple HIG), 48×48px (Material).
- Horizontal scroll: yasak. \`overflow-x: hidden\` yerine içeriği düzelt.
- Font boyutu: minimum 16px body (zoom tetiklememek için).
- Tap delay: \`touch-action: manipulation\` ile 300ms gecikmeyi kaldır.

---

## 7. Ölçüm ve Kanıt

\`\`\`bash
# Lighthouse CI
npx lhci autorun --collect.url=https://example.com

# WebPageTest
# webpagetest.org → Advanced Settings → Mobile (4G)

# Chrome DevTools
# Performance panel → Record → mobile throttling
\`\`\`

Raporlama: before/after metrik tablosu. Kanıtsız "optimize edildi" ifadesi kullanılmaz.
`
  },

  'ag-schema-expert': {
    name: 'ag-schema-expert',
    description: 'Yapısal veri uzmanı. Schema.org / JSON-LD (Product, WebSite, BreadcrumbList, Article, FAQPage).',
    content: `# Yapısal Veri Uzmanı (Schema.org) — Uzmanlık Dokümanı

Rolü: Google Rich Results uyumlu JSON-LD üretimi. Sahte veri yasak; yalnızca gerçek verilerle çalış.

---

## 1. Desteklenen Schema Tipleri

| Tip | Kullanım | Rich Result |
|---|---|---|
| \`Product\` | E-ticaret ürün sayfası | Ürün carousel, fiyat |
| \`Article\` / \`BlogPosting\` | Blog, haber | Top stories |
| \`FAQPage\` | SSS sayfası | FAQ accordion |
| \`HowTo\` | Adım adım rehber | HowTo snippet |
| \`BreadcrumbList\` | Tüm sayfalar | Breadcrumb |
| \`WebSite\` | Ana sayfa | Sitelinks searchbox |
| \`Organization\` | Hakkımızda | Knowledge panel |
| \`LocalBusiness\` | Yerel işletme | Local pack |
| \`Event\` | Etkinlik sayfası | Event listing |
| \`Recipe\` | Tarif sayfası | Recipe card |

---

## 2. @graph Yapısı (Önerilen)

\`\`\`json
<script type="application/ld+json">
{
  "@context": "https://schema.org",
  "@graph": [
    {
      "@type": "WebSite",
      "@id": "https://example.com/#website",
      "url": "https://example.com/",
      "name": "Site Adı",
      "inLanguage": "tr"
    },
    {
      "@type": "BreadcrumbList",
      "itemListElement": [
        { "@type": "ListItem", "position": 1, "name": "Ana Sayfa", "item": "https://example.com/" },
        { "@type": "ListItem", "position": 2, "name": "Kategori", "item": "https://example.com/kategori/" }
      ]
    }
  ]
}
</script>
\`\`\`

---

## 3. Product Şeması (E-ticaret)

\`\`\`json
{
  "@type": "Product",
  "name": "Ürün Adı",
  "image": ["https://example.com/urun.jpg"],
  "description": "Ürün açıklaması.",
  "sku": "SKU-123",
  "brand": { "@type": "Brand", "name": "Marka" },
  "offers": {
    "@type": "Offer",
    "url": "https://example.com/urun",
    "priceCurrency": "TRY",
    "price": "299.00",
    "availability": "https://schema.org/InStock"
  }
}
\`\`\`

**aggregateRating YASAĞI:** Gerçek değerlendirme verisi olmadan \`aggregateRating\` bloğu üretilmez.

---

## 4. FAQPage Şeması

\`\`\`json
{
  "@type": "FAQPage",
  "mainEntity": [
    {
      "@type": "Question",
      "name": "Soru metni?",
      "acceptedAnswer": {
        "@type": "Answer",
        "text": "Cevap metni."
      }
    }
  ]
}
\`\`\`

---

## 5. Doğrulama Adımı

Her şema çıktısı için:
1. [Google Rich Results Test](https://search.google.com/test/rich-results)
2. [Schema.org Validator](https://validator.schema.org/)
3. Hata yoksa → ✅ GEÇTI | Hata varsa → 🔴 KALDI (hata mesajı ile rapor)
`
  },

  'ag-seo-structure-tester': {
    name: 'ag-seo-structure-tester',
    description: 'Sayfa SEO test uzmanı. H1-H6 hiyerarşisi, landmark etiketleri ve JSON-LD yapı denetimi.',
    content: `# Sayfa SEO Yapı Test Uzmanı — Uzmanlık Dokümanı

Rolü: sayfa içi SEO ve anlamsal yapı denetimi. Yalnızca raporlar; kod değişikliği yapmaz.

---

## 1. Başlık Hiyerarşisi Denetimi

### Kurallar
- [ ] Sayfada **yalnızca bir H1** var mı?
- [ ] H1 ana anahtar kelimeyi içeriyor mu?
- [ ] Seviye atlamak yok (H1→H3 doğrudan geçiş yasak)?
- [ ] Boş heading yok (\`<h2></h2>\`)?
- [ ] Heading text'i anlamlı (logo alt yazısı, dekoratif metin değil)?

### Geçerli Hiyerarşi
\`\`\`
H1 → Ana Başlık
  H2 → Bölüm 1
    H3 → Alt Bölüm 1.1
    H3 → Alt Bölüm 1.2
  H2 → Bölüm 2
\`\`\`

---

## 2. Landmark Etiket Denetimi

| Etiket | Kontrol |
|---|---|
| \`<header>\` | Sayfada mevcut, site başlığı içeriyor |
| \`<main>\` | Yalnızca bir tane, temel içeriği kapsıyor |
| \`<nav>\` | Navigasyon menüsü için kullanılıyor |
| \`<footer>\` | Sayfanın altında, footer içeriği kapsıyor |
| \`<aside>\` | Yan içerik için (opsiyonel ama varsa doğru kullanım) |

---

## 3. JSON-LD Yapı Denetimi

- [ ] \`<script type="application/ld+json">\` mevcut mu?
- [ ] \`@context\` ve \`@type\` var mı?
- [ ] JSON sözdizimi geçerli mi (parse hatası yok)?
- [ ] Google Rich Results Test'ten hatasız geçiyor mu?
- [ ] Sahte \`aggregateRating\` (gerçek verisi olmadan) var mı? → KRİTİK BULGU

---

## 4. Rapor Formatı

\`\`\`
## SEO Yapı Test Raporu — [URL veya Sayfa Adı]
Tarih: [YYYY-MM-DD]

### Başlık Hiyerarşisi
✅ Tek H1: "[H1 metni]"
❌ Seviye atlama: H2 → H4 geçişi (satır X)
⚠️ H2 metni kısa (< 3 kelime): "[metin]"

### Landmark Etiketleri
✅ <header> mevcut
✅ <main> mevcut
❌ <nav> eksik — navigasyon <div class="nav"> içinde

### JSON-LD
✅ @context: https://schema.org
✅ @type: Product
❌ aggregateRating: gerçek veri kaynağı yok — KRİTİK

### Özet
Kritik: 2 | Uyarı: 1 | Geçti: 5
\`\`\`
`
  },

  'ag-technical-writer': {
    name: 'ag-technical-writer',
    description: 'Teknik yazar. docs/ altında README, kurulum, API kılavuzu ve changelog dokümantasyonu.',
    content: `# Teknik Yazar — Uzmanlık Dokümanı

Rolü: geliştiriciye yönelik teknik dokümantasyon. Pazarlama değil; pratik, net, doğru içerik.

---

## 1. Dokümantasyon Mimarisi

\`\`\`
docs/
├── README.md            # Proje özeti, kurulum, hızlı başlangıç
├── CHANGELOG.md         # Keep a Changelog formatı
├── CONTRIBUTING.md      # Katkı rehberi
├── api/
│   └── endpoints.md     # API endpoint belgeleri
├── guides/
│   ├── installation.md  # Detaylı kurulum
│   └── configuration.md # Yapılandırma seçenekleri
└── architecture.md      # Sistem mimarisi
\`\`\`

---

## 2. README Yapısı

\`\`\`markdown
# Proje Adı

Kısa açıklama (1-2 cümle). Ne yapar, kim için?

## Özellikler
- Madde 1
- Madde 2

## Gereksinimler
- PHP 8.3+
- MySQL 8.0+

## Kurulum
\`\`\`bash
composer install
cp .env.example .env
php artisan migrate
\`\`\`

## Kullanım
[Örnek kod veya komut]

## API
[Kısa referans veya docs/ linkine yönlendirme]

## Katkı
[CONTRIBUTING.md linkine yönlendirme]

## Lisans
[Lisans türü]
\`\`\`

---

## 3. API Endpoint Belgeleme Şablonu

\`\`\`markdown
### POST /api/v1/products

Yeni ürün oluşturur.

**Yetkilendirme:** Bearer Token gerekli

**İstek Gövdesi:**
| Alan | Tip | Zorunlu | Açıklama |
|---|---|---|---|
| name | string | ✅ | Ürün adı (max 255 karakter) |
| price | number | ✅ | Fiyat (TRY, ≥0) |

**Başarılı Yanıt (201):**
\`\`\`json
{
  "data": { "id": 1, "name": "Ürün", "price": 299.00 },
  "message": "Ürün oluşturuldu."
}
\`\`\`

**Hata Yanıtları:**
| Kod | Açıklama |
|---|---|
| 400 | Geçersiz veri |
| 401 | Yetki hatası |
\`\`\`

---

## 4. Changelog Formatı (Keep a Changelog)

\`\`\`markdown
# Changelog

## [1.2.0] - 2026-09-11
### Added
- Yeni özellik açıklaması

### Changed
- Değişen davranış açıklaması

### Fixed
- Düzeltilen hata açıklaması

### Security
- Güvenlik yaması açıklaması
\`\`\`

---

## 5. Yazım Kuralları
- Ters piramit: en önemli bilgi ilk satırda.
- Cümle ≤ 20 kelime. Paragraf ≤ 5 cümle.
- Aktif cümle: "Sistem veriyi işler" değil "Veriyi işle" (komut kipinde talimat).
- Yasak sözcükler: "mükemmel", "güçlü", "kapsamlı", "esnek" (dolgu sıfatlar).
- Kod örnekleri: gerçek çalışır kod. Pseudocode yerine gerçek syntax.
`
  }
};

// ─── YARDIMCI FONKSİYONLAR ───────────────────────────────────────────────────
function ensureDir(dirPath) {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
    console.log(`  📁 Dizin oluşturuldu: ${dirPath}`);
  }
}

function writeUTF8(filePath, content) {
  ensureDir(path.dirname(filePath));
  fs.writeFileSync(filePath, content, { encoding: 'utf8' });
  // BOM kontrolü
  const buf = fs.readFileSync(filePath);
  if (buf[0] === 0xEF && buf[1] === 0xBB && buf[2] === 0xBF) {
    fs.writeFileSync(filePath, buf.slice(3));
    console.log(`  🔧 BOM temizlendi: ${filePath}`);
  }
}

function buildAgentMd(name, cfg) {
  const skillRef = cfg.skillAlso
    ? `skills: ${cfg.skill} (+ ${cfg.skillAlso})`
    : `skills: ${cfg.skill}`;
  const moduleNote = cfg.skillModule ? `\nskills-module: ${cfg.skillModule}` : '';

  return `---
name: ${name}
description: ${cfg.description}
${skillRef}${moduleNote}
---

${cfg.body}
${TOOL_PATHS}`;
}

// ─── ÇALIŞTIR ────────────────────────────────────────────────────────────────
console.log('\n══════════════════════════════════════════');
console.log(' AGENTS ↔ SKILLS SYNC & OPTIMIZE');
console.log('══════════════════════════════════════════\n');

// 1. Tüm agent.md dosyalarını güncelle
console.log('▶ AGENT.MD DOSYALARI GÜNCELLENİYOR...\n');
for (const [agentName, cfg] of Object.entries(agents)) {
  const agentDir = path.join(AGENTS_DIR, agentName);
  ensureDir(agentDir);
  const agentFile = path.join(agentDir, 'agent.md');
  const content = buildAgentMd(agentName, cfg);
  writeUTF8(agentFile, content);
  console.log(`  ✅ ${agentName}/agent.md → skill: ${cfg.skill}`);
}

// 2. Yeni skill klasörleri ve SKILL.md dosyaları oluştur
console.log('\n▶ YENİ SKILL\'LER OLUŞTURULUYOR...\n');
for (const [skillDir, skillCfg] of Object.entries(newSkills)) {
  const skillPath = path.join(SKILLS_DIR, skillDir);
  ensureDir(skillPath);
  const skillFile = path.join(skillPath, 'SKILL.md');

  const frontmatter = `---\nname: ${skillCfg.name}\ndescription: ${skillCfg.description}\n---\n\n`;
  const fullContent = frontmatter + skillCfg.content;
  writeUTF8(skillFile, fullContent);
  console.log(`  ✅ ${skillDir}/SKILL.md oluşturuldu`);
}

// 3. ag-php-developer'a database.md modülü ekle (eğer yoksa)
console.log('\n▶ ag-php-developer DATABASE MODÜLÜ EKLENİYOR...\n');
const dbModulePath = path.join(SKILLS_DIR, 'ag-php-developer', 'database.md');
if (!fs.existsSync(dbModulePath)) {
  const dbContent = `# Veritabanı Modülü — ag-php-developer

Bu modül \`ag-php-developer\` şemsiyesinin bir parçasıdır.
Derin DBA görevleri için \`ag-database-expert\` skill'ini kullan.

---

## 1. PHP-MySQL Entegrasyon Standartları

### PDO Bağlantısı

\`\`\`php
<?php
declare(strict_types=1);

$pdo = new PDO(
    'mysql:host=localhost;dbname=mydb;charset=utf8mb4',
    $_ENV['DB_USER'],
    $_ENV['DB_PASS'],
    [
        PDO::ATTR_ERRMODE            => PDO::ERRMODE_EXCEPTION,
        PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
        PDO::ATTR_EMULATE_PREPARES   => false,
    ]
);
\`\`\`

### Prepared Statement Zorunluluğu

\`\`\`php
// ✅ DOĞRU
$stmt = $pdo->prepare('SELECT * FROM users WHERE email = :email');
$stmt->execute([':email' => $email]);
$user = $stmt->fetch();

// ❌ YASAK — SQL Injection
$result = $pdo->query("SELECT * FROM users WHERE email = '$email'");
\`\`\`

---

## 2. Migration Kuralı (Çift Adımlı)

\`\`\`
1. database/Migrations/NNN_ad.sql dosyasını oluştur (bir sonraki numara)
2. Aktif veritabanına doğrudan PDO ile uygula
3. Geri alma (rollback) SQL'ini de yaz
\`\`\`

\`\`\`sql
-- database/Migrations/001_create_users.sql
CREATE TABLE users (
    id         BIGINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    email      VARCHAR(255) NOT NULL UNIQUE,
    created_at TIMESTAMP    NOT NULL DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
\`\`\`

---

## 3. N+1 Sorgu Yasağı

\`\`\`php
// ❌ YASAK — N+1
foreach ($orders as $order) {
    $user = $pdo->query("SELECT * FROM users WHERE id = {$order['user_id']}")->fetch();
}

// ✅ DOĞRU — JOIN ile tek sorgu
$stmt = $pdo->prepare('
    SELECT o.*, u.name, u.email
    FROM orders o
    JOIN users u ON u.id = o.user_id
    WHERE o.status = :status
');
$stmt->execute([':status' => 'active']);
$orders = $stmt->fetchAll();
\`\`\`

---

## 4. Sayfalama — Keyset (Cursor)

\`\`\`php
// ❌ YASAK — büyük tabloda yavaş
$stmt = $pdo->prepare('SELECT * FROM posts LIMIT :limit OFFSET :offset');

// ✅ DOĞRU — keyset pagination
$stmt = $pdo->prepare('
    SELECT * FROM posts
    WHERE id > :last_id
    ORDER BY id ASC
    LIMIT :limit
');
$stmt->execute([':last_id' => $lastId, ':limit' => $perPage]);
\`\`\`

---

## 5. Veritabanı Kolonu Tipleri (Best Practice)

| Tip | Kullanım |
|---|---|
| \`BIGINT UNSIGNED\` | Primary key, foreign key |
| \`VARCHAR(255)\` | E-posta, isim, slug |
| \`TEXT\` | Uzun metin (≤65KB) |
| \`MEDIUMTEXT\` | Makale içeriği |
| \`DECIMAL(10,2)\` | Para birimi (FLOAT değil!) |
| \`TINYINT(1)\` | Boolean flag |
| \`TIMESTAMP\` | created_at, updated_at (UTC) |
| \`JSON\` | MySQL 5.7+ yapılandırma/metadata |

---

## 6. İndeks Stratejisi

\`\`\`sql
-- Composite index: sık birlikte sorgulanan kolonlar
CREATE INDEX idx_orders_user_status ON orders (user_id, status);

-- Covering index: SELECT kolonları da index'e dahil
CREATE INDEX idx_posts_slug_title ON posts (slug, title);

-- EXPLAIN ile doğrula
EXPLAIN SELECT * FROM orders WHERE user_id = 1 AND status = 'active';
\`\`\`
`;
  writeUTF8(dbModulePath, dbContent);
  console.log('  ✅ ag-php-developer/database.md oluşturuldu');
} else {
  console.log('  ℹ️  ag-php-developer/database.md zaten mevcut, atlandı');
}

// 4. ag-php-developer SKILL.md güncelle (database modülünü tabloya ekle)
console.log('\n▶ ag-php-developer SKILL.md GÜNCELLEME KONTROL...');
const phpSkillFile = path.join(SKILLS_DIR, 'ag-php-developer', 'SKILL.md');
if (!fs.existsSync(phpSkillFile)) { console.log('  ℹ️  ag-php-developer/SKILL.md hedefte bulunamadı, atlandı'); return; }
const phpSkillContent = fs.readFileSync(phpSkillFile, { encoding: 'utf8' });
if (!phpSkillContent.includes('database.md')) {
  const updatedContent = phpSkillContent.replace(
    '| **Framework Uzmanı**',
    '| **Veritabanı** | [`database.md`](database.md) | MySQL/MariaDB, PostgreSQL, PDO prepared statement, N+1 engelleme, migration (çift adımlı), keyset pagination, indeks stratejisi. |\n| **Framework Uzmanı**'
  );
  writeUTF8(phpSkillFile, updatedContent);
  console.log('  ✅ ag-php-developer/SKILL.md → database.md satırı eklendi');
} else {
  console.log('  ℹ️  ag-php-developer/SKILL.md zaten database.md içeriyor');
}

// ─── ÖZET RAPOR ─────────────────────────────────────────────────────────────
console.log('\n══════════════════════════════════════════');
console.log(' TAMAMLANDI — ÖZET RAPOR');
console.log('══════════════════════════════════════════');
console.log(`  Agent.md güncellendi : ${Object.keys(agents).length} dosya`);
console.log(`  Yeni Skill oluşturuldu: ${Object.keys(newSkills).length} klasör`);
console.log('  ag-php-developer/database.md: eklendi');
console.log('\n  Eşleştirme:');
console.log('  frontend-developer    → ag-php-developer (frontend-developer.md)');
console.log('  backend-developer     → ag-php-developer (backend-developer.md + database.md)');
console.log('  php-security-expert   → ag-php-developer (php-security.md)');
console.log('  database-expert       → ag-database-expert + ag-php-developer/database.md');
console.log('  php-developer         → ag-php-developer (TÜM modüller)');
console.log('══════════════════════════════════════════\n');
