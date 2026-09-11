const fs = require('fs');
const path = require('path');

const GEMINI  = '$HOME/.gemini';
const DESKTOP = '$HOME/Desktop/gemini/.gemini';

function mkdirp(p) { fs.mkdirSync(p, { recursive: true }); }
function writeUtf8(p, c) { mkdirp(path.dirname(p)); fs.writeFileSync(p, c, 'utf8'); }
function copyFile(src, dst) { mkdirp(path.dirname(dst)); fs.copyFileSync(src, dst); }

// ============================================================
// 1. ag-schema-expert SKILL.md — .gemini ve Desktop'a yaz
// ============================================================
const schemaSkill = [
  '---',
  'name: ag-schema-expert',
  'description: Yapısal veri uzmanı. Schema.org / JSON-LD (Product, WebSite, BreadcrumbList, Article, FAQPage).',
  '---',
  '',
  '# AG-SCHEMA-EXPERT — JSON-LD Yapısal Veri Uzmanı',
  '',
  'Schema.org standartlarına ve Google Rich Results yönergelerine tam uyumlu JSON-LD üretir ve denetler.',
  '',
  '## Desteklenen Şema Tipleri',
  '- **E-ticaret:** Product, Offer, AggregateRating, Review',
  '- **İçerik:** Article, BlogPosting, NewsArticle, HowTo, Recipe, FAQPage',
  '- **Organizasyon:** Organization, LocalBusiness, WebSite, WebPage',
  '- **Navigasyon:** BreadcrumbList, SiteLinksSearchBox',
  '- **Etkinlik:** Event, Course, JobPosting',
  '',
  '## Temel Kurallar',
  '',
  '1. **@graph container:** Tüm şema tipleri sayfa başına tek `@graph` bloğunda birleştirilir.',
  '2. **aggregateRating:** YALNIZCA gerçek verilerle; sahte/rastgele puan üretmek KESİNLİKLE YASAKTIR.',
  '3. **Güvenli output:** JSON-LD içindeki `</script>` kapanma etiketleri `<\\/script>` olarak kaçılır.',
  '4. **Knowledge Vault:** Organization şemasına `sameAs` içinde Wikidata QID eklenir.',
  '5. **Doğrulama:** Her şema için Google Rich Results Test URL\'si verilir.',
  '6. **@id kullanımı:** Sayfa içi çapraz referanslar `@id` ile yapılır; inline tekrar edilmez.',
  '7. **Sayfa başına bir blok:** Birden fazla JSON-LD script etiketi yerine tek `@graph` array kullanılır.',
  '',
  '## Çıktı Doğrulama',
  '- Google Rich Results Test: https://search.google.com/test/rich-results',
  '- Schema.org Validator: https://validator.schema.org',
].join('\n');

const schemaPaths = [
  GEMINI  + '/config/skills/ag-schema-expert/SKILL.md',
  DESKTOP + '/config/skills/ag-schema-expert/SKILL.md',
];
for (const p of schemaPaths) {
  writeUtf8(p, schemaSkill);
  console.log('[OK] ' + p.replace(/.*\.gemini/, '.gemini'));
}

// ============================================================
// 2. ag-seo-structure-tester SKILL.md
// ============================================================
const seoStructSkill = [
  '---',
  'name: ag-seo-structure-tester',
  'description: Sayfa SEO test uzmanı. H1-H6 hiyerarşisi, landmark etiketleri ve JSON-LD yapı denetimi.',
  '---',
  '',
  '# AG-SEO-STRUCTURE-TESTER — Sayfa İçi SEO Yapı Denetçisi',
  '',
  'Bir sayfanın semantik yapısını, başlık hiyerarşisini ve JSON-LD doğruluğunu denetler.',
  '> **Not:** Bu ajan yalnızca RAPORLAR. Kod değişikliği yapmaz, öneri sunmaz.',
  '',
  '## Denetim Alanları',
  '',
  '### 1. Başlık Hiyerarşisi (ENG-04 / ENG-14)',
  '- Tek H1 zorunlu — birden fazla H1 → KRİTİK HATA',
  '- Seviye atlama yasak: H1→H3 (H2 atlanırsa HATA), H2→H5 (HATA)',
  '- H1 sayfanın birincil anahtar kelimesini içermeli',
  '- Boş heading (içeriksiz `<h*></h*>`) → KRİTİK HATA',
  '- Footer veya header içinde H1 → KRİTİK HATA',
  '',
  '### 2. Landmark Etiketleri (ENG-11)',
  '- `<main>`, `<header>`, `<footer>`, `<nav>` doğru ve tekil kullanım',
  '- `<main>` içinde H1 bulunmalı',
  '- `<aside>` varsa anlamsal içerik taşımalı',
  '',
  '### 3. JSON-LD Yapı (ENG-08 / ENG-15)',
  '- Geçerli `@context` ve `@type` varlığı',
  '- `@graph` container kullanımı tercih edilir',
  '- Çözülemeyen `@id` referansları — UYARI',
  '- Google Rich Results Test uyumluluğu',
  '',
  '## Çıktı Formatı (Kesin — Değiştirilemez)',
  '```',
  '✅ GEÇTİ: [madde]',
  '❌ KALDI: [madde]',
  '⚠️ UYARI: [madde]',
  '```',
].join('\n');

const seoStructPaths = [
  GEMINI  + '/config/skills/ag-seo-structure-tester/SKILL.md',
  DESKTOP + '/config/skills/ag-seo-structure-tester/SKILL.md',
];
for (const p of seoStructPaths) {
  writeUtf8(p, seoStructSkill);
  console.log('[OK] ' + p.replace(/.*\.gemini/, '.gemini'));
}

// ============================================================
// 3. Agent ↔ Skill harita doğrulaması — tüm agent.md'leri oku
//    ve skills: alanının doğru olduğunu kontrol et
// ============================================================
console.log('\n=== Agent ↔ Skill Harita Kontrolü ===');

const agentSkillMap = {
  'accessibility-expert':     'ag-accessibility-expert',
  'backend-developer':        'ag-php-developer (module: backend-developer.md + database.md)',
  'bug-hunter':               'ag-bug-hunter',
  'content-writer':           'ag-content-writer',
  'database-expert':          'ag-database-expert + ag-php-developer (module: database.md)',
  'devops-engineer':          'ag-devops-engineer',
  'frontend-developer':       'ag-php-developer (module: frontend-developer.md)',
  'html-export-expert':       'ag-html-export-expert',
  'mobile-optimization-expert': 'ag-mobile-optimization-expert',
  'php-developer':            'ag-php-developer (all modules)',
  'php-security-expert':      'ag-php-developer (module: php-security.md)',
  'project-manager':          'ag-project-manager',
  'schema-expert':            'ag-schema-expert',
  'seo-expert':               'ag-seo-expert',
  'seo-structure-tester':     'ag-seo-structure-tester',
  'technical-writer':         'ag-technical-writer',
  'test-engineer':            'ag-test-engineer',
};

const agentsDir = GEMINI + '/config/agents';
for (const [agName, expectedSkill] of Object.entries(agentSkillMap)) {
  const agPath = path.join(agentsDir, agName, 'agent.md');
  if (!fs.existsSync(agPath)) {
    console.log('[EKSIK] agent/' + agName + '/agent.md — dosya yok!');
    continue;
  }
  const content = fs.readFileSync(agPath, 'utf8');
  const skillMatch = content.match(/^skills:\s*(.+)$/m);
  const currentSkill = skillMatch ? skillMatch[1].trim() : '(yok)';
  const ok = currentSkill.toLowerCase().includes(expectedSkill.split(' ')[0].toLowerCase());
  console.log((ok ? '[OK] ' : '[UYARI] ') + agName + ' → skills: ' + currentSkill);
}

// ============================================================
// 4. Desktop rules/ → .gemini/config/rules/ senkron kontrol
// ============================================================
console.log('\n=== Rules Senkron Kontrolü ===');
const rulesDir = GEMINI + '/config/rules';
const desktopRulesDir = DESKTOP + '/config/rules';
const ruleFiles = fs.readdirSync(rulesDir);
for (const rf of ruleFiles) {
  const dstPath = path.join(desktopRulesDir, rf);
  if (!fs.existsSync(dstPath)) {
    copyFile(path.join(rulesDir, rf), dstPath);
    console.log('[KOPYALANDI] rules/' + rf + ' -> Desktop');
  } else {
    console.log('[OK] rules/' + rf + ' Desktop\'ta mevcut');
  }
}

// ============================================================
// 5. Son özet
// ============================================================
console.log('\n=== SON DURUM ===');
console.log('.gemini/config/agents: ' + fs.readdirSync(agentsDir).length + ' agent klasörü');
console.log('.gemini/config/skills: ' + fs.readdirSync(GEMINI+'/config/skills').length + ' skill');
console.log('.gemini/config/rules:  ' + fs.readdirSync(rulesDir).length + ' kural dosyası');
console.log('Desktop/config/skills: ' + fs.readdirSync(DESKTOP+'/config/skills').length + ' skill');
console.log('Desktop/config/rules:  ' + fs.readdirSync(desktopRulesDir).length + ' kural dosyası');
console.log('\n[TAMAMLANDI]');
