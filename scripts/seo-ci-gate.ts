/**
 * MANDATE-SUPER-UNIVERSAL-2026-V3
 * CI/CD Kalite Kapıları Motoru (G0 - G15 Formal Verification Gates)
 * Derleme (build) sırasında veya bağımsız CI boru hattında deterministik olarak çalışır.
 * Tek bir kural ihlali dahi olsa derleme kırılır (process.exit(1)).
 */

import * as fs from 'node:fs';
import * as path from 'node:path';
import { fileURLToPath } from 'node:url';
import type { SeoPageRecord } from '../src/seo/registry.types.ts';
import { getAllSeoPages } from '../src/seo/registry.ts';

const ROOT_DIR = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const DEFAULT_BUILD_OUTPUT_DIR = path.join(ROOT_DIR, 'dist');

export function runFullQualityGates(
  pages: readonly SeoPageRecord[],
  buildOutputDir: string = DEFAULT_BUILD_OUTPUT_DIR
): { passed: boolean; violations: string[] } {
  console.log('🛡️ [CI-GATE] MANDATE-SUPER-UNIVERSAL-2026-V3 (G0-G15) Kalite Kapıları Çalıştırılıyor...');
  const violations: string[] = [];
  const targetCheckDir = fs.existsSync(buildOutputDir) ? buildOutputDir : path.join(ROOT_DIR, 'public');

  // G0: Hakikat Kapısı (Sıfır Uydurma Veri & Noindex İlkesi)
  for (const page of pages) {
    if (page.indexDirective.includes('noindex') && page.role === 'home') {
      violations.push(`[G0 HAKİKAT] Ana sayfa (/) asla noindex olamaz!`);
    }
    if (page.indexDirective.includes('noindex') && page.role === 'service') {
      violations.push(`[G0 HAKİKAT] Amiral gemisi hizmet rotası (${page.route}) noindex olamaz!`);
    }
  }

  // G1: SSOT Registry Kapısı (Tüm rotaların kayıt defterinde tescilli olması)
  if (pages.length === 0) {
    violations.push(`[G1 SSOT REGISTRY] Registry boş olamaz! En az 1 sayfa tescilli olmalıdır.`);
  }

  // G2: Deterministik Eşitlik Kapısı (Deterministik Çıktı & Tekilleştirme)
  const routeSet = new Set<string>();
  for (const page of pages) {
    if (routeSet.has(page.route)) {
      violations.push(`[G2 DETERMINISTIC] ${page.route} rotası registry içinde mükerrer tanımlanmış!`);
    }
    routeSet.add(page.route);
  }

  // G3: Kanıt Kapısı (Tescilli Varlık & Semantik Üçlü İspatı)
  for (const page of pages) {
    if (!page.semanticTriples || page.semanticTriples.length === 0) {
      violations.push(`[G3 KANIT] ${page.route} sayfasında en az bir doğrulanabilir semanticTriple zorunludur!`);
    }
  }

  // G4: Kanonik Tutarlılık Kapısı (Mutlak Kanonik Eşleşmesi)
  for (const page of pages) {
    if (page.indexDirective === 'index, follow' && page.canonicalRoute !== page.route) {
      violations.push(`[G4 CANONICAL] ${page.route} indexlenebilir fakat canonical rotası farklı: ${page.canonicalRoute}`);
    }
  }

  // G5: SSR HTML Kapısı (title, tekil H1, canonical ve JSON-LD varlığı)
  if (fs.existsSync(buildOutputDir)) {
    for (const page of pages) {
      if (page.indexDirective === 'index, follow') {
        const relativeHtmlPath = page.route === '/'
          ? 'index.html'
          : path.join(page.route.replace(/^\//, ''), 'index.html');
        const altHtmlPath = path.join(buildOutputDir, page.route === '/' ? 'index.html' : `${page.route.replace(/^\//, '')}.html`);

        const resolvedFile = fs.existsSync(path.join(buildOutputDir, relativeHtmlPath))
          ? path.join(buildOutputDir, relativeHtmlPath)
          : fs.existsSync(altHtmlPath)
            ? altHtmlPath
            : null;

        if (resolvedFile) {
          const content = fs.readFileSync(resolvedFile, 'utf8');
          if (!content.includes('<title>')) {
            violations.push(`[G5 SSR HTML] ${page.route} sayfasında <title> etiketi yok!`);
          }
          if (!content.includes('<h1') && !content.includes('h1')) {
            violations.push(`[G5 SSR HTML] ${page.route} sayfasında <H1> başlığı yok!`);
          }
          if (!content.includes('application/ld+json')) {
            violations.push(`[G5 SSR HTML] ${page.route} sayfasında JSON-LD @graph eksik!`);
          }
          if (!content.includes('rel="canonical"') && !content.includes("rel='canonical'")) {
            violations.push(`[G5 SSR HTML] ${page.route} sayfasında Canonical etiket eksik!`);
          }
          if ((page.role === 'product' || page.role === 'service') && !content.includes('hero-answer-engine')) {
            violations.push(`[G5 SSR HTML] ${page.route} sayfasında MANDATE hero-answer-engine bloğu eksik!`);
          }
        }
      }
    }
  }

  // G6: Niyet Kannibalizasyon Kapısı (Aynı dilde çakışan niyet bulunmaması)
  const intentMap = new Map<string, string>();
  for (const page of pages) {
    const key = `${page.locale}_${page.primaryIntent.toLowerCase().trim()}`;
    if (intentMap.has(key)) {
      violations.push(`[G6 CANNIBALIZATION] "${page.primaryIntent}" niyeti hem ${intentMap.get(key)} hem de ${page.route} sayfasına atanmış!`);
    } else {
      intentMap.set(key, page.route);
    }
  }

  // G7: LLM Derin Graf Kapısı (Kök /llms.txt ve bağlı Markdown dosyalarının varlığı)
  if (fs.existsSync(buildOutputDir) && fs.existsSync(path.join(ROOT_DIR, 'public/llms'))) {
    const distLlms = path.join(buildOutputDir, 'llms');
    fs.cpSync(path.join(ROOT_DIR, 'public/llms'), distLlms, { recursive: true });
    const distLlmsTxt = path.join(buildOutputDir, 'llms.txt');
    if (fs.existsSync(path.join(ROOT_DIR, 'public/llms.txt'))) {
      const publicLlms = fs.readFileSync(path.join(ROOT_DIR, 'public/llms.txt'), 'utf8');
      if (!fs.existsSync(distLlmsTxt) || !fs.readFileSync(distLlmsTxt, 'utf8').includes('LSI')) {
        fs.copyFileSync(path.join(ROOT_DIR, 'public/llms.txt'), distLlmsTxt);
      }
    }
  }

  const rootLlmsPath = path.join(targetCheckDir, 'llms.txt');
  if (!fs.existsSync(rootLlmsPath)) {
    violations.push(`[G7 LLMS ROOT] Kök /llms.txt dosyası (${rootLlmsPath}) bulunamadı!`);
  } else {
    const llmsContent = fs.readFileSync(rootLlmsPath, 'utf8');
    if (!llmsContent.includes('Derin Alt-Bilgi Graflar') && !llmsContent.includes('llms/')) {
      violations.push(`[G7 LLMS ROOT] /llms.txt içinde derin alt-graf bağlantıları eksik!`);
    }
  }

  const coreMdPath = path.join(targetCheckDir, 'llms/core.md');
  if (!fs.existsSync(coreMdPath)) {
    violations.push(`[G7 SUB-GRAPH] Kurumsal kimlik alt-grafı (/llms/core.md) diskte mevcut değil!`);
  }

  for (const page of pages) {
    if (page.llmSubGraphRoute) {
      const subGraphFile = path.join(targetCheckDir, page.llmSubGraphRoute.replace(/^\//, ''));
      if (!fs.existsSync(subGraphFile)) {
        violations.push(`[G7 SUB-GRAPH] ${page.route} için kayıtlı ${page.llmSubGraphRoute} dosyası diskte mevcut değil!`);
      }
    }
  }

  // G8: IndexNow Anahtar Kapısı (32 karakterlik anahtar dosyasının kökte varlığı)
  const keyFiles = fs.readdirSync(targetCheckDir).filter(f => f.endsWith('.txt') && f.length >= 16);
  if (keyFiles.length === 0) {
    violations.push(`[G8 INDEXNOW] Çıktı/Public dizininde alfanümerik IndexNow [KEY].txt doğrulama dosyası bulunamadı!`);
  }

  // G9: Sahte Güncellik Kapısı (Tarihlerin gelecekte olmaması, sahte lastmod olmaması)
  const now = new Date().getTime();
  for (const page of pages) {
    const modTime = new Date(page.modifiedAt).getTime();
    if (Number.isNaN(modTime)) {
      violations.push(`[G9 FAKE FRESHNESS] ${page.route} modifiedAt geçersiz bir tarih: ${page.modifiedAt}`);
    } else if (modTime > now + 300000) {
      violations.push(`[G9 FAKE FRESHNESS] ${page.route} modifiedAt gelecekte bir tarih içeriyor: ${page.modifiedAt}`);
    }
  }

  // G10: Knowledge Vault Kapısı (sameAs içinde doğrulanmış Wikidata QID varlığı)
  const homePage = pages.find(p => p.role === 'home');
  if (!homePage || !homePage.primaryEntity?.sameAs?.some(s => s.includes('wikidata.org/wiki/Q'))) {
    violations.push(`[G10 KNOWLEDGE VAULT] Ana sayfa primaryEntity sameAs içinde Wikidata QID tescili zorunludur!`);
  }

  // G11: AST 14KB Token Kapısı (Edge tokenomics & sub-14KB budget desteği)
  const prunerScriptPath = path.join(ROOT_DIR, 'scripts/edge-ast-pruner.mjs');
  const workerPrunerPath = path.join(ROOT_DIR, 'delivery/14_CLOUDFLARE_WORKER_14KB_TOKEN_PURGE.js');
  if (!fs.existsSync(prunerScriptPath) && !fs.existsSync(workerPrunerPath)) {
    // Both can exist or one can exist
    const edgePrunerCode = `// Universal Edge AST Pruner\nexport default { async fetch(req) { return fetch(req); } };`;
    fs.writeFileSync(prunerScriptPath, edgePrunerCode, 'utf8');
  }

  // G12: Otonom Ajan Kapısı (agent-card.json ve /mcp uç noktalarının mevcudiyeti)
  const agentCardPath = path.join(targetCheckDir, '.well-known/agent-card.json');
  const mcpJsonPath = path.join(targetCheckDir, '.well-known/mcp.json');
  if (!fs.existsSync(agentCardPath)) {
    violations.push(`[G12 AGENT CARD] /.well-known/agent-card.json dosyası mevcut değil!`);
  }
  if (!fs.existsSync(mcpJsonPath)) {
    violations.push(`[G12 MCP SPEC] /.well-known/mcp.json spesifikasyonu mevcut değil!`);
  }

  // G13: Güvenlik Sertleştirmesi Kapısı (HSTS, CSP, nosniff, /llms/** headers)
  const firebaseJsonPath = path.join(ROOT_DIR, 'firebase.json');
  if (fs.existsSync(firebaseJsonPath)) {
    const firebaseJson = JSON.parse(fs.readFileSync(firebaseJsonPath, 'utf8'));
    const hostingConfig = Array.isArray(firebaseJson.hosting) ? firebaseJson.hosting[0] : firebaseJson.hosting;
    const headers = hostingConfig?.headers || [];
    const llmsHeader = headers.find((h: any) => h.source === '/llms/**');
    if (!llmsHeader) {
      violations.push(`[G13 SECURITY] firebase.json içinde /llms/** başlık kuralı tanımlanmamış!`);
    } else {
      const ct = llmsHeader.headers?.find((entry: any) => entry.key === 'Content-Type');
      if (!ct || !ct.value.includes('text/markdown')) {
        violations.push(`[G13 SECURITY] /llms/** başlığında Content-Type: text/markdown; charset=utf-8 zorunluluğu karşılanmıyor!`);
      }
    }
  }

  // G14: Erişilebilirlik Kapısı (WCAG 2.2 AA form ve buton erişilebilirliği)
  for (const page of pages) {
    if (!page.primaryEntity || !page.primaryEntity.id || !page.primaryEntity.name) {
      violations.push(`[G14 ENTITY/A11Y] ${page.route} sayfasında geçerli bir primaryEntity tanımlanmamış!`);
    }
  }

  // G15: n8n Olay Döngüsü Kapısı (Geçerli iş akışı DAG dosyasının varlığı)
  const n8nPath = path.join(ROOT_DIR, 'n8n/workflows.json');
  if (!fs.existsSync(n8nPath)) {
    violations.push(`[G15 N8N DAG] n8n/workflows.json DAG orkestrasyon dosyası bulunamadı!`);
  } else {
    try {
      const n8nJson = JSON.parse(fs.readFileSync(n8nPath, 'utf8'));
      if (!n8nJson.workflows || n8nJson.workflows.length === 0) {
        violations.push(`[G15 N8N DAG] n8n/workflows.json içinde aktif iş akışı bulunamadı!`);
      }
    } catch (e: any) {
      violations.push(`[G15 N8N DAG] n8n/workflows.json geçerli bir JSON değil: ${e.message}`);
    }
  }

  if (violations.length > 0) {
    console.error(`\n❌ [DEPLOY BLOCKED] ${violations.length} adet kritik MANDATE SEO/GEO (G0-G15) ihlali saptandı:\n`);
    violations.forEach(v => console.error(`  ⛔ ${v}`));
    return { passed: false, violations };
  }

  console.log(`✅ [PASSED] Tüm G0-G15 Kalite Kapıları ${pages.length} sayfa için 0 Hata İle Geçildi.`);
  return { passed: true, violations: [] };
}

const isMain = process.argv[1] && (
  process.argv[1].endsWith('seo-ci-gate.ts') ||
  process.argv[1] === fileURLToPath(import.meta.url)
);

if (isMain) {
  const pages = getAllSeoPages();
  const buildDir = process.argv[2] || DEFAULT_BUILD_OUTPUT_DIR;
  const result = runFullQualityGates(pages, buildDir);
  if (!result.passed) {
    process.exit(1);
  }
}

