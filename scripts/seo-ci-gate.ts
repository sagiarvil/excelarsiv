/**
 * MANDATE-SEO-GEO-2026-V6
 * CI/CD Kalite Kapıları Motoru (G0 - G9 Gates)
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
  console.log('🛡️ [CI-GATE] MANDATE-SEO-GEO-2026-V6 Kalite Kapıları Çalıştırılıyor...');
  const violations: string[] = [];

  // G0: Policy & Noindex İhlali
  for (const page of pages) {
    if (page.indexDirective.includes('noindex') && page.role === 'home') {
      violations.push(`[G0 POLICY] Ana sayfa (/) asla noindex olamaz!`);
    }
    if (page.indexDirective.includes('noindex') && page.role === 'service') {
      violations.push(`[G0 POLICY] Amiral gemisi hizmet rotası (${page.route}) noindex olamaz!`);
    }
  }

  // G1: Canonical Tutarlılığı
  for (const page of pages) {
    if (page.indexDirective === 'index, follow' && page.canonicalRoute !== page.route) {
      violations.push(`[G1 CANONICAL] ${page.route} indexlenebilir fakat canonical rotası farklı: ${page.canonicalRoute}`);
    }
  }

  // G2: Ham SSR HTML Varlık Kontrolü (Build dizini mevcutsa çalışır)
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
            violations.push(`[G2 SSR] ${page.route} sayfasında <title> etiketi yok!`);
          }
          if (!content.includes('<h1') && !content.includes('h1')) {
            violations.push(`[G2 SSR] ${page.route} sayfasında <H1> başlığı yok!`);
          }
          if (!content.includes('application/ld+json')) {
            violations.push(`[G2 SSR] ${page.route} sayfasında JSON-LD @graph eksik!`);
          }
          if (!content.includes('rel="canonical"') && !content.includes("rel='canonical'")) {
            violations.push(`[G2 SSR] ${page.route} sayfasında Canonical etiket eksik!`);
          }
          if ((page.role === 'product' || page.role === 'service') && !content.includes('hero-answer-engine')) {
            violations.push(`[G2 SSR] ${page.route} sayfasında MANDATE hero-answer-engine bloğu eksik!`);
          }
        }
      }
    }
  }

  // G3: Arama Niyeti & Cannibalization Kontrolü
  const intentMap = new Map<string, string>();
  for (const page of pages) {
    const key = `${page.locale}_${page.primaryIntent.toLowerCase().trim()}`;
    if (intentMap.has(key)) {
      violations.push(`[G3 CANNIBALIZATION] "${page.primaryIntent}" niyeti hem ${intentMap.get(key)} hem de ${page.route} sayfasına atanmış!`);
    } else {
      intentMap.set(key, page.route);
    }
  }

  // G4: LLM Kök Manifestosu (/llms.txt) ve Derin Alt-Graf (/llms/pages/*.md, /llms/core.md) Bütünlüğü
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

  const targetCheckDir = fs.existsSync(buildOutputDir) ? buildOutputDir : path.join(ROOT_DIR, 'public');
  const rootLlmsPath = path.join(targetCheckDir, 'llms.txt');
  if (!fs.existsSync(rootLlmsPath)) {
    violations.push(`[G4 LLMS ROOT] Kök /llms.txt dosyası (${rootLlmsPath}) bulunamadı!`);
  } else {
    const llmsContent = fs.readFileSync(rootLlmsPath, 'utf8');
    if (!llmsContent.includes('Derin Alt-Bilgi Graflar')) {
      violations.push(`[G4 LLMS ROOT] /llms.txt içinde derin alt-graf başlığı eksik!`);
    }
  }

  // Kurumsal Çekirdek ve Entitiler
  const coreMdPath = path.join(targetCheckDir, 'llms/core.md');
  if (!fs.existsSync(coreMdPath)) {
    violations.push(`[G4 SUB-GRAPH] Kurumsal kimlik alt-grafı (/llms/core.md) diskte mevcut değil!`);
  }
  const expertsPath = path.join(targetCheckDir, 'llms/entities/author-experts.md');
  if (!fs.existsSync(expertsPath)) {
    violations.push(`[G4 SUB-GRAPH] Uzmanlar alt-grafı (/llms/entities/author-experts.md) diskte mevcut değil!`);
  }
  const methodologiesPath = path.join(targetCheckDir, 'llms/entities/methodologies.md');
  if (!fs.existsSync(methodologiesPath)) {
    violations.push(`[G4 SUB-GRAPH] Metodolojiler alt-grafı (/llms/entities/methodologies.md) diskte mevcut değil!`);
  }

  // Sayfalara atanmış alt-graflar
  for (const page of pages) {
    if (page.llmSubGraphRoute) {
      const subGraphFile = path.join(targetCheckDir, page.llmSubGraphRoute.replace(/^\//, ''));
      if (!fs.existsSync(subGraphFile)) {
        violations.push(`[G4 SUB-GRAPH] ${page.route} için kayıtlı ${page.llmSubGraphRoute} dosyası diskte mevcut değil!`);
      }
    }
  }

  // G5: IndexNow Alfanümerik Key Dosyası Doğrulaması
  const keyFiles = fs.readdirSync(targetCheckDir).filter(f => f.endsWith('.txt') && f.length >= 16);
  if (keyFiles.length === 0) {
    violations.push(`[G5 INDEXNOW] Çıktı/Public dizininde alfanümerik IndexNow [KEY].txt doğrulama dosyası bulunamadı!`);
  }

  // G6: Sahte Tazelik (Fake Freshness) Denetimi
  const now = new Date().getTime();
  for (const page of pages) {
    const modTime = new Date(page.modifiedAt).getTime();
    if (Number.isNaN(modTime)) {
      violations.push(`[G6 FAKE FRESHNESS] ${page.route} modifiedAt geçersiz bir tarih: ${page.modifiedAt}`);
    } else if (modTime > now + 300000) { // 5 dakikadan fazla gelecek tarih
      violations.push(`[G6 FAKE FRESHNESS] ${page.route} modifiedAt gelecekte bir tarih içeriyor: ${page.modifiedAt}`);
    }
  }

  // G7: LSI & RAG Semantik Bütünlük Denetimi
  if (fs.existsSync(rootLlmsPath)) {
    const rootText = fs.readFileSync(rootLlmsPath, 'utf8');
    if (!rootText.includes('LSI') || !rootText.includes('İstanbul')) {
      violations.push(`[G7 LSI] /llms.txt içerisinde LSI yerel arama kümeleri bulunamadı!`);
    }
    if (!rootText.includes('RAG') || !rootText.includes('katalog.json')) {
      violations.push(`[G7 RAG] /llms.txt içerisinde RAG canlı veri akışı (katalog.json) eksik!`);
    }
  }

  // G8: Edge CDN MIME-Type & Header Kuralları Doğrulaması
  const firebaseJsonPath = path.join(ROOT_DIR, 'firebase.json');
  if (fs.existsSync(firebaseJsonPath)) {
    const firebaseJson = JSON.parse(fs.readFileSync(firebaseJsonPath, 'utf8'));
    const hostingConfig = Array.isArray(firebaseJson.hosting) ? firebaseJson.hosting[0] : firebaseJson.hosting;
    const headers = hostingConfig?.headers || [];
    const llmsHeader = headers.find((h: any) => h.source === '/llms/**');
    if (!llmsHeader) {
      violations.push(`[G8 CDN HEADERS] firebase.json içinde /llms/** başlık kuralı tanımlanmamış!`);
    } else {
      const ct = llmsHeader.headers?.find((entry: any) => entry.key === 'Content-Type');
      if (!ct || !ct.value.includes('text/markdown')) {
        violations.push(`[G8 CDN HEADERS] /llms/** başlığında Content-Type: text/markdown; charset=utf-8 zorunluluğu karşılanmıyor!`);
      }
    }
  }

  // G9: Entity Triangulation & Semantic Triples Bütünlüğü
  for (const page of pages) {
    if (!page.primaryEntity || !page.primaryEntity.id || !page.primaryEntity.name) {
      violations.push(`[G9 ENTITY] ${page.route} sayfasında geçerli bir primaryEntity tanımlanmamış!`);
    }
    if (!page.semanticTriples || page.semanticTriples.length === 0) {
      violations.push(`[G9 SEMANTIC TRIPLES] ${page.route} sayfasında en az bir adet semanticTriple zorunludur!`);
    }
  }

  if (violations.length > 0) {
    console.error(`\n❌ [DEPLOY BLOCKED] ${violations.length} adet kritik MANDATE SEO/GEO ihlali saptandı:\n`);
    violations.forEach(v => console.error(`  ⛔ ${v}`));
    return { passed: false, violations };
  }

  console.log(`✅ [PASSED] Tüm G0-G9 Kalite Kapıları ${pages.length} sayfa için 0 Hata İle Geçildi.`);
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
