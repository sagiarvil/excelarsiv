/**
 * ExcelArşiv Enterprise Traffic & Revenue Engine Deterministic Test Suite
 * Iron Law of Verification PASS Checker
 */

import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { readFileSync, existsSync } from 'node:fs';
import {
  buildIndexNowPayload,
  analyzeSemanticDensity,
  generateHighIntentTrafficMatrix,
  INDEXNOW_ENDPOINTS
} from './traffic-revenue-engine.mjs';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const rootDir = resolve(__dirname, '..');

let passCount = 0;
let failCount = 0;

function assert(condition, message) {
  if (condition) {
    passCount++;
    console.log(`  ✅ PASS: ${message}`);
  } else {
    failCount++;
    console.error(`  ❌ FAIL: ${message}`);
  }
}

async function runTests() {
  console.log('🚀 [Traffic & Revenue Engine Test] Başlatılıyor...');

  // Test 1: IndexNow Endpoints
  console.log('\n--- Test 1: IndexNow Multi-Hub Endpoints ---');
  assert(INDEXNOW_ENDPOINTS.length >= 3, 'En az 3 IndexNow uç noktası mevcut');
  assert(INDEXNOW_ENDPOINTS.includes('https://api.indexnow.org/indexnow'), 'IndexNow.org endpoint tanımlı');
  assert(INDEXNOW_ENDPOINTS.includes('https://www.bing.com/indexnow'), 'Bing endpoint tanımlı');
  assert(INDEXNOW_ENDPOINTS.includes('https://yandex.com/indexnow'), 'Yandex endpoint tanımlı');

  // Test 2: IndexNow Payload Generator
  console.log('\n--- Test 2: IndexNow Payload Builder ---');
  const sampleUrls = [
    'https://excelarsiv.com/sablonlar/',
    'https://excelarsiv.com/sablon/13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi',
    '/sablon/akilli-kasa-defteri-ve-nakit-kontrol-sistemi'
  ];
  const payload = buildIndexNowPayload({
    host: 'excelarsiv.com',
    key: 'testkey123',
    urlList: sampleUrls
  });
  assert(payload.host === 'excelarsiv.com', 'Payload host doğru');
  assert(payload.key === 'testkey123', 'Payload key doğru');
  assert(payload.urlList.length === 3, 'Tüm URLler normalize edildi ve korundu');
  assert(payload.urlList.every((u) => u.startsWith('https://excelarsiv.com')), 'Tüm URLler tam canonical https://');

  // Test 3: Semantik Yoğunluk ve 14KB AST Byte Gate Analizi
  console.log('\n--- Test 3: Semantik Yoğunluk & AST Byte Gate ---');
  const sampleValidHtml = `
    <!DOCTYPE html>
    <html lang="tr">
    <head>
      <title>13 Haftalık Nakit Akışı - Excel Arşiv</title>
      <link rel="alternate" type="application/rss+xml" href="https://excelarsiv.com/rss.xml" />
      <link rel="alternate" type="application/atom+xml" href="https://excelarsiv.com/atom.xml" />
      <link rel="help" href="https://excelarsiv.com/llms.txt" />
      <script type="application/ld+json">{"@context":"https://schema.org","@type":"Product","name":"Nakit Akışı"}</script>
    </head>
    <body>
      <h1>13 Haftalık Nakit Akışı Modeli</h1>
      <p>Kurumsal nakit projeksiyonu ve stres testi.</p>
    </body>
    </html>
  `;
  const semResult = analyzeSemanticDensity(sampleValidHtml);
  assert(semResult.passed === true, 'Standart semantik HTML doğrulamasından başarıyla geçti');
  assert(semResult.first14KBHasEntity === true, 'İlk 14KB AST bütçesinde varlık tespit edildi');
  assert(semResult.h1Count === 1, 'Tekil H1 kuralı sağlandı');
  assert(semResult.discovery.hasRssRef === true, 'RSS autodiscovery linki doğrulandı');
  assert(semResult.discovery.hasAtomRef === true, 'Atom autodiscovery linki doğrulandı');

  // Test 4: Yüksek Ticari Niyet Matrisi (Revenue Routing)
  console.log('\n--- Test 4: High Intent Traffic & Revenue Matrix ---');
  let catalog = [];
  try {
    const raw = readFileSync(resolve(rootDir, 'public', 'katalog.json'), 'utf8');
    const parsed = JSON.parse(raw);
    catalog = Array.isArray(parsed) ? parsed : (parsed.products || []);
  } catch (err) {
    console.error('Katalog okunamadı:', err.message);
  }

  assert(catalog.length > 0, `Katalog yüklendi (toplam ${catalog.length} ürün)`);
  const matrix = generateHighIntentTrafficMatrix(catalog);
  assert(matrix.pillars.length === 4, '4 kritik gelir sütunu hesaplandı');
  assert(matrix.pillars[0].productCount > 0, 'Nakit akışı kategorisi ürünlerle eşleşti');
  assert(matrix.trafficRecommendations.length >= 3, 'Trafik ve büyüme strateji önerileri üretildi');

  console.log(`\n========================================`);
  console.log(`📊 TRAFFIC & REVENUE ENGINE TEST SONUCU:`);
  console.log(`   Toplam Geçen: ${passCount}`);
  console.log(`   Toplam Kalan: ${failCount}`);
  console.log(`========================================`);

  if (failCount > 0) {
    process.exit(1);
  }
}

runTests().catch((err) => {
  console.error('Test hatası:', err);
  process.exit(1);
});
