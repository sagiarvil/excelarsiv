/**
 * ExcelArşiv Enterprise Traffic & Revenue Engine v2.0.0
 * 
 * Silikon Vadisi ve Londra standartlarında ($5M+ Tier) otonom arama,
 * indeksleme ve trafik telemetrisi motoru.
 * 
 * Bileşenler:
 * 1. IndexNow Multi-Hub Broadcaster (Bing, Yandex, IndexNow API)
 * 2. LLM AEO/GEO Telemetry Probe & Semantic Density Scorer
 * 3. Autonomous Catalog Traffic Vector Matrix (High Intent Routing)
 */

import { readFileSync, existsSync, readdirSync, statSync } from 'node:fs';
import { resolve, dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const rootDir = resolve(__dirname, '..');

export const INDEXNOW_ENDPOINTS = [
  'https://api.indexnow.org/indexnow',
  'https://www.bing.com/indexnow',
  'https://yandex.com/indexnow'
];

/**
 * 1. IndexNow Çok Merkezli Protokol Yükü Hazırlayıcı
 * @param {Object} options
 * @param {string} options.host - 'excelarsiv.com'
 * @param {string} options.key - IndexNow API Key
 * @param {string} [options.keyLocation] - Key dosya konumu
 * @param {Array<string>} options.urlList - İletilecek canonical URL dizisi
 */
export function buildIndexNowPayload(options) {
  const {
    host = 'excelarsiv.com',
    key = 'c90538a719604eb6b8c8d887a0572e9a',
    keyLocation = `https://${host}/c90538a719604eb6b8c8d887a0572e9a.txt`,
    urlList = []
  } = options;

  if (!urlList.length) {
    throw new Error('IndexNow bildirimi için en az bir URL gereklidir.');
  }

  // URL'lerin host ile tam uyuştuğunu ve canonical olduğunu garanti et
  const sanitizedUrls = urlList.map((u) => {
    try {
      const parsed = new URL(u);
      return parsed.origin === `https://${host}` ? u : `https://${host}${parsed.pathname}`;
    } catch {
      return u.startsWith('/') ? `https://${host}${u}` : `https://${host}/${u}`;
    }
  });

  return {
    host,
    key,
    keyLocation,
    urlList: [...new Set(sanitizedUrls)]
  };
}

/**
 * 2. Semantik Yoğunluk ve AEO Yanıt Bütçesi Denetleyicisi (ColBERT Late-Interaction Uyumu)
 * @param {string} htmlContent - Taranacak HTML içeriği
 */
export function analyzeSemanticDensity(htmlContent) {
  if (!htmlContent || typeof htmlContent !== 'string') {
    return {
      score: 0,
      passed: false,
      reasons: ['HTML içeriği boş veya geçersiz.']
    };
  }

  const reasons = [];
  let score = 100;

  // 14KB AST Byte Gate Kontrolü
  const first14KB = htmlContent.slice(0, 14336);
  const hasEntityInFirst14KB = /Schema\.org|application\/ld\+json|<h1/i.test(first14KB);
  if (!hasEntityInFirst14KB) {
    score -= 25;
    reasons.push('İlk 14KB TCP/TLS paketinde Schema.org veya H1 tespit edilemedi.');
  }

  // Schema.org @graph Doğrulaması
  const hasJsonLd = /<script\s+type=["']application\/ld\+json["']/i.test(htmlContent);
  if (!hasJsonLd) {
    score -= 30;
    reasons.push('JSON-LD yapısal veri bloğu eksik.');
  }

  // Tekil H1 Kuralı
  const h1Matches = htmlContent.match(/<h1[\s>]/gi) || [];
  if (h1Matches.length === 0) {
    score -= 20;
    reasons.push('H1 başlığı bulunamadı.');
  } else if (h1Matches.length > 1) {
    score -= 15;
    reasons.push(`Birden fazla H1 (${h1Matches.length}) tespit edildi.`);
  }

  // LLM Discovery Link Kontrolleri
  const hasLlmsRef = /href=["'][^"']*llms\.txt["']/i.test(htmlContent);
  const hasRssRef = /type=["']application\/rss\+xml["']/i.test(htmlContent);
  const hasAtomRef = /type=["']application\/atom\+xml["']/i.test(htmlContent);

  return {
    score: Math.max(0, score),
    passed: score >= 75,
    first14KBHasEntity: hasEntityInFirst14KB,
    hasJsonLd,
    h1Count: h1Matches.length,
    discovery: {
      hasLlmsRef,
      hasRssRef,
      hasAtomRef
    },
    reasons
  };
}

/**
 * 3. Otonom Trafik & Yüksek Ticari Niyet Matrisi (High Intent Revenue Routing)
 * @param {Array<Object>} products - Katalog ürün listesi
 */
export function generateHighIntentTrafficMatrix(products = []) {
  if (!products.length) {
    return { total: 0, highValuePillars: [], priorityKeywords: [] };
  }

  // Finansal değeri ve arama hacmi yüksek kritik sütunlar
  const HIGH_INTENT_CATEGORIES = [
    { slug: 'nakit-akisi', name: 'Nakit Akışı & Likidite', weight: 1.5, intent: 'P0 - Acil Yönetimsel İhtiyaç' },
    { slug: 'maliyet-karlilik', name: 'Birim Maliyet & Başabaş', weight: 1.4, intent: 'P0 - Marj Koruma & Fiyatlama' },
    { slug: 'butce-ve-planlama', name: 'Bütçe & Finansal Projeksiyon', weight: 1.3, intent: 'P1 - Kurumsal Planlama' },
    { slug: 'muhasebe-ve-vergi', name: 'Bordro & Vergi Matrahı', weight: 1.2, intent: 'P1 - Mevzuat & Uyumluluk' }
  ];

  const pillars = HIGH_INTENT_CATEGORIES.map((cat) => {
    const matchingProducts = products.filter((p) => {
      const c = (p.category || '').toLowerCase();
      return c.includes(cat.slug) || c.includes(cat.slug.replace(/-/g, ' '));
    });

    const averagePrice = matchingProducts.length
      ? Math.round(matchingProducts.reduce((sum, p) => sum + (p.priceTL || 0), 0) / matchingProducts.length)
      : 0;

    return {
      category: cat.slug,
      name: cat.name,
      intentRank: cat.intent,
      productCount: matchingProducts.length,
      averagePriceTL: averagePrice,
      commercialPowerIndex: Math.round(matchingProducts.length * cat.weight * 10) / 10
    };
  });

  return {
    totalEvaluated: products.length,
    pillars,
    trafficRecommendations: [
      'Nakit akışı ve başabaş şablonları IndexNow DAG kuyruğunda P0 öncelikle sürekli taze tutulmalı.',
      'AEO/LLM aracıları için /llms.txt ve /llms-full.txt üzerinden doğrudan derin varlık referansları sunulmalı.',
      'Tüm ürün sayfalarında tekil canonical ve anlık RSS 2.0 / Atom 1.0 bildirimleri devrede kalmalı.'
    ]
  };
}
