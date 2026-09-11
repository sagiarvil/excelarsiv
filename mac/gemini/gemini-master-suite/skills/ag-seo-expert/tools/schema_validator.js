'use strict';
/**
 * schema_validator.js — Google yapılandırılmış veri / Rich Results doğrulayıcı.
 *
 * Kullanım:
 *   node schema_validator.js <URL | yerel.html> [--json]
 *
 * Denetler:
 *  - Tüm JSON-LD blokları geçerli JSON mu, @context schema.org mu
 *  - @graph yapısı, @id benzersiz mi, dahili @id referansları çözülüyor mu
 *  - Tip bazında Google zorunlu/önerilen alanlar (Organization, WebSite, WebPage,
 *    BreadcrumbList, Product, Offer/AggregateOffer, SoftwareApplication, Service,
 *    Article/BlogPosting, FAQPage, LocalBusiness, ImageObject)
 *  - Google politika riski: uydurma aggregateRating/review, boş/placeholder sameAs,
 *    fiyatsız Offer, geçmiş priceValidUntil
 */
const L = require('./_lib');

const REQUIRED = {
  Organization: { must: ['name', 'url'], should: ['logo', 'sameAs'] },
  WebSite: { must: ['name', 'url'], should: ['potentialAction'] },
  WebPage: { must: ['name'], should: ['url', 'isPartOf'] },
  BreadcrumbList: { must: ['itemListElement'], should: [] },
  ListItem: { must: ['position', 'name'], should: ['item'] },
  Product: { must: ['name'], should: ['image', 'description', 'offers', 'brand'] },
  Offer: { must: ['price', 'priceCurrency'], should: ['availability', 'url', 'priceValidUntil'] },
  AggregateOffer: { must: ['lowPrice', 'priceCurrency'], should: ['highPrice', 'offerCount'] },
  SoftwareApplication: { must: ['name', 'applicationCategory'], should: ['offers', 'operatingSystem', 'aggregateRating'] },
  Service: { must: ['name'], should: ['provider', 'areaServed', 'serviceType'] },
  Article: { must: ['headline'], should: ['image', 'datePublished', 'author', 'publisher'] },
  BlogPosting: { must: ['headline'], should: ['image', 'datePublished', 'author', 'publisher'] },
  NewsArticle: { must: ['headline'], should: ['image', 'datePublished', 'author', 'publisher'] },
  FAQPage: { must: ['mainEntity'], should: [] },
  Question: { must: ['name', 'acceptedAnswer'], should: [] },
  Answer: { must: ['text'], should: [] },
  LocalBusiness: { must: ['name', 'address'], should: ['telephone', 'openingHoursSpecification', 'geo'] },
  ImageObject: { must: ['url'], should: ['width', 'height'] },
  AggregateRating: { must: ['ratingValue', 'ratingCount'], should: ['bestRating'] },
  Review: { must: ['reviewRating', 'author'], should: [] },
};
const PLACEHOLDER_RX = /(example\.com|yourdomain|placeholder|xxx+|lorem|test\.test|\{\{)/i;

function typeList(node) {
  const t = node && node['@type'];
  return (Array.isArray(t) ? t : [t]).filter(Boolean);
}
function has(node, key) {
  return node && node[key] !== undefined && node[key] !== null && node[key] !== '' &&
    !(Array.isArray(node[key]) && node[key].length === 0);
}

async function main() {
  const args = L.parseArgs();
  const input = args._[0];
  if (!input) { console.error('Kullanım: node schema_validator.js <URL|dosya> [--json]'); process.exit(2); }

  const { html, status } = await L.loadHtml(input);
  const rep = L.makeReport('SCHEMA / RICH RESULTS DOĞRULAMA — ' + input);
  if (status && status >= 400) rep.warn('HTTP', 'Sayfa HTTP ' + status + ' döndü');

  const blocks = L.extractJsonLd(html);
  if (blocks.length === 0) { rep.fail('LD-00', 'Sayfada hiç JSON-LD bloğu yok'); finish(rep, args); return; }
  rep.pass('LD-00', blocks.length + ' JSON-LD bloğu bulundu');

  blocks.forEach((b, i) => {
    if (!b.ok) { rep.fail('LD-PARSE', `Blok #${i + 1} JSON hatası: ${b.error}`); return; }
    const ctx = b.data['@context'] || (b.data['@graph'] && '(graph)') || null;
    if (!/schema\.org/i.test(JSON.stringify(b.data['@context'] || ''))) {
      rep.warn('LD-CTX', `Blok #${i + 1} @context schema.org değil: ${JSON.stringify(ctx)}`);
    }
  });

  const nodes = L.flattenNodes(blocks);
  rep.info('LD-CNT', nodes.length + ' node (graph düzleştirildi)');

  // @id benzersizlik + referans çözümü
  const ids = new Map();
  for (const n of nodes) if (n && n['@id']) {
    if (ids.has(n['@id'])) rep.warn('LD-ID', 'Tekrar eden @id: ' + n['@id']);
    ids.set(n['@id'], n);
  }
  const refRx = /"@id"\s*:\s*"([^"]+)"/g;
  let rm; const seenRef = new Set();
  const flat = JSON.stringify(nodes);
  while ((rm = refRx.exec(flat))) seenRef.add(rm[1]);
  for (const ref of seenRef) {
    if (!ids.has(ref) && /#/.test(ref)) rep.warn('LD-REF', 'Çözülemeyen dahili @id referansı: ' + ref);
  }

  // Tip bazında alan denetimi
  const foundTypes = new Set();
  for (const n of nodes) {
    for (const t of typeList(n)) {
      foundTypes.add(t);
      const spec = REQUIRED[t];
      if (!spec) continue;
      for (const k of spec.must) if (!has(n, k)) rep.fail('REQ', `${t}: zorunlu alan eksik → ${k}`);
      for (const k of spec.should) if (!has(n, k)) rep.warn('REC', `${t}: önerilen alan eksik → ${k}`);
    }

    // Politika riskleri
    if (has(n, 'sameAs')) {
      const arr = Array.isArray(n.sameAs) ? n.sameAs : [n.sameAs];
      for (const s of arr) if (typeof s !== 'string' || !/^https?:\/\/.+\..+/.test(s) || PLACEHOLDER_RX.test(s))
        rep.warn('POL-SAMEAS', 'Geçersiz/placeholder sameAs: ' + JSON.stringify(s));
    }
    if (typeList(n).includes('Offer')) {
      const p = parseFloat(n.price);
      if (!(p > 0)) rep.fail('POL-OFFER', 'Offer.price ≤ 0 / geçersiz — Google geçersiz Offer sayar');
      if (n.priceValidUntil && Date.parse(n.priceValidUntil) < Date.now())
        rep.warn('POL-PVU', 'priceValidUntil geçmişte: ' + n.priceValidUntil);
    }
    if (typeList(n).includes('AggregateRating')) {
      const rc = parseFloat(n.ratingCount || n.reviewCount);
      if (!(rc > 0)) rep.fail('POL-RATING', 'aggregateRating var ama ratingCount/reviewCount yok — uydurma veri riski');
    }
  }

  // Taban blok beklentisi
  if (!foundTypes.has('Organization')) rep.warn('BASE', 'Organization node yok (taban blok beklenir)');
  if (!foundTypes.has('WebSite')) rep.warn('BASE', 'WebSite node yok (taban blok beklenir)');
  rep.info('TYPES', 'Bulunan tipler: ' + [...foundTypes].sort().join(', '));

  finish(rep, args);
}

function finish(rep, args) {
  if (args.json) {
    console.log(JSON.stringify({ summary: rep.summary(), text: rep.toText() }, null, 2));
    process.exit(rep.summary().fail ? 1 : 0);
  }
  process.exit(rep.print());
}

main().catch((e) => { console.error('HATA:', e.message); process.exit(2); });
