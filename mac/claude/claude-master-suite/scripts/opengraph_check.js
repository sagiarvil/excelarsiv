'use strict';
/**
 * opengraph_check.js — <head> SEO + Open Graph + Twitter Card denetimi.
 * Kullanım: node opengraph_check.js <URL|dosya> [--json]
 */
const L = require('./_lib');

async function main() {
  const args = L.parseArgs();
  const input = args._[0];
  if (!input) { console.error('Kullanım: node opengraph_check.js <URL|dosya> [--json]'); process.exit(2); }
  const { html, status, finalUrl } = await L.loadHtml(input);
  const rep = L.makeReport('OPEN GRAPH / META DENETİMİ — ' + input);
  if (status >= 400) rep.warn('HTTP', 'HTTP ' + status);

  const title = L.first(L.rx.title, html);
  if (!title) rep.fail('TITLE', '<title> yok');
  else {
    rep.pass('TITLE', `"${title}" (${title.length} kr)`);
    if (title.length < 15) rep.warn('TITLE', 'Title çok kısa (<15)');
    if (title.length > 65) rep.warn('TITLE', 'Title uzun (>65) — SERP kırpılır');
  }

  const desc = L.first(L.rx.metaName('description'), html);
  if (!desc) rep.fail('DESC', 'meta description yok');
  else {
    rep.pass('DESC', `${desc.length} kr`);
    if (desc.length < 70) rep.warn('DESC', 'Description kısa (<70)');
    if (desc.length > 165) rep.warn('DESC', 'Description uzun (>165) — kırpılır');
  }

  const canon = L.first(L.rx.linkRel('canonical'), html);
  if (!canon) rep.fail('CANON', 'rel=canonical yok');
  else {
    rep.pass('CANON', canon);
    if (!/^https?:\/\//i.test(canon)) rep.warn('CANON', 'Canonical mutlak URL değil');
  }

  const robots = L.first(L.rx.metaName('robots'), html);
  if (robots && /noindex/i.test(robots)) rep.warn('ROBOTS', 'Sayfa noindex: ' + robots);
  else rep.info('ROBOTS', robots || '(yok — indexlenir)');

  const viewport = L.first(L.rx.metaName('viewport'), html);
  if (!viewport) rep.fail('VIEWPORT', 'meta viewport yok (mobil uyumsuz)');

  // Open Graph
  const ogRequired = ['og:title', 'og:type', 'og:url', 'og:image', 'og:description'];
  for (const p of ogRequired) {
    const v = L.first(L.rx.metaProp(p), html);
    if (!v) rep.fail('OG', `${p} yok`);
    else rep.pass('OG', `${p} = ${v.slice(0, 80)}`);
  }
  const ogImg = L.first(L.rx.metaProp('og:image'), html);
  if (ogImg && !/^https?:\/\//i.test(ogImg)) rep.warn('OG', 'og:image mutlak URL değil');
  const siteName = L.first(L.rx.metaProp('og:site_name'), html);
  if (!siteName) rep.warn('OG', 'og:site_name yok');
  const locale = L.first(L.rx.metaProp('og:locale'), html);
  if (!locale) rep.warn('OG', 'og:locale yok (ör. tr_TR)');

  // Twitter
  const tCard = L.first(L.rx.metaName('twitter:card'), html);
  if (!tCard) rep.fail('TW', 'twitter:card yok');
  else rep.pass('TW', 'twitter:card = ' + tCard);
  for (const n of ['twitter:title', 'twitter:description', 'twitter:image']) {
    if (!L.first(L.rx.metaName(n), html)) rep.warn('TW', n + ' yok');
  }

  // og:type=article ise zaman damgaları
  const ogType = L.first(L.rx.metaProp('og:type'), html);
  if (ogType === 'article') {
    for (const p of ['article:published_time', 'article:modified_time']) {
      if (!L.first(L.rx.metaProp(p), html)) rep.warn('OG-ART', p + ' yok');
    }
  }

  rep.info('URL', finalUrl);
  if (args.json) { console.log(JSON.stringify({ summary: rep.summary(), text: rep.toText() }, null, 2)); process.exit(rep.summary().fail ? 1 : 0); }
  process.exit(rep.print());
}
main().catch((e) => { console.error('HATA:', e.message); process.exit(2); });
