'use strict';
/**
 * web_vitals_hints.js — Core Web Vitals statik ipuçları (lab değil, kaynak taraması).
 * Kullanım: node web_vitals_hints.js <URL|dosya> [--json]
 */
const L = require('./_lib');

async function main() {
  const args = L.parseArgs();
  if (!args._[0]) { console.error('Kullanım: node web_vitals_hints.js <URL|dosya>'); process.exit(2); }
  const { html, headers } = await L.loadHtml(args._[0]);
  const rep = L.makeReport('CORE WEB VITALS — STATİK İPUÇLARI — ' + args._[0]);

  const bytes = Buffer.byteLength(html, 'utf8');
  rep.info('HTML', bytes + ' bayt');
  if (bytes > 102400) rep.warn('HTML', 'HTML > 100KB — TTFB/LCP riski');
  rep.info('AST-14KB', bytes <= 14336 ? 'İlk 14KB bütçesi içinde' : 'İlk 14KB bütçesi aşıldı (kritik içerik geç)');

  // Render-blocking
  const headMatch = html.match(/<head[^>]*>([\s\S]*?)<\/head>/i);
  const head = headMatch ? headMatch[1] : html;
  const blkCss = [...head.matchAll(/<link[^>]*rel=["']stylesheet["'][^>]*>/gi)]
    .filter((m) => !/\bmedia=["'](print|[^"']*\bprint)["']/i.test(m[0]) && !/\bonload=/i.test(m[0]));
  if (blkCss.length > 2) rep.warn('CSS', blkCss.length + ' render-blocking stylesheet (<=2 hedefle)');
  else rep.pass('CSS', blkCss.length + ' render-blocking stylesheet');

  const headScripts = [...head.matchAll(/<script\b[^>]*\bsrc=["'][^"']+["'][^>]*>/gi)]
    .filter((m) => !/\b(async|defer|type=["']module["'])\b/i.test(m[0]));
  if (headScripts.length) rep.fail('JS', headScripts.length + ' <head> içinde async/defer olmayan script');
  else rep.pass('JS', '<head> scriptleri async/defer/module');

  // 3. parti
  const ext = [...html.matchAll(L.rx.scriptSrc)].map((m) => m[1]).filter((s) => /^https?:\/\//i.test(s));
  const thirdParty = ext.filter((s) => !/(^\/|localhost)/.test(s));
  if (thirdParty.length) rep.warn('3P', thirdParty.length + ' harici script: ' + [...new Set(thirdParty.map((u) => { try { return new URL(u).host; } catch { return u; } }))].join(', '));

  // Görseller — boyut / lazy / modern format
  const imgs = html.match(L.rx.imgTag) || [];
  let noDim = 0, noLazy = 0, legacy = 0;
  imgs.forEach((t) => {
    if (L.attr(t, 'width') === null || L.attr(t, 'height') === null) noDim++;
    if (!/\bloading=["']lazy["']/i.test(t) && !/\bfetchpriority=["']high["']/i.test(t)) noLazy++;
    const src = L.attr(t, 'src') || '';
    if (/\.(jpe?g|png|gif)(\?|$)/i.test(src)) legacy++;
  });
  if (imgs.length) {
    rep[noDim ? 'warn' : 'pass']('IMG-CLS', `${noDim}/${imgs.length} görselde width/height yok (CLS riski)`);
    rep[noLazy > 1 ? 'warn' : 'pass']('IMG-LAZY', `${noLazy}/${imgs.length} görselde loading=lazy yok`);
    if (legacy) rep.warn('IMG-FMT', `${legacy} görsel jpg/png/gif — webp/avif düşün`);
  }

  // preconnect / dns-prefetch
  if (thirdParty.length && !/rel=["'](preconnect|dns-prefetch)["']/i.test(head))
    rep.warn('PRECONN', 'Harici köken var ama preconnect/dns-prefetch yok');

  // Font
  if (/<link[^>]*fonts?\.(googleapis|gstatic)/i.test(head) && !/rel=["']preconnect["'][^>]*gstatic/i.test(head))
    rep.warn('FONT', 'Google Fonts var, gstatic preconnect yok');
  if (/@font-face/i.test(html) && !/font-display\s*:/i.test(html))
    rep.warn('FONT', '@font-face var, font-display yok (FOIT)');

  // Sıkıştırma / cache başlıkları (URL ise)
  if (headers && headers['content-encoding']) rep.pass('GZIP', 'content-encoding: ' + headers['content-encoding']);
  else if (headers && Object.keys(headers).length) rep.warn('GZIP', 'content-encoding başlığı yok (br/gzip?)');
  if (headers && !headers['cache-control']) rep.warn('CACHE', 'cache-control başlığı yok');

  if (args.json) { console.log(JSON.stringify({ summary: rep.summary(), text: rep.toText() }, null, 2)); process.exit(rep.summary().fail ? 1 : 0); }
  process.exit(rep.print());
}
main().catch((e) => { console.error('HATA:', e.message); process.exit(2); });
