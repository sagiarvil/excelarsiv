'use strict';
/**
 * sitemap_robots_check.js — robots.txt, sitemap.xml, llms.txt, agent-card.json denetimi.
 * Kullanım: node sitemap_robots_check.js <origin-URL> [--max=40] [--json]
 *   <origin-URL>: https://alan.tld  (yol verilirse origin'e indirilir)
 */
const L = require('./_lib');
const { URL } = require('url');

async function main() {
  const args = L.parseArgs();
  if (!args._[0]) { console.error('Kullanım: node sitemap_robots_check.js <https://alan.tld> [--max=40]'); process.exit(2); }
  const origin = new URL(args._[0]).origin;
  const max = parseInt(args.max || '40', 10);
  const rep = L.makeReport('ROBOTS / SITEMAP / LLM YÜZEYLERİ — ' + origin);

  // robots.txt
  let sitemaps = [];
  try {
    const r = await L.fetchUrl(origin + '/robots.txt');
    if (r.status === 200 && r.body.trim()) {
      rep.pass('ROBOTS', 'robots.txt bulundu (' + r.body.length + ' bayt)');
      if (/Disallow:\s*\/\s*$/im.test(r.body)) rep.fail('ROBOTS', 'Tüm site Disallow: / ile kapalı');
      sitemaps = [...r.body.matchAll(/^\s*Sitemap:\s*(\S+)/gim)].map((m) => m[1]);
      if (!sitemaps.length) rep.warn('ROBOTS', 'robots.txt içinde Sitemap: satırı yok');
      else rep.pass('ROBOTS', sitemaps.length + ' sitemap bildirimi');
    } else rep.fail('ROBOTS', 'robots.txt yok / HTTP ' + r.status);
  } catch (e) { rep.fail('ROBOTS', 'robots.txt alınamadı: ' + e.message); }

  if (!sitemaps.length) sitemaps = [origin + '/sitemap.xml'];

  // sitemap(ler)
  const urls = [];
  for (const sm of sitemaps.slice(0, 5)) {
    try {
      const r = await L.fetchUrl(sm);
      if (r.status !== 200) { rep.fail('SITEMAP', sm + ' → HTTP ' + r.status); continue; }
      if (!/<(urlset|sitemapindex)\b/i.test(r.body)) { rep.fail('SITEMAP', sm + ' geçerli XML değil'); continue; }
      const isIndex = /<sitemapindex\b/i.test(r.body);
      const locs = [...r.body.matchAll(/<loc>\s*([^<\s]+)\s*<\/loc>/gi)].map((m) => m[1]);
      rep.pass('SITEMAP', `${sm} → ${locs.length} ${isIndex ? 'alt-sitemap' : 'URL'}`);
      if (isIndex) {
        for (const sub of locs.slice(0, 5)) {
          try {
            const rs = await L.fetchUrl(sub);
            [...rs.body.matchAll(/<loc>\s*([^<\s]+)\s*<\/loc>/gi)].forEach((m) => urls.push(m[1]));
          } catch (e) { rep.warn('SITEMAP', 'alt sitemap alınamadı: ' + sub); }
        }
      } else locs.forEach((u) => urls.push(u));
      if (/<lastmod>\s*<\/lastmod>/i.test(r.body)) rep.warn('SITEMAP', 'boş <lastmod> var');
    } catch (e) { rep.fail('SITEMAP', sm + ' alınamadı: ' + e.message); }
  }

  // URL örnekleme — HTTP durum kontrolü
  const sample = [...new Set(urls)].slice(0, max);
  let bad = 0;
  for (const u of sample) {
    try {
      const r = await L.fetchUrl(u, { method: 'GET' });
      if (r.status >= 400) { rep.fail('URL', `${r.status} ← ${u}`); bad++; }
      else if (r.url.replace(/\/$/, '') !== u.replace(/\/$/, '')) rep.warn('URL', `redirect: ${u} → ${r.url}`);
    } catch (e) { rep.fail('URL', 'erişilemedi: ' + u); bad++; }
  }
  if (sample.length) rep[bad ? 'warn' : 'pass']('URL', `${sample.length} URL örneklendi, ${bad} hatalı`);

  // LLM / makine yüzeyleri
  for (const p of ['/llms.txt', '/llms-full.txt', '/.well-known/agent-card.json', '/.well-known/ai-plugin.json']) {
    try {
      const r = await L.fetchUrl(origin + p);
      if (r.status === 200 && r.body.trim()) rep.pass('LLM', p + ' mevcut');
      else rep.warn('LLM', p + ' yok (HTTP ' + r.status + ')');
    } catch (e) { rep.warn('LLM', p + ' alınamadı'); }
  }

  if (args.json) { console.log(JSON.stringify({ summary: rep.summary(), text: rep.toText() }, null, 2)); process.exit(rep.summary().fail ? 1 : 0); }
  process.exit(rep.print());
}
main().catch((e) => { console.error('HATA:', e.message); process.exit(2); });
