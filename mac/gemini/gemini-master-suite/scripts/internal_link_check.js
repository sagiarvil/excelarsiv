'use strict';
/**
 * internal_link_check.js — sığ tarama ile kırık iç link + zayıf çapa metni.
 * Kullanım: node internal_link_check.js <başlangıç-URL> [--max=60] [--depth=1] [--json]
 */
const L = require('./_lib');
const { URL } = require('url');

const GENERIC = /^(t[ıi]kla(y[ıi]n)?|buraya?|devam[ıi]?|read more|daha fazla|link|here|more|>>|»)$/i;

async function main() {
  const args = L.parseArgs();
  if (!args._[0]) { console.error('Kullanım: node internal_link_check.js <URL> [--max=60] [--depth=1]'); process.exit(2); }
  const start = new URL(args._[0]);
  const origin = start.origin;
  const max = parseInt(args.max || '60', 10);
  const maxDepth = parseInt(args.depth || '1', 10);
  const rep = L.makeReport('İÇ LİNK DENETİMİ — ' + origin);

  const seen = new Set();
  const queue = [[start.toString(), 0]];
  const statusCache = new Map();
  let checked = 0;

  async function head(u) {
    if (statusCache.has(u)) return statusCache.get(u);
    let st;
    try { st = (await L.fetchUrl(u)).status; } catch { st = 0; }
    statusCache.set(u, st);
    return st;
  }

  while (queue.length && checked < max) {
    const [url, depth] = queue.shift();
    if (seen.has(url)) continue;
    seen.add(url); checked++;
    let html;
    try { const r = await L.fetchUrl(url); html = r.body; if (r.status >= 400) { rep.fail('PAGE', `${r.status} ← ${url}`); continue; } }
    catch (e) { rep.fail('PAGE', 'erişilemedi: ' + url); continue; }

    let m; L.rx.aHref.lastIndex = 0;
    const links = [];
    while ((m = L.rx.aHref.exec(html))) links.push({ href: m[1], text: L.stripTags(m[2]) });

    for (const { href, text } of links) {
      if (/^(mailto:|tel:|javascript:|#)/i.test(href)) continue;
      let abs;
      try { abs = new URL(href, url); } catch { rep.warn('HREF', 'geçersiz href: ' + href + ' @ ' + url); continue; }
      if (abs.origin !== origin) continue;
      abs.hash = '';
      const clean = abs.toString();
      const st = await head(clean);
      if (st === 0) rep.fail('LINK', `erişilemedi: ${clean}  (kaynak: ${url})`);
      else if (st >= 400) rep.fail('LINK', `${st}: ${clean}  (kaynak: ${url})`);
      else if (st >= 300) rep.warn('LINK', `${st} redirect: ${clean}`);
      if (!text) rep.warn('ANCHOR', 'boş çapa metni → ' + clean);
      else if (GENERIC.test(text)) rep.warn('ANCHOR', `jenerik çapa "${text}" → ${clean}`);
      if (depth < maxDepth && !seen.has(clean) && st && st < 300) queue.push([clean, depth + 1]);
    }
  }
  rep.info('CRAWL', `${checked} sayfa tarandı, ${statusCache.size} benzersiz iç link kontrol edildi`);
  if (args.json) { console.log(JSON.stringify({ summary: rep.summary(), text: rep.toText() }, null, 2)); process.exit(rep.summary().fail ? 1 : 0); }
  process.exit(rep.print());
}
main().catch((e) => { console.error('HATA:', e.message); process.exit(2); });
