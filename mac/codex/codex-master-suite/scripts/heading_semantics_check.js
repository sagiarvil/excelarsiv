'use strict';
/**
 * heading_semantics_check.js — başlık hiyerarşisi + anlamsal/landmark yapı.
 * Kullanım: node heading_semantics_check.js <URL|dosya> [--json]
 */
const L = require('./_lib');

async function main() {
  const args = L.parseArgs();
  if (!args._[0]) { console.error('Kullanım: node heading_semantics_check.js <URL|dosya>'); process.exit(2); }
  const { html } = await L.loadHtml(args._[0]);
  const rep = L.makeReport('BAŞLIK & ANLAMSAL YAPI — ' + args._[0]);

  const hs = [];
  let m; L.rx.headings.lastIndex = 0;
  while ((m = L.rx.headings.exec(html))) hs.push({ level: +m[1][1], text: L.stripTags(m[2]) });

  const h1 = hs.filter((h) => h.level === 1);
  if (h1.length === 0) rep.fail('H1', 'H1 yok');
  else if (h1.length > 1) rep.fail('H1', h1.length + ' adet H1 (tek olmalı): ' + h1.map((h) => `"${h.text}"`).join(', '));
  else rep.pass('H1', `Tek H1: "${h1[0].text}"`);

  let prev = 0, skips = 0;
  for (const h of hs) {
    if (prev && h.level > prev + 1) { rep.fail('SKIP', `Seviye atlama: H${prev} → H${h.level} ("${h.text}")`); skips++; }
    prev = h.level;
  }
  if (!skips && hs.length) rep.pass('SKIP', 'Başlık seviyesi atlaması yok');
  for (const h of hs) if (!h.text) rep.warn('EMPTY', `Boş H${h.level} etiketi`);
  rep.info('OUTLINE', hs.map((h) => '  '.repeat(h.level - 1) + 'H' + h.level + ' ' + h.text).join('\n') || '(başlık yok)');

  // Landmark / anlamsal etiketler
  const need = { '<header': 'header', '<nav': 'nav', '<main': 'main', '<footer': 'footer' };
  for (const [tag, name] of Object.entries(need)) {
    const c = (html.match(new RegExp(tag + '\\b', 'gi')) || []).length;
    if (name === 'main' && c !== 1) rep.fail('LANDMARK', `<main> ${c} adet (tam 1 olmalı)`);
    else if (c === 0) rep.warn('LANDMARK', `<${name}> yok`);
    else rep.pass('LANDMARK', `<${name}> ×${c}`);
  }
  if (!/<html[^>]*\blang=/i.test(html)) rep.fail('LANG', '<html lang> yok');
  else rep.pass('LANG', L.first(/<html[^>]*\blang=["']([^"']+)["']/i, html));

  // Görsel alt metni
  const imgs = html.match(L.rx.imgTag) || [];
  const noAlt = imgs.filter((t) => L.attr(t, 'alt') === null);
  if (imgs.length) rep[noAlt.length ? 'warn' : 'pass']('IMG-ALT', `${imgs.length} <img>, ${noAlt.length} alt eksik`);

  if (args.json) { console.log(JSON.stringify({ summary: rep.summary(), text: rep.toText() }, null, 2)); process.exit(rep.summary().fail ? 1 : 0); }
  process.exit(rep.print());
}
main().catch((e) => { console.error('HATA:', e.message); process.exit(2); });
