'use strict';
/**
 * super-seo-motor ortak kütüphane — bağımlılıksız (yalnız Node çekirdeği).
 * Tüm araçlar buradan beslenir. SABİT YOL YOK: kök otomatik bulunur / argümandan alınır.
 */
const fs = require('fs');
const path = require('path');
const http = require('http');
const https = require('https');
const { URL } = require('url');

const ROOT_MARKERS = ['composer.json', 'package.json', '.git', 'artisan', path.join('public', 'index.php')];

/** Verilen dizinden yukarı çıkarak proje kökünü bul. Bulunamazsa startDir döner. */
function findProjectRoot(startDir) {
  let dir = path.resolve(startDir || process.cwd());
  while (true) {
    for (const m of ROOT_MARKERS) {
      if (fs.existsSync(path.join(dir, m))) return dir;
    }
    const parent = path.dirname(dir);
    if (parent === dir) return path.resolve(startDir || process.cwd());
    dir = parent;
  }
}

/** [PROJE_KÖKÜ]/docs/seotest yolu. */
function seotestDir(root) {
  return path.join(root || findProjectRoot(), 'docs', 'seotest');
}

function ensureDir(d) { fs.mkdirSync(d, { recursive: true }); }

/** Basit argv ayrıştırıcı: --key=val, --flag, konumsal. */
function parseArgs(argv) {
  const out = { _: [] };
  for (const a of (argv || process.argv.slice(2))) {
    if (a.startsWith('--')) {
      const [k, ...r] = a.slice(2).split('=');
      out[k] = r.length ? r.join('=') : true;
    } else out._.push(a);
  }
  return out;
}

/** URL getir (yönlendirme takipli). {status, headers, body, url} döner. */
function fetchUrl(target, opts = {}, redirects = 0) {
  return new Promise((resolve, reject) => {
    let u;
    try { u = new URL(target); } catch (e) { return reject(new Error('Geçersiz URL: ' + target)); }
    const mod = u.protocol === 'https:' ? https : http;
    const req = mod.request(u, {
      method: opts.method || 'GET',
      headers: Object.assign({
        'User-Agent': 'super-seo-motor/1.0 (+audit)',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
      }, opts.headers || {}),
      timeout: opts.timeout || 20000,
    }, (res) => {
      const loc = res.headers.location;
      if (res.statusCode >= 300 && res.statusCode < 400 && loc && redirects < 6) {
        res.resume();
        const next = new URL(loc, u).toString();
        return resolve(fetchUrl(next, opts, redirects + 1));
      }
      const chunks = [];
      res.on('data', (c) => chunks.push(c));
      res.on('end', () => resolve({
        status: res.statusCode,
        headers: res.headers,
        body: Buffer.concat(chunks).toString('utf8'),
        url: u.toString(),
      }));
    });
    req.on('timeout', () => req.destroy(new Error('timeout')));
    req.on('error', reject);
    if (opts.body) req.write(opts.body);
    req.end();
  });
}

/** Girdi URL de olabilir yerel dosya de. HTML metni döndürür. */
async function loadHtml(input) {
  if (/^https?:\/\//i.test(input)) {
    const r = await fetchUrl(input);
    return { html: r.body, status: r.status, finalUrl: r.url, headers: r.headers };
  }
  const p = path.resolve(input);
  return { html: fs.readFileSync(p, 'utf8'), status: 200, finalUrl: 'file://' + p, headers: {} };
}

/* ---------- HTML yardımcıları (regex tabanlı, DOM'suz) ---------- */
const rx = {
  title: /<title[^>]*>([\s\S]*?)<\/title>/i,
  metaName: (n) => new RegExp('<meta[^>]*name=["\']' + n + '["\'][^>]*content=["\']([^"\']*)["\']', 'i'),
  metaProp: (p) => new RegExp('<meta[^>]*property=["\']' + p + '["\'][^>]*content=["\']([^"\']*)["\']', 'i'),
  linkRel: (r) => new RegExp('<link[^>]*rel=["\']' + r + '["\'][^>]*href=["\']([^"\']*)["\']', 'i'),
  jsonLd: /<script[^>]*type=["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/gi,
  headings: /<(h[1-6])[^>]*>([\s\S]*?)<\/\1>/gi,
  imgTag: /<img\b[^>]*>/gi,
  scriptSrc: /<script\b[^>]*\bsrc=["']([^"']+)["'][^>]*>/gi,
  aHref: /<a\b[^>]*\bhref=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi,
};
function first(re, s) { const m = s.match(re); return m ? m[1].trim() : null; }
function stripTags(s) { return String(s || '').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim(); }
function attr(tag, name) {
  const m = tag.match(new RegExp('\\b' + name + '=["\']([^"\']*)["\']', 'i'));
  return m ? m[1] : null;
}

/** Tüm JSON-LD bloklarını çıkar ve parse et. [{ok,raw,data,error}] */
function extractJsonLd(html) {
  const out = [];
  let m;
  rx.jsonLd.lastIndex = 0;
  while ((m = rx.jsonLd.exec(html))) {
    const raw = m[1].trim();
    try { out.push({ ok: true, raw, data: JSON.parse(raw) }); }
    catch (e) { out.push({ ok: false, raw, error: e.message }); }
  }
  return out;
}

/** @graph'ı düzleştirip tüm node'ları tek diziye indir. */
function flattenNodes(parsedBlocks) {
  const nodes = [];
  for (const b of parsedBlocks) {
    if (!b.ok) continue;
    const d = b.data;
    if (Array.isArray(d)) d.forEach((n) => nodes.push(n));
    else if (d && Array.isArray(d['@graph'])) d['@graph'].forEach((n) => nodes.push(n));
    else if (d) nodes.push(d);
  }
  return nodes;
}

/* ---------- Raporlama ---------- */
function makeReport(title) {
  const rows = [];
  let pass = 0, fail = 0, warn = 0;
  const api = {
    pass: (id, msg) => { rows.push(['PASS', id, msg]); pass++; },
    fail: (id, msg) => { rows.push(['FAIL', id, msg]); fail++; },
    warn: (id, msg) => { rows.push(['WARN', id, msg]); warn++; },
    info: (id, msg) => { rows.push(['INFO', id, msg]); },
    summary: () => ({ pass, fail, warn, total: rows.length }),
    toText: () => {
      const L = [`# ${title}`, ''];
      for (const [s, id, msg] of rows) L.push(`- [${s}] ${id} — ${msg}`);
      L.push('', `SONUÇ: PASS=${pass}  FAIL=${fail}  WARN=${warn}`);
      return L.join('\n');
    },
    print: () => { console.log(api.toText()); return fail === 0 ? 0 : 1; },
  };
  return api;
}

module.exports = {
  fs, path, findProjectRoot, seotestDir, ensureDir, parseArgs,
  fetchUrl, loadHtml, rx, first, stripTags, attr,
  extractJsonLd, flattenNodes, makeReport,
};
