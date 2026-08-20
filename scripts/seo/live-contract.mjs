// Canlı SEO sözleşmesi (madde 17): production sonrası tüm sitemap URL'lerini doğrular.
// robots.txt -> sitemap.xml -> her child 200 -> her URL 200 + self-canonical
// + noindex yok + H1 + JSON-LD. Index lastmod ISO + future değil.
// + Answer Graph kritik crawler/retired-route sözleşmeleri.
import { resolve } from 'node:path';
import { isImageSitemapChild } from './validate-gates.mjs';

const BASE = process.env.SEO_BASE_URL ?? 'https://excelarsiv.com';
const RETRY = 6;
const RETRY_DELAY_MS = 2000;
const FUTURE_GRACE_MS = 5 * 60 * 1000;
const TRANSIENT_STATUS = new Set([408, 425, 429, 500, 502, 503, 504]);
const RETIRED_ROUTES = new Map([
  ['/excel-araclari', '/sablonlar'],
  ['/paketler', '/sablon/kobi-finans-yonetim-paketi'],
]);
const COMPARISON_ROUTE = '/neden-excel-arsiv';

const failures = [];
let checks = 0;

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

function errorLabel(error) {
  if (!error) return 'fetch_failed';
  if (error.name === 'AbortError' || error.code === 'ABORT_ERR') return 'abort';
  return String(error.cause?.code || error.code || error.message || 'fetch_failed');
}

async function get(url, { redirect = 'follow' } = {}) {
  let lastError = '';
  for (let attempt = 1; attempt <= RETRY; attempt++) {
    const timeoutMs = attempt === RETRY ? 45_000 : 20_000;
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    try {
      const res = await fetch(url, { signal: controller.signal, redirect });
      const payload = {
        status: res.status,
        location: res.headers.get('location') ?? '',
        contentType: res.headers.get('content-type') ?? '',
        text: await res.text(),
        error: '',
      };
      if (TRANSIENT_STATUS.has(payload.status) && attempt < RETRY) {
        await sleep(RETRY_DELAY_MS * attempt);
        continue;
      }
      return payload;
    } catch (error) {
      lastError = errorLabel(error);
      if (attempt === RETRY) return { status: 0, location: '', contentType: '', text: '', error: lastError };
      await sleep(RETRY_DELAY_MS * attempt);
    } finally {
      clearTimeout(timer);
    }
  }
  return { status: 0, location: '', contentType: '', text: '', error: lastError };
}

function check(ok, message) {
  checks++;
  if (!ok) failures.push(message);
}

function parseIndex(xml) {
  return [...xml.matchAll(/<sitemap>([\s\S]*?)<\/sitemap>/g)]
    .map((block) => ({
      loc: block[1].match(/<loc>([^<]+)<\/loc>/)?.[1]?.trim(),
      lastmod: block[1].match(/<lastmod>([^<]+)<\/lastmod>/)?.[1]?.trim(),
    }))
    .filter((entry) => entry.loc);
}

function parseUrls(xml) {
  return [...xml.matchAll(/<url>([\s\S]*?)<\/url>/g)]
    .map((block) => block[1].match(/<loc>([^<]+)<\/loc>/)?.[1]?.trim())
    .filter(Boolean);
}

const robots = await get(`${BASE}/robots.txt`);
check(robots.status === 200, `robots.txt HTTP ${robots.status}`);
check(robots.text.includes(`Sitemap: ${BASE}/sitemap.xml`), 'robots.txt sitemap.xml bildirmiyor');
check(/User-agent:\s*OAI-SearchBot[\s\S]*?Allow:\s*\//i.test(robots.text), 'robots.txt OAI-SearchBot için Allow: / içermiyor');

for (const [route, destination] of RETIRED_ROUTES) {
  const retired = await get(`${BASE}${route}`, { redirect: 'manual' });
  check(retired.status === 301, `retired route 301 dönmüyor: ${route} (HTTP ${retired.status})`);
  const expectedLocation = new URL(destination, BASE).href;
  const actualLocation = retired.location ? new URL(retired.location, BASE).href : '';
  check(actualLocation === expectedLocation, `retired route yanlış hedefe yönleniyor: ${route} -> ${retired.location || 'Location yok'}`);
}

const indexRes = await get(`${BASE}/sitemap.xml`);
check(indexRes.status === 200, `sitemap.xml HTTP ${indexRes.status}`);
check(indexRes.text.includes('<sitemapindex'), 'sitemap.xml sitemapindex köküne sahip değil');
const indexContentType = (indexRes.contentType ?? '').split(';')[0].trim().toLowerCase();
check(
  ['application/xml', 'text/xml'].includes(indexContentType),
  `sitemap.xml Content-Type uygun değil: ${indexRes.contentType || 'yok'}`,
);

const indexEntries = indexRes.status === 200 ? parseIndex(indexRes.text) : [];
check(indexEntries.length > 0, 'sitemap index boş');
const seenChildren = new Set();
for (const entry of indexEntries) {
  if (seenChildren.has(entry.loc)) check(false, `index duplicate child: ${entry.loc}`);
  seenChildren.add(entry.loc);
  if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(entry.lastmod ?? '')) check(false, `index child lastmod ISO değil: ${entry.loc}`);
  if (Date.parse(entry.lastmod ?? '') > Date.now() + FUTURE_GRACE_MS) check(false, `index child future lastmod: ${entry.loc}`);
}

const allUrls = new Set();
let imageSitemapCount = 0;
for (const entry of indexEntries) {
  const child = await get(entry.loc);
  check(child.status === 200, `child HTTP ${child.status}: ${entry.loc}`);
  if (child.status !== 200) continue;
  check(child.text.includes('<urlset'), `child <urlset> eksik: ${entry.loc}`);
  check(!child.text.includes('<changefreq>'), `child changefreq içeriyor: ${entry.loc}`);
  check(!child.text.includes('<priority>'), `child priority içeriyor: ${entry.loc}`);
  const urls = parseUrls(child.text);
  check(urls.length > 0, `child 0 URL: ${entry.loc}`);

  if (isImageSitemapChild(entry.loc, child.text)) {
    imageSitemapCount += 1;
    check(child.text.includes('xmlns:image='), `image sitemap namespace eksik: ${entry.loc}`);
    check(child.text.includes('<image:loc>'), `image sitemap image:loc yok: ${entry.loc}`);
    continue;
  }

  for (const loc of urls) {
    if (allUrls.has(loc)) check(false, `duplicate URL: ${loc}`);
    allUrls.add(loc);
  }
}
check(imageSitemapCount > 0, 'sitemap index içinde image sitemap child yok');
check(allUrls.has(`${BASE}/`), 'homepage sitemap içinde yok');
check(allUrls.has(`${BASE}${COMPARISON_ROUTE}`), 'karşılaştırma otorite sayfası sitemap içinde yok');
for (const route of RETIRED_ROUTES.keys()) check(!allUrls.has(`${BASE}${route}`), `retired route sitemap içinde: ${route}`);

const homepage = await get(`${BASE}/`);
check(homepage.status === 200, `homepage HTTP ${homepage.status}`);

const comparison = await get(`${BASE}${COMPARISON_ROUTE}`);
check(comparison.status === 200, `karşılaştırma otorite sayfası HTTP ${comparison.status}`);
check(comparison.text.includes('ExcelArşiv') && comparison.text.includes('ChatGPT') && comparison.text.includes('Boş Excel'), 'karşılaştırma otorite sayfası üç yaklaşımı birlikte açıklamıyor');

for (const url of allUrls) {
  const res = await get(url);
  check(res.status === 200, `URL HTTP ${res.status}: ${url}${res.error ? ` (${res.error})` : ''}`);
  if (res.status !== 200) continue;
  const canonical = res.text.match(/<link[^>]+rel=["']canonical["'][^>]*href=["']([^"']+)["']/i)?.[1];
  check(canonical === url, `self-canonical değil: ${url} (canonical=${canonical})`);
  const robotsMeta = res.text.match(/<meta[^>]+name=["']robots["'][^>]*content=["']([^"']+)["']/i)?.[1] ?? '';
  check(!robotsMeta.toLowerCase().includes('noindex'), `noindex sayfa sitemap içinde: ${url}`);
  check(/<h1[\s>]/i.test(res.text), `H1 eksik: ${url}`);
  check(res.text.includes('application/ld+json'), `JSON-LD eksik: ${url}`);
}

console.log(`SEO LIVE CONTRACT: ${checks} kontrol, ${failures.length} hata`);
if (failures.length > 0) {
  for (const f of failures) console.error(`  - ${f}`);
  process.exit(1);
}
console.log('SEO LIVE CONTRACT PASS');
