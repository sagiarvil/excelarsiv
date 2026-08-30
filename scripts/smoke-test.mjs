// Smoke test — dist/ HTML'lerini render + kırık iç link + içerik için doğrular.
// Bağımlılıksız (node >= 18). Kullanım: npm test
import { readFileSync, readdirSync, statSync } from 'node:fs';
import { join, extname, resolve } from 'node:path';

const dist = resolve(process.cwd(), 'dist');
const SPECIAL_LEAD_PAGE = 'ozel-excel-sistemleri/index.html';
const HOME_PAGE = 'index.html';
const VERIFIED_WHATSAPP_PREFIX = 'https://wa.me/905393333303?text=';

function walk(dir, out = []) {
  for (const entry of readdirSync(dir)) {
    const full = join(dir, entry);
    if (statSync(full).isDirectory()) {
      walk(full, out);
    } else if (extname(full) === '.html') {
      out.push(full);
    }
  }
  return out;
}

const pages = walk(dist).map((file) => file.replace(dist + '/', ''));
const failures = [];

function existsAsFile(base) {
  try {
    return statSync(base).isFile();
  } catch {
    return false;
  }
}

for (const page of pages) {
  const html = readFileSync(join(dist, page), 'utf8');

  if (!/<title>[\s\S]*<\/title>/.test(html)) {
    failures.push(`${page}: <title> eksik`);
  }
  if (!/<main[\s>]/.test(html)) {
    failures.push(`${page}: <main> eksik`);
  }

  // WhatsApp doğrudan sipariş kanalı olarak yasak kalır.
  // İzin verilen tek yüzeyler: ana sayfadaki danışma CTA'sı ve özel Excel sistemleri lead/ön görüşme CTA'ları.
  const whatsappLinks = [...html.matchAll(/href="(https:\/\/wa\.me\/[^"]+)"/g)].map((m) => m[1]);
  if (whatsappLinks.length > 0) {
    if (page !== SPECIAL_LEAD_PAGE && page !== HOME_PAGE) {
      failures.push(`${page}: WhatsApp sipariş linki yasak`);
    } else {
      for (const href of whatsappLinks) {
        if (!href.startsWith(VERIFIED_WHATSAPP_PREFIX)) {
          failures.push(`${page}: doğrulanmamış WhatsApp lead linki -> ${href}`);
        }
      }
      if (page === SPECIAL_LEAD_PAGE && !html.includes('data-special-premium-v17')) {
        failures.push(`${page}: WhatsApp lead istisnası için special premium V17 contract eksik`);
      }
      if (page === HOME_PAGE && !html.includes('data-dual-funnel-home-v17')) {
        failures.push(`${page}: WhatsApp lead istisnası için home dual-funnel V17 contract eksik`);
      }
    }
  }

  // Kırık iç link kontrolü: yalnızca kök-relative href'ler (dış linkler hariç).
  // Query string (?q=...) ve fragment (#...) yolun parçası değildir.
  const links = [...html.matchAll(/href="\/([^"#?]*?)(?:[?#][\s\S]*?)?"/g)]
    .map((m) => m[1].replace(/\/+$/, ''))
    .filter((p) => p.length > 0);
  for (const target of [...new Set(links)]) {
    if (target.includes('.')) {
      if (existsAsFile(join(dist, target))) continue;
      failures.push(`${page}: kırık link -> /${target}`);
      continue;
    }
    const exists =
      existsAsFile(join(dist, `${target}.html`)) || existsAsFile(join(dist, target, 'index.html'));
    if (!exists) {
      failures.push(`${page}: kırık link -> /${target}`);
    }
  }

  const scripts = [...html.matchAll(/<script(?!\s+type="application\/ld\+json")[^>]*>([\s\S]*?)<\/script>/g)]
    .map((m) => ({ src: /src=/.test(m[0]), body: m[1] }))
    .filter((s) => !s.src && s.body.trim().length > 0);
  for (const s of scripts) {
    try {
      new Function(s.body);
    } catch (e) {
      failures.push(`${page}: geçersiz inline JS -> ${e.message}`);
    }
  }

  if (html.includes('data-template-grid')) {
    if (!html.includes('data-template-grid-wrap')) {
      failures.push(`${page}: filtre boş durumu için data-template-grid-wrap yok`);
    }
    const storyCards = html.match(/<a\b[^>]*\bdata-story-card\b/g)?.length ?? 0;
    const articleCards = html.match(/<article\b[^>]*\bdata-template-card\b/g)?.length ?? 0;
    const cards = storyCards + articleCards;
    if (cards < 1) failures.push(`${page}: katalog kartı yok`);
    if (storyCards > 0 && articleCards > 0) {
      failures.push(`${page}: iki kart varyantı aynı anda render edilmiş`);
    }
    const fakeVisual = html.match(/product-visual|pv-kpi|●\s*Canlı/g);
    if (fakeVisual) failures.push(`${page}: sahte KPI/Canlı rozeti -> ${fakeVisual[0]}`);
    const canli = html.match(/Canlı/g)?.length ?? 0;
    if (canli > 0) failures.push(`${page}: Gate 2 Canlı=${canli}`);

    if (storyCards > 0) {
      const storyLinks = [...html.matchAll(/<a\b[^>]*\bdata-story-card\b[^>]*href="([^"]+)"/g)];
      for (const match of storyLinks) {
        if (!match[1].startsWith('/sablon/')) {
          failures.push(`${page}: story kart detay slug'a gitmiyor -> ${match[1]}`);
        }
      }
      const storySlugs = html.match(/data-story-slug="[^"]*"/g)?.length ?? 0;
      if (storySlugs !== storyCards) {
        failures.push(`${page}: data-story-slug ${storySlugs}/${storyCards}`);
      }
    } else {
      const orders = [...html.matchAll(/class="card__cta card__cta--primary"[^>]*href="([^"]+)"/g)];
      if (orders.length !== articleCards) {
        failures.push(`${page}: Satın Al CTA sayısı ${orders.length}, kart ${articleCards}`);
      }
      for (const match of orders) {
        const href = match[1];
        if (!/^https:\/\/www\.shopier\.com\/\d+$/.test(href)) {
          failures.push(`${page}: Satın Al Shopier ürün linkine gitmiyor -> ${href}`);
        }
      }
      const details = [...html.matchAll(/class="card__cta card__cta--ghost"[^>]*href="([^"]+)"/g)];
      for (const match of details) {
        if (!match[1].startsWith('/sablon/')) {
          failures.push(`${page}: Detay slug'a gitmiyor -> ${match[1]}`);
        }
      }
      const focuses = html.match(/data-focus="result"/g)?.length ?? 0;
      if (focuses !== articleCards) {
        failures.push(`${page}: screenshotFocus=result ${focuses}/${articleCards}`);
      }
    }
  }

  if (page === HOME_PAGE) {
    if (/product-visual|pv-kpi/.test(html)) {
      failures.push(`${page}: ana sayfa sahte KPI görseli`);
    }
    if (!html.includes('premium-card__shot')) {
      failures.push(`${page}: ana sayfa gerçek ekran görüntüsü yok`);
    }
  }
}

if (failures.length > 0) {
  console.error('SMOKE TEST KALDI');
  for (const f of failures) console.error('  - ' + f);
  process.exit(1);
}

console.log(`SMOKE TEST GEÇTİ — ${pages.length} sayfa render, kırık iç link yok`);
