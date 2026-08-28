#!/usr/bin/env node
import { readdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const distDir = 'dist';
// Typography mandate bazı sayfalarda attribute sırasını/ek attribute'ları değiştirebiliyor.
// Dış Inter stylesheet'ini tag içindeki attribute sırasından bağımsız olarak kaldır.
const externalFontPattern = /<link\b[^>]*href=["']https:\/\/fonts\.googleapis\.com\/css2\?family=Inter[^"']*["'][^>]*>/gi;
const interToken = '--ea-font-sans:Inter,ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;';
const localToken = '--ea-font-sans:"Manrope",ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;';
const specialLightPath = join(distDir, 'ozel-excel-sistemleri', 'index.html');

function walk(dir) {
  const out = [];
  for (const name of readdirSync(dir)) {
    const path = join(dir, name);
    if (statSync(path).isDirectory()) out.push(...walk(path));
    else if (path.endsWith('.html')) out.push(path);
  }
  return out;
}

function isSelfHostedSpecialLight(path, html) {
  return path === specialLightPath
    && html.includes('data-special-light-v1')
    && /\/fonts\/inter-(?:400|500|600|700)-latin-ext\.woff2/i.test(html);
}

const files = walk(distDir);
if (files.length < 10) throw new Error(`LOCAL TYPOGRAPHY GATE: suspicious HTML count ${files.length}`);

let transformed = 0;
for (const path of files) {
  const before = readFileSync(path, 'utf8');
  let after = before.replace(externalFontPattern, '');

  // Premium-light özel sayfa kendi self-hosted Inter yüzeyini tasarım sözleşmesi olarak korur.
  // Diğer tüm sayfalar site-wide Manrope tokenına normalize edilir.
  if (!isSelfHostedSpecialLight(path, after)) after = after.replaceAll(interToken, localToken);

  if (after !== before) {
    writeFileSync(path, after, 'utf8');
    transformed += 1;
  }
}

let compliant = 0;
for (const path of files) {
  const html = readFileSync(path, 'utf8');
  if (/fonts\.googleapis\.com\/css2\?family=Inter/i.test(html)) {
    throw new Error(`LOCAL TYPOGRAPHY GATE: external Inter blocker remains in ${path}`);
  }

  if (isSelfHostedSpecialLight(path, html)) {
    if (!html.includes('@font-face') || !html.includes('font-family:Inter')) {
      throw new Error(`LOCAL TYPOGRAPHY GATE: premium light self-hosted Inter contract incomplete in ${path}`);
    }
    compliant += 1;
    continue;
  }

  if (html.includes('--ea-font-sans:Inter,')) {
    throw new Error(`LOCAL TYPOGRAPHY GATE: Inter token remains in ${path}`);
  }
  compliant += 1;
}

if (compliant !== files.length) {
  throw new Error(`LOCAL TYPOGRAPHY GATE: expected ${files.length} compliant pages, found ${compliant}`);
}

console.log(`LOCAL TYPOGRAPHY GATE PASS — ${compliant}/${files.length} HTML sayfası self-hosted font zincirinde; ${transformed} sayfa build sırasında normalize edildi; Google Fonts render blocker yok.`);
