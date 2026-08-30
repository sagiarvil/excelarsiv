#!/usr/bin/env node
import { readdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const distDir = 'dist';
const externalFontPattern = /<link\b[^>]*href=["']https:\/\/fonts\.googleapis\.com\/css2\?family=Inter[^"']*["'][^>]*>/gi;
const interToken = '--ea-font-sans:Inter,ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;';
const localToken = '--ea-font-sans:"Manrope",ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;';
const specialLightPath = join(distDir, 'ozel-excel-sistemleri', 'index.html');
const specialLightSourcePath = join('src', 'pages', 'ozel-excel-sistemleri.astro');

function walk(dir) {
  const out = [];
  for (const name of readdirSync(dir)) {
    const path = join(dir, name);
    if (statSync(path).isDirectory()) out.push(...walk(path));
    else if (path.endsWith('.html')) out.push(path);
  }
  return out;
}

function isSpecialRoute(path) {
  return path === specialLightPath;
}

function isSelfHostedSpecialLight(path, html) {
  return isSpecialRoute(path) && html.includes('data-special-light-v1');
}

// Özel sayfanın font kontratı kaynak dosyada doğrulanır. V5 itibarıyla sayfa Apple'ın
// sistem font zincirini kullanabilir; bu modda harici/proprietary font dosyası bundling'i yoktur.
const specialSource = readFileSync(specialLightSourcePath, 'utf8');
const sourceOwnedApple = specialSource.includes('data-special-source-v5')
  && specialSource.includes('--font:-apple-system,BlinkMacSystemFont')
  && specialSource.includes('"SF Pro Text"')
  && specialSource.includes('"SF Pro Display"');
const livingWorkbookSource = !sourceOwnedApple
  && specialSource.includes('class="workbook"')
  && specialSource.includes('font-family:Manrope');

const requiredSpecialFontTokens = sourceOwnedApple
  ? [
      'data-special-source-v5',
      '--font:-apple-system,BlinkMacSystemFont',
      '"SF Pro Text"',
      '"SF Pro Display"',
      'font-size:17px',
    ]
  : livingWorkbookSource
    ? [
        '@font-face',
        'font-family:Manrope',
        "font-family:'IBM Plex Mono'",
        "/fonts/manrope-latin-ext.woff2",
        "/fonts/manrope-latin.woff2",
        "/fonts/ibm-plex-mono-500-latin-ext.woff2",
        "/fonts/ibm-plex-mono-500-latin.woff2",
      ]
    : [
        '@font-face',
        'font-family:Inter',
        "/fonts/inter-400-latin-ext.woff2",
        "/fonts/inter-500-latin-ext.woff2",
        "/fonts/inter-600-latin-ext.woff2",
        "/fonts/inter-700-latin-ext.woff2",
      ];

for (const token of requiredSpecialFontTokens) {
  if (!specialSource.includes(token)) {
    throw new Error(`LOCAL TYPOGRAPHY GATE: special source font contract missing: ${token}`);
  }
}
if (/fonts\.googleapis\.com\/css2\?family=Inter/i.test(specialSource)) {
  throw new Error('LOCAL TYPOGRAPHY GATE: special source may not load Inter from Google Fonts');
}
if (sourceOwnedApple && /@font-face|manrope|ibm plex mono/i.test(specialSource)) {
  throw new Error('LOCAL TYPOGRAPHY GATE: Apple system-font special source may not bundle legacy typography');
}

const files = walk(distDir);
if (files.length < 10) throw new Error(`LOCAL TYPOGRAPHY GATE: suspicious HTML count ${files.length}`);

let transformed = 0;
for (const path of files) {
  const before = readFileSync(path, 'utf8');
  let after = before.replace(externalFontPattern, '');

  // V5 özel sayfa kaynak tarafından yönetilir. Bu aşamada özel route'a Manrope/Inter
  // normalizasyonu uygulanmaz; final source gate daha sonra route'u byte-level yeniden kurar.
  if (!isSpecialRoute(path)) after = after.replaceAll(interToken, localToken);

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

  if (isSpecialRoute(path)) {
    if (!sourceOwnedApple && !isSelfHostedSpecialLight(path, html)) {
      throw new Error('LOCAL TYPOGRAPHY GATE: special route lost supported typography namespace');
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

const specialMode = sourceOwnedApple ? 'Apple-system-v5' : livingWorkbookSource ? 'Manrope+IBM-Plex-Mono' : 'legacy-Inter';
console.log(`LOCAL TYPOGRAPHY GATE PASS — ${compliant}/${files.length} HTML sayfası font kontratında; ${transformed} sayfa build sırasında normalize edildi; special route font mode=${specialMode}; Google Fonts render blocker yok.`);
