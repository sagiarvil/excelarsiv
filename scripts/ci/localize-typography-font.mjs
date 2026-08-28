#!/usr/bin/env node
import { readdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const distDir = 'dist';
// Typography mandate bazı sayfalarda attribute sırasını/ek attribute'ları değiştirebiliyor.
// Dış Inter stylesheet'ini tag içindeki attribute sırasından bağımsız olarak kaldır.
const externalFontPattern = /<link\b[^>]*href=["']https:\/\/fonts\.googleapis\.com\/css2\?family=Inter[^"']*["'][^>]*>/gi;
// Yalnız typography custom property başlangıcını normalize eder. Böylece minifier veya
// farklı post-process yazımları stack'in devamını değiştirse bile dış Inter bağımlılığı kalmaz.
const interPropertyPrefix = /--ea-font-sans\s*:\s*["']?Inter["']?\s*,/gi;
const localPropertyPrefix = '--ea-font-sans:"Manrope",';

function walk(dir) {
  const out = [];
  for (const name of readdirSync(dir)) {
    const path = join(dir, name);
    if (statSync(path).isDirectory()) out.push(...walk(path));
    else if (path.endsWith('.html')) out.push(path);
  }
  return out;
}

const files = walk(distDir);
if (files.length < 10) throw new Error(`LOCAL TYPOGRAPHY GATE: suspicious HTML count ${files.length}`);

let changed = 0;
for (const path of files) {
  const before = readFileSync(path, 'utf8');
  let after = before.replace(externalFontPattern, '');
  after = after.replace(interPropertyPrefix, localPropertyPrefix);
  if (after !== before) {
    writeFileSync(path, after, 'utf8');
    changed += 1;
  }
}

if (changed !== files.length) {
  throw new Error(`LOCAL TYPOGRAPHY GATE: expected ${files.length} localized pages, changed ${changed}`);
}

for (const path of files) {
  const html = readFileSync(path, 'utf8');
  if (/fonts\.googleapis\.com\/css2\?family=Inter/i.test(html)) {
    throw new Error(`LOCAL TYPOGRAPHY GATE: external Inter blocker remains in ${path}`);
  }
  if (/--ea-font-sans\s*:\s*["']?Inter["']?\s*,/i.test(html)) {
    throw new Error(`LOCAL TYPOGRAPHY GATE: Inter token remains in ${path}`);
  }
}

console.log(`LOCAL TYPOGRAPHY GATE PASS — ${files.length} HTML sayfası self-hosted Manrope zincirine alındı; Google Fonts render blocker kaldırıldı.`);
