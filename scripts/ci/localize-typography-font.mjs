#!/usr/bin/env node
import { readdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const distDir = 'dist';
// Typography mandate bazı sayfalarda attribute sırasını/ek attribute'ları değiştirebiliyor.
// Dış Inter stylesheet'ini tag içindeki attribute sırasından bağımsız olarak kaldır.
const externalFontPattern = /<link\b[^>]*href=["']https:\/\/fonts\.googleapis\.com\/css2\?family=Inter[^"']*["'][^>]*>/gi;
// Post-process zincirindeki farklı typography blokları aynı font stack'ini küçük yazım
// farklarıyla üretebilir. Yalnız --ea-font-sans custom property değerini normalize et;
// serbest metin veya başka CSS tokenlarını değiştirme.
const interTokenPattern = /--ea-font-sans\s*:\s*["']?Inter["']?\s*,\s*ui-sans-serif\s*,\s*system-ui\s*,\s*-apple-system\s*,\s*BlinkMacSystemFont\s*,\s*["']Segoe UI["']\s*,\s*Arial\s*,\s*sans-serif\s*;/gi;
const localToken = '--ea-font-sans:"Manrope",ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;';

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
  after = after.replace(interTokenPattern, localToken);
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
