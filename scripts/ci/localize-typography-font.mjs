#!/usr/bin/env node
import { readdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const distDir = 'dist';
const externalFontPattern = /<link\s+data-ea-typography-font\s+href="https:\/\/fonts\.googleapis\.com\/css2\?family=Inter:[^"]+"\s+rel="stylesheet">/g;
const interToken = '--ea-font-sans:Inter,ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;';
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
  after = after.replaceAll(interToken, localToken);
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
  if (html.includes('fonts.googleapis.com/css2?family=Inter')) {
    throw new Error(`LOCAL TYPOGRAPHY GATE: external Inter blocker remains in ${path}`);
  }
  if (html.includes('--ea-font-sans:Inter,')) {
    throw new Error(`LOCAL TYPOGRAPHY GATE: Inter token remains in ${path}`);
  }
}

console.log(`LOCAL TYPOGRAPHY GATE PASS — ${files.length} HTML sayfası self-hosted Manrope zincirine alındı; Google Fonts render blocker kaldırıldı.`);
