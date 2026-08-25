#!/usr/bin/env node
import { existsSync, readFileSync, readdirSync, statSync, writeFileSync } from 'node:fs';
import { join, normalize, resolve } from 'node:path';

const DIST = resolve('dist');
const PRODUCT_DIR = join(DIST, 'sablon');
// Product detail pages are performance-critical and currently ship their
// layout/slug CSS as multiple render-blocking requests. Inlining the complete
// product-page CSS set removes those round trips without weakening any
// Lighthouse budget. The largest current bundle is ~130 KiB uncompressed.
const MAX_INLINE_BYTES = 160 * 1024;
const MARKER = 'data-inline-product-css';

function walk(dir) {
  const out = [];
  if (!existsSync(dir)) return out;
  for (const name of readdirSync(dir)) {
    const full = join(dir, name);
    const stat = statSync(full);
    if (stat.isDirectory()) out.push(...walk(full));
    else if (name === 'index.html') out.push(full);
  }
  return out;
}

function safeAssetPath(href) {
  if (!href.startsWith('/_astro/') || !href.endsWith('.css')) return null;
  const path = normalize(join(DIST, href.slice(1)));
  if (!path.startsWith(DIST) || !existsSync(path)) return null;
  return path;
}

let pages = 0;
let linksInlined = 0;
let bytesInlined = 0;

for (const file of walk(PRODUCT_DIR)) {
  let html = readFileSync(file, 'utf8');
  if (html.includes(MARKER)) continue;

  let changed = false;
  html = html.replace(/<link\b([^>]*\brel=["']stylesheet["'][^>]*\bhref=["']([^"']+\.css)["'][^>]*)>/gi, (tag, attrs, href) => {
    const asset = safeAssetPath(href);
    if (!asset) return tag;
    const size = statSync(asset).size;
    if (size > MAX_INLINE_BYTES) return tag;
    const css = readFileSync(asset, 'utf8');
    changed = true;
    linksInlined++;
    bytesInlined += size;
    return `<style ${MARKER} data-source="${href}">${css}</style>`;
  });

  if (changed) {
    writeFileSync(file, html);
    pages++;
  }
}

if (!pages) {
  console.error('Product CSS inline gate: no product HTML updated.');
  process.exit(1);
}

console.log(`PRODUCT CSS INLINE PASS — pages=${pages}, links=${linksInlined}, bytes=${bytesInlined}, max=${MAX_INLINE_BYTES}`);
