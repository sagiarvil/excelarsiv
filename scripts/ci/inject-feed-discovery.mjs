#!/usr/bin/env node
/**
 * RSS 2.0 & Atom 1.0 Autodiscovery Enjektörü
 * Tüm HTML sayfalarının <head> etiketine RSS ve Atom autodiscovery linklerini enjekte eder.
 */

import { readdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const DIST = 'dist';
const RSS_LINK = '<link rel="alternate" type="application/rss+xml" title="Excel Arşiv RSS 2.0 Beslemesi" href="https://excelarsiv.com/rss.xml" />';
const ATOM_LINK = '<link rel="alternate" type="application/atom+xml" title="Excel Arşiv Atom 1.0 Beslemesi" href="https://excelarsiv.com/atom.xml" />';

function walk(dir) {
  const out = [];
  for (const name of readdirSync(dir)) {
    const full = join(dir, name);
    const stat = statSync(full);
    if (stat.isDirectory()) out.push(...walk(full));
    else if (name.endsWith('.html')) out.push(full);
  }
  return out;
}

const htmlFiles = walk(DIST);
let updatedCount = 0;

for (const filePath of htmlFiles) {
  let content = readFileSync(filePath, 'utf8');
  let modified = false;

  if (!content.includes('href="https://excelarsiv.com/rss.xml"') && !content.includes('href="/rss.xml"')) {
    if (content.includes('</head>')) {
      content = content.replace('</head>', `  ${RSS_LINK}\n  ${ATOM_LINK}\n</head>`);
      modified = true;
    }
  }

  if (modified) {
    writeFileSync(filePath, content, 'utf8');
    updatedCount++;
  }
}

console.log(`FEED AUTODISCOVERY PASS — ${updatedCount} HTML sayfasına RSS 2.0 & Atom 1.0 linkleri enjekte edildi.`);
