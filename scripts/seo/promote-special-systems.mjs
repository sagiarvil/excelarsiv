#!/usr/bin/env node
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const DIST = 'dist';
const SPECIAL_URL = 'https://excelarsiv.com/ozel-excel-sistemleri';
const files = {
  short: join(DIST, 'llms.txt'),
  full: join(DIST, 'llms-full.txt'),
  pages: join(DIST, 'sitemap-pages.xml'),
  index: join(DIST, 'sitemap.xml'),
};

for (const [name, path] of Object.entries(files)) {
  if (!existsSync(path)) throw new Error(`SPECIAL SYSTEMS DISCOVERY: missing ${name} artifact at ${path}`);
}

function count(text, needle) {
  return text.split(needle).length - 1;
}

function addPromotedLine(text, anchor, line, label) {
  if (text.includes(line)) return text;
  const occurrences = count(text, anchor);
  if (occurrences !== 1) throw new Error(`SPECIAL SYSTEMS DISCOVERY: ${label} anchor count=${occurrences}`);
  return text.replace(anchor, `${anchor}\n${line}`);
}

let short = readFileSync(files.short, 'utf8');
short = addPromotedLine(
  short,
  '- Katalog: https://excelarsiv.com/sablonlar',
  `- İhtiyaca özel Excel sistemleri: ${SPECIAL_URL}`,
  'llms.txt catalog',
);
writeFileSync(files.short, short, 'utf8');

let full = readFileSync(files.full, 'utf8');
full = addPromotedLine(
  full,
  '- Katalog: https://excelarsiv.com/sablonlar',
  `- İhtiyaca özel Excel sistemleri: ${SPECIAL_URL}`,
  'llms-full.txt catalog',
);
writeFileSync(files.full, full, 'utf8');

const pages = readFileSync(files.pages, 'utf8');
const specialCount = count(pages, `<loc>${SPECIAL_URL}</loc>`);
if (specialCount !== 1) {
  throw new Error(`SPECIAL SYSTEMS DISCOVERY: sitemap-pages.xml must contain special systems exactly once; found ${specialCount}`);
}

const index = readFileSync(files.index, 'utf8');
if (!index.includes('<loc>https://excelarsiv.com/sitemap-pages.xml</loc>')) {
  throw new Error('SPECIAL SYSTEMS DISCOVERY: sitemap index does not reference sitemap-pages.xml');
}

if (!readFileSync(files.short, 'utf8').includes(`- İhtiyaca özel Excel sistemleri: ${SPECIAL_URL}`)) {
  throw new Error('SPECIAL SYSTEMS DISCOVERY: llms.txt promotion missing');
}
if (!readFileSync(files.full, 'utf8').includes(`- İhtiyaca özel Excel sistemleri: ${SPECIAL_URL}`)) {
  throw new Error('SPECIAL SYSTEMS DISCOVERY: llms-full.txt promotion missing');
}

console.log('SPECIAL SYSTEMS DISCOVERY PASS — promoted in llms.txt + llms-full.txt; sitemap-pages inclusion locked');
