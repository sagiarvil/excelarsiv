import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, readdirSync } from 'node:fs';

const catalog = JSON.parse(readFileSync('commerce/catalog.json', 'utf8'));
const templateDir = 'src/content/templates';

function mdxPrice(slug: string): number {
  const source = readFileSync(`${templateDir}/${slug}.mdx`, 'utf8');
  const match = source.match(/^priceTL:\s*(\d+)\s*$/m);
  assert.ok(match, `${slug}: priceTL missing from MDX`);
  return Number(match[1]);
}

test('every commerce product price matches its tier and MDX SSOT', () => {
  for (const [slug, product] of Object.entries<any>(catalog.products)) {
    const tier = catalog.tiers[product.tier];
    assert.ok(tier, `${slug}: unknown commerce tier ${product.tier}`);
    assert.equal(mdxPrice(slug), tier.priceTL, `${slug}: MDX price drifted from commerce tier`);
  }
});

test('homepage product count, minimum price and visible product prices are derived from template data', () => {
  const home = readFileSync('src/pages/index.astro', 'utf8');
  assert.match(home, /const templates = await getAllTemplates\(\)/);
  assert.match(home, /const productCount = templates\.length/);
  assert.match(home, /Math\.min\(\.\.\.templates\.map\(\(item\) => item\.priceTL\)\)/);
  assert.match(home, /item\.priceTL\.toLocaleString\(['"]tr-TR['"]\)/);
  assert.doesNotMatch(home, /price:\s*['"]\d[\d.]*\s*TL/i);
  assert.doesNotMatch(home, /(?:^|[^\w])(?:2490|2\.490|990)\s*TL/i);
});

test('decision pages render visible and JSON-LD prices from TemplateViewModel SSOT', () => {
  const page = readFileSync('src/pages/karar/[slug].astro', 'utf8');
  const registry = readFileSync('src/data/kararPages.ts', 'utf8');
  assert.match(page, /const templates = await getAllTemplates\(\)/);
  assert.match(page, /price:\s*item\.priceTL/);
  assert.match(page, /primary\.priceTL\.toLocaleString/);
  assert.match(page, /item\.priceTL\.toLocaleString/);
  assert.doesNotMatch(registry, /\bprice(?:TL)?\s*:\s*\d+/i);
});

test('catalog count and minimum price are derived from current templates', () => {
  const catalogPage = readFileSync('src/pages/sablonlar.astro', 'utf8');
  assert.match(catalogPage, /templates\.length/);
  assert.match(catalogPage, /Math\.min\(\.\.\.templates\.map\(\(template\) => template\.priceTL\)\)/);
  assert.doesNotMatch(catalogPage, /50\s*(?:satışta sistem|ürün)/i);
  assert.doesNotMatch(catalogPage, /990\s*TL\s*başlayan fiyat/i);
});

test('every template MDX participating in commerce has a matching catalog product', () => {
  const mdxSlugs = readdirSync(templateDir).filter((name) => name.endsWith('.mdx')).map((name) => name.slice(0, -4));
  for (const slug of Object.keys(catalog.products)) assert.ok(mdxSlugs.includes(slug), `${slug}: commerce product missing MDX`);
});
