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

test('homepage proof prices are derived from template data, never hardcoded', () => {
  const home = readFileSync('src/pages/index.astro', 'utf8');
  const proof = readFileSync('src/components/home/CommerceProofGallery.astro', 'utf8');
  assert.match(home, /CommerceProofGallery templates=\{allTemplates\}/);
  assert.doesNotMatch(home, /ProofGallery from ['"]\.\.\/components\/home\/ProofGallery\.astro/);
  assert.match(proof, /template\.priceTL\.toLocaleString/);
  assert.doesNotMatch(proof, /price:\s*['"]\d[\d.]*\s*TL/i);
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
