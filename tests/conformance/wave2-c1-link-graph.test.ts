import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { analyzeLinkGraph } from '../../scripts/seo/link-graph.ts';

const ROOT = resolve(fileURLToPath(new URL('../../', import.meta.url)));

test('link graph counts incoming links and identifies orphan routes', () => {
  const pages = [
    { route: '/', html: '<a href="/a">A</a><a href="/b">B</a>' },
    { route: '/a', html: '<a href="/b">B</a>' },
    { route: '/b', html: '<a href="/a">A</a><a href="/">Home</a>' },
  ];
  const registry = {
    records: [
      { pageId: 'home', route: '/', status: 'live', type: 'home' },
      { pageId: 'a', route: '/a', status: 'live', type: 'other' },
      { pageId: 'b', route: '/b', status: 'live', type: 'other' },
    ],
  };
  const result = analyzeLinkGraph(pages, registry, 2);
  assert.equal(result.rows.find((row) => row.route === '/a')?.internalLinksIn, 2);
  assert.equal(result.rows.find((row) => row.route === '/b')?.internalLinksIn, 2);
  assert.equal(result.rows.find((row) => row.route === '/')?.internalLinksIn, 1);
  assert.deepEqual(result.orphans.map((row) => row.route), ['/']);
  assert.equal(result.suggestions[0]?.targetRoute, '/');
});

test('C1/CWV critical pages defer below-fold rendering and keep heavy mobile hero images off the critical path', () => {
  const home = readFileSync(resolve(ROOT, 'src/pages/index.astro'), 'utf8');
  const homeHero = readFileSync(resolve(ROOT, 'src/components/home/HeroSection.astro'), 'utf8');
  const productHero = readFileSync(resolve(ROOT, 'src/components/product/ProductHeroPremium.astro'), 'utf8');
  assert.match(home, /content-visibility:\s*auto/);
  assert.match(home, /contain-intrinsic-size:\s*auto\s+760px/);
  assert.match(home, /<HeroSection searchIndex=\{searchIndex\}\s*\/>/);
  assert.match(homeHero, /<source media="\(min-width: 721px\)" srcset="\/images\/hero\.jpg"/);
  assert.match(homeHero, /data:image\/gif;base64/);
  assert.match(homeHero, /\.hero-artwork\s*\{[\s\S]*aspect-ratio:\s*3\s*\/\s*1/);
  assert.match(homeHero, /@media \(max-width: 720px\)[\s\S]*\.hero-artwork\s*\{[\s\S]*display:\s*none/);
  assert.match(productHero, /source media="\(min-width: 761px\)" srcset=\{primary\.src\}/);
  assert.match(productHero, /\.product-hero__visual\{display:none\}/);
  assert.match(productHero, /product-hero__mobile-proof-link/);
  assert.match(productHero, /:global\(\.product-page > \.product-section\).*content-visibility:auto/);
});

test('C1/CWV seven-section homepage keeps decision core visible and defers below-fold work', () => {
  const home = readFileSync(resolve(ROOT, 'src/pages/index.astro'), 'utf8');
  const homeHero = readFileSync(resolve(ROOT, 'src/components/home/HeroSection.astro'), 'utf8');
  const heroSearch = readFileSync(resolve(ROOT, 'src/components/home/HeroSearch.astro'), 'utf8');
  const inlineSections = home.match(/<section\b/g) ?? [];
  const heroSections = homeHero.match(/<section\b/g) ?? [];
  assert.equal(inlineSections.length + heroSections.length, 7);
  assert.match(home, /class="decision-help" data-experience-stage/);
  assert.match(home, /homeDecisionSlugs/);
  assert.match(home, /href=\{`\/karar\/\$\{page\.slug\}`\}/);
  assert.match(home, /getTemplateSearchIndex/);
  assert.match(homeHero, /id="hero-search"/);
  assert.match(heroSearch, /input[\s\S]*type="search"/);
  assert.match(home, /--font-sans:\s*ui-sans-serif/);
  assert.match(home, /--font-mono-face:\s*ui-monospace/);
});
