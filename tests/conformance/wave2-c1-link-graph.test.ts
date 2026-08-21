import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { analyzeLinkGraph } from '../../scripts/seo/link-graph.ts';

const ROOT = resolve(fileURLToPath(new URL('../../', import.meta.url)));

test('link graph counts incoming links and identifies orphan routes', () => {
  const pages = new Map([
    ['/', '<a href="/a">A</a>'],
    ['/a', '<a href="/b">B</a>'],
    ['/b', '<a href="/a">A</a><a href="/">Home</a>'],
  ]);
  const registry = [
    { route: '/', status: 'live' },
    { route: '/a', status: 'live' },
    { route: '/b', status: 'live' },
  ];
  const result = analyzeLinkGraph(pages, registry, 2);
  assert.equal(result.rows.find((row) => row.route === '/a')?.internalLinksIn, 2);
  assert.equal(result.rows.find((row) => row.route === '/b')?.internalLinksIn, 2);
  assert.equal(result.rows.find((row) => row.route === '/')?.internalLinksIn, 1);
  assert.deepEqual(result.orphans.map((row) => row.route), ['/']);
  assert.equal(result.suggestions[0]?.targetRoute, '/');
});

test('C1/CWV critical pages defer below-fold rendering and keep heavy mobile hero images off the critical path', () => {
  const home = readFileSync(resolve(ROOT, 'src/pages/index.astro'), 'utf8');
  const productHero = readFileSync(resolve(ROOT, 'src/components/product/ProductHeroPremium.astro'), 'utf8');
  assert.match(home, /content-visibility:\s*auto/);
  assert.match(home, /contain-intrinsic-size:\s*auto\s+760px/);
  assert.match(home, /<source media="\(min-width: 721px\)" srcset=\{heroProduct\.kapak\}/);
  assert.match(home, /data:image\/gif;base64/);
  assert.match(productHero, /source media="\(min-width: 761px\)" srcset=\{primary\.src\}/);
  assert.match(productHero, /\.product-hero__visual\{display:none\}/);
  assert.match(productHero, /product-hero__mobile-proof-link/);
  assert.match(productHero, /:global\(\.product-page > \.product-section\).*content-visibility:auto/);
});

test('C1/CWV seven-section homepage keeps decision core visible and defers below-fold work', () => {
  const home = readFileSync(resolve(ROOT, 'src/pages/index.astro'), 'utf8');
  const sections = home.match(/<section\b/g) ?? [];
  assert.equal(sections.length, 7);
  assert.match(home, /class="decision-help" data-experience-stage/);
  assert.match(home, /homeDecisionSlugs/);
  assert.match(home, /href=\{`\/karar\/\$\{page\.slug\}`\}/);
  assert.doesNotMatch(home, /input type="search"/);
  assert.match(home, /--font-sans:\s*ui-sans-serif/);
  assert.match(home, /--font-mono-face:\s*ui-monospace/);
});
