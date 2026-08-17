import test from 'node:test';
import assert from 'node:assert/strict';
import { existsSync, readFileSync, readdirSync } from 'node:fs';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { catalogScreenshotAlt, pickCatalogScreenshot } from '../../src/lib/catalog-screenshot.ts';

const ROOT = resolve(fileURLToPath(new URL('../../', import.meta.url)));

test('katalog ekranı harita yoksa hesap ızgarasını (-2) boş panoya tercih eder', () => {
  const picked = pickCatalogScreenshot([
    { src: '/screenshots/ornek-1.png', alt: 'NAKIT_GIRISLERI — Haftalık nakit girişi girişleri' },
    { src: '/screenshots/ornek-2.png', alt: 'HAFTALIK_PLAN — Haftalık plan' },
    { src: '/screenshots/ornek-3.png', alt: 'PANO sayfası — Yönetim panosu özeti' },
  ]);
  assert.equal(picked?.src, '/screenshots/ornek-2.png');
});

test('katalog ekranı dolu Excel ızgarasını (-1) seçer', () => {
  const picked = pickCatalogScreenshot([
    { src: '/screenshots/cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi-1.png', alt: 'Girdi' },
    { src: '/screenshots/cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi-2.png', alt: 'Hesap' },
    { src: '/screenshots/cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi-3.png', alt: 'PANO' },
  ]);
  assert.equal(picked?.src, '/screenshots/cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi-1.png');
});

test('katalog ekranı Proof-Demo üründe hesap karesini (-2) seçer', () => {
  const picked = pickCatalogScreenshot([
    { src: '/screenshots/13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi-1.png', alt: 'Girdi' },
    { src: '/screenshots/13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi-2.png', alt: 'Hesap' },
    { src: '/screenshots/13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi-3.png', alt: 'PANO' },
  ]);
  assert.equal(picked?.src, '/screenshots/13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi-2.png');
});

test('katalog ekranı tek karede o kareyi döndürür', () => {
  const picked = pickCatalogScreenshot([{ src: '/screenshots/tek.png', alt: 'Tek sayfa' }]);
  assert.equal(picked?.src, '/screenshots/tek.png');
});

test('kart alt metni sahte Canlı rozeti taşımaz', () => {
  assert.equal(catalogScreenshotAlt('13 Haftalık Nakit'), '13 Haftalık Nakit gerçek Excel karar ekranı');
  assert.equal(/Canlı/.test(catalogScreenshotAlt('13 Haftalık Nakit')), false);
});

test('her MDX ekran görüntüsü public/ altında durur', () => {
  const dir = resolve(ROOT, 'src/content/templates');
  const missing = [];
  for (const file of readdirSync(dir).filter((name) => name.endsWith('.mdx'))) {
    const source = readFileSync(resolve(dir, file), 'utf8');
    for (const match of source.matchAll(/src:\s*'(\/screenshots\/[^']+)'/g)) {
      const abs = resolve(ROOT, 'public', match[1].slice(1));
      if (!existsSync(abs)) missing.push(`${file}: ${match[1]}`);
    }
  }
  assert.deepEqual(missing, []);
});

test('listeleme ve sızıntı yüzeyleri tek kart bileşenine bağlı', () => {
  const grid = readFileSync(resolve(ROOT, 'src/components/catalog/TemplateGrid.astro'), 'utf8');
  const card = readFileSync(resolve(ROOT, 'src/components/catalog/KatalogUrunKarti.astro'), 'utf8');
  const product = readFileSync(resolve(ROOT, 'src/pages/sablon/[slug].astro'), 'utf8');
  const notFound = readFileSync(resolve(ROOT, 'src/pages/404.astro'), 'utf8');
  const home = readFileSync(resolve(ROOT, 'src/components/home/PremiumFeaturedTemplates.astro'), 'utf8');
  const visual = readFileSync(resolve(ROOT, 'src/components/ProductCard.astro'), 'utf8');
  const filter = readFileSync(resolve(ROOT, 'src/scripts/catalog-filter.ts'), 'utf8');
  const tools = readFileSync(resolve(ROOT, 'src/pages/excel-araclari.astro'), 'utf8');
  const legacy = readFileSync(resolve(ROOT, 'src/components/catalog/TemplateCard.astro'), 'utf8');

  assert.match(grid, /data-template-grid-wrap/);
  assert.match(grid, /KatalogUrunKarti/);
  assert.equal(/wa\.me/.test(grid), false);
  assert.match(card, /orderUrl=\{template\.shopierUrl\}/);
  assert.equal(/wa\.me/.test(card), false);
  assert.equal(/wa\.me/.test(visual), false);
  assert.match(product, /KatalogUrunKarti/);
  assert.equal(/TemplateCard/.test(product), false);
  assert.match(notFound, /KatalogUrunKarti/);
  assert.equal(/TemplateCard/.test(notFound), false);
  assert.match(home, /premium-card__shot/);
  assert.equal(/CatalogProductVisual/.test(home), false);
  assert.match(visual, /aspect-ratio:\s*4\s*\/\s*3/);
  assert.match(visual, /object-fit:\s*cover/);
  assert.match(visual, /catalogShotOdak/);
  assert.match(visual, /transform:\s*scale/);
  assert.match(visual, /card__workbook/);
  assert.match(visual, /border-radius:\s*0/);
  assert.equal(/●\s*Canlı|pv-kpi|product-visual/.test(visual), false);
  assert.match(filter, /data-template-grid-wrap/);
  assert.equal(/wa\.me/.test(tools), false);
  assert.equal(/wa\.me/.test(legacy), false);
  const shopierSrc = readFileSync(resolve(ROOT, 'src/lib/shopier.ts'), 'utf8');
  assert.match(shopierSrc, /export function shopierUrlForPrice/);
});
