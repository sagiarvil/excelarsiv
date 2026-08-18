import test from 'node:test';
import assert from 'node:assert/strict';
import { existsSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = resolve(fileURLToPath(new URL('../../', import.meta.url)));

test('satıştaki her ürünün premium katalog kapağı vardır', () => {
  const catalog = JSON.parse(readFileSync(resolve(ROOT, 'commerce/catalog.json'), 'utf8'));
  const slugs = Object.keys(catalog.products ?? {});
  const missing = slugs.filter((slug) => !existsSync(resolve(ROOT, 'public/images/kapak', `${slug}.webp`)));
  assert.equal(slugs.length, 50);
  assert.deepEqual(missing, []);
});

test('premium resolver mevcut kapak sözleşmesini ezmeden eklenmiştir', () => {
  const source = readFileSync(resolve(ROOT, 'src/lib/kapak.ts'), 'utf8');
  assert.equal(source.includes('export function premiumKapakUrl'), true);
  assert.equal(source.includes('/images/kapak/${slug}.webp'), true);
});

test('katalog kapak zinciri premium katmana bağlıdır', () => {
  const kapak = readFileSync(resolve(ROOT, 'src/lib/kapak.ts'), 'utf8');
  const candidates = [
    'src/components/catalog/KatalogUrunKarti.astro',
    'src/pages/sablonlar.astro',
    'src/pages/sablonlar/index.astro',
    'src/components/ProductCard.astro',
    'src/components/home/PremiumFeaturedTemplates.astro',
  ];
  const directSurfaceBinding = candidates
    .filter((p) => existsSync(resolve(ROOT, p)))
    .map((p) => readFileSync(resolve(ROOT, p), 'utf8'))
    .some((source) => source.includes('premiumKapakUrl'));

  // Bazı repo revizyonlarında katalog yüzeyi src/lib/kapak.ts içindeki mevcut
  // resolverı çağırır ve URL'yi karta prop olarak geçirir. Bu durumda premium
  // seçim mevcut resolver gövdesine fail-safe override olarak eklenir.
  const inPlaceResolverBinding =
    kapak.includes('PREMIUM_KAPAK_SLUGS.has(__premiumSlug)') &&
    kapak.includes('/images/kapak/${__premiumSlug}.webp');

  assert.equal(directSurfaceBinding || inPlaceResolverBinding, true);
});
