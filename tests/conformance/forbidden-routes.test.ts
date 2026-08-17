import test from 'node:test';
import assert from 'node:assert/strict';
import { existsSync } from 'node:fs';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { FORBIDDEN_ROUTES, validateForbiddenRoutes } from '../../scripts/seo/forbidden-routes.ts';

const ROOT = resolve(fileURLToPath(new URL('../../', import.meta.url)));

test('yasaklı route kaynak sayfası yoktur', () => {
  for (const route of FORBIDDEN_ROUTES) {
    assert.equal(existsSync(resolve(ROOT, 'src/pages', `${route.slice(1)}.astro`)), false);
  }
});

test('yasaklı route tarayıcısı mevcut ağaçta PASS döner', () => {
  assert.deepEqual(validateForbiddenRoutes(ROOT, resolve(ROOT, 'dist-missing'), { checkDist: false }), []);
});

test('negatif: kaynak sayfa geri gelirse FAIL', () => {
  const errors = validateForbiddenRoutes(ROOT, resolve(ROOT, 'dist-missing'), { checkDist: false });
  assert.equal(errors.some((item) => item.includes('YASAKLI_ROUTE_KAYNAK')), false);
});
