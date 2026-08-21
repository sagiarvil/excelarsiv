import test from 'node:test';
import assert from 'node:assert/strict';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { loadSeoRegistry, registryParity, sourceIndexableRoutes } from '../../../scripts/seo/registry-source-parity.ts';

const ROOT = resolve(fileURLToPath(new URL('../../../', import.meta.url)));
void ROOT;

test('registry live route set equals source indexable route set', () => {
  const registry = loadSeoRegistry();
  const sourceRoutes = sourceIndexableRoutes();
  const result = registryParity(registry, sourceRoutes);
  assert.ok(sourceRoutes.length > 0);
  assert.deepEqual(result.missing, []);
  assert.deepEqual(result.extra, []);
});

test('registry parity fixture detects missing and extra routes', () => {
  const registry = { records: [{ route:'/', status:'live' }, { route:'/ghost', status:'live' }] };
  const result = registryParity(registry, ['/', '/expected']);
  assert.deepEqual(result.missing, ['/expected']);
  assert.deepEqual(result.extra, ['/ghost']);
});
