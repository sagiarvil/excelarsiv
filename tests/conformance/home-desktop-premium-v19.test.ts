import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = resolve(fileURLToPath(new URL('../../', import.meta.url)));

test('desktop premium V19 stays desktop-only and runs after mobile V18', () => {
  const pkg = JSON.parse(readFileSync(resolve(ROOT, 'package.json'), 'utf8'));
  const finalizer = readFileSync(resolve(ROOT, 'scripts/ci/finalize-home-desktop-premium-v19.mjs'), 'utf8');
  const css = readFileSync(resolve(ROOT, 'src/styles/home-desktop-premium-v19.css'), 'utf8');
  const build = String(pkg.scripts?.build ?? '');
  const mobilePos = build.indexOf('finalize-mobile-premium-v18.mjs');
  const desktopPos = build.indexOf('finalize-home-desktop-premium-v19.mjs');
  assert.ok(mobilePos >= 0, 'V18 mobile finalizer must stay in build chain');
  assert.ok(desktopPos > mobilePos, 'V19 desktop finalizer must run after V18');
  assert.match(css, /@media\(min-width:1021px\)/);
  assert.doesNotMatch(css, /@media\(max-width:/);
  assert.doesNotMatch(css, /overflow-x\s*:\s*hidden/);
  assert.match(finalizer, /data-desktop-premium-v19/);
  assert.match(css, /\.hero-panel::after/);
  assert.match(css, /\.hero-search-rail/);
  assert.match(css, /\.finance-pillars/);
});
