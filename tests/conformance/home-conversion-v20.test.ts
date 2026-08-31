import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = resolve(fileURLToPath(new URL('../../', import.meta.url)));

test('homepage conversion V20 preserves transparent dual-path persuasion', () => {
  const pkg = JSON.parse(readFileSync(resolve(ROOT, 'package.json'), 'utf8'));
  const script = readFileSync(resolve(ROOT, 'scripts/ci/finalize-home-conversion-v20.mjs'), 'utf8');
  const css = readFileSync(resolve(ROOT, 'src/styles/home-conversion-v20.css'), 'utf8');
  const build = String(pkg.scripts?.build ?? '');
  const desktopPos = build.indexOf('finalize-home-desktop-premium-v19.mjs');
  const conversionPos = build.indexOf('finalize-home-conversion-v20.mjs');
  const retiredGuardPos = build.indexOf('assert-no-retired-whatsapp-v17.mjs');
  assert.ok(desktopPos >= 0 && conversionPos > desktopPos, 'V20 must run after V19');
  assert.ok(retiredGuardPos > conversionPos, 'global retired-number guard must run after V20');
  assert.match(script, /Hazır sistem/);
  assert.match(script, /Size özel/);
  assert.match(script, /İşletmenize Özel Sistemi İnceleyin/);
  assert.match(script, /home_custom_system_deep/);
  assert.match(script, /home_custom_whatsapp/);
  assert.match(script, /905393333303/);
  assert.match(css, /custom-build-v20/);
  assert.doesNotMatch(script, /stokta son|son \d+|geri sayım|sadece bugün/i);
  assert.doesNotMatch(css, /stokta son|son \d+|geri sayım|sadece bugün/i);
});
