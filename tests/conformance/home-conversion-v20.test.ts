import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = resolve(fileURLToPath(new URL('../../', import.meta.url)));

test('homepage conversion V20 preserves Size Özel funnel while restoring the original homepage hero', () => {
  const pkg = JSON.parse(readFileSync(resolve(ROOT, 'package.json'), 'utf8'));
  const script = readFileSync(resolve(ROOT, 'scripts/ci/finalize-home-conversion-v20.mjs'), 'utf8');
  const hero = readFileSync(resolve(ROOT, 'src/components/home/HeroSection.astro'), 'utf8');
  const css = readFileSync(resolve(ROOT, 'src/styles/home-conversion-v20.css'), 'utf8');
  const build = String(pkg.scripts?.build ?? '');
  const desktopPos = build.indexOf('finalize-home-desktop-premium-v19.mjs');
  const conversionPos = build.indexOf('finalize-home-conversion-v20.mjs');
  const retiredGuardPos = build.indexOf('assert-no-retired-whatsapp-v17.mjs');

  assert.ok(desktopPos >= 0 && conversionPos > desktopPos, 'V20 must run after V19 so the original hero reset is authoritative');
  assert.ok(retiredGuardPos > conversionPos, 'global retired-number guard must run after V20');

  assert.match(script, /Hazır sistem/);
  assert.match(script, /Size Özel/);
  assert.match(script, /İşletmenize Özel Sistemi İnceleyin/);
  assert.match(script, /home_custom_system_deep/);
  assert.match(script, /home_custom_whatsapp/);
  assert.match(script, /905393333303/);
  assert.match(css, /custom-build-v20/);

  assert.match(hero, /class="hero-mobile-copy"/);
  assert.match(hero, /srcset="\/images\/hero\.jpg"/);
  assert.match(hero, /aspect-ratio:\s*3\s*\/\s*1/);
  assert.match(hero, /İşletmeler İçin Finansal Karar ve Excel Sistemleri/);
  assert.match(hero, /Hazır Finansal Sistemleri İnceleyin/);
  assert.match(hero, /İşletmenize Özel Sistem Kuralım/);
  assert.match(hero, /class="hero-route-actions"/);
  assert.match(hero, /home_ready_systems/);
  assert.match(hero, /home_custom_system/);
  assert.doesNotMatch(hero, /class="hero-panel"/);
  assert.doesNotMatch(hero, /class="hero-copy"/);

  assert.match(script, /home-original-hero-v21-css/);
  assert.match(script, /original homepage hero contract missing/);
  assert.match(script, /dual hero CTA contract missing/);
  assert.match(script, /redesigned split hero leaked back into homepage/);
  assert.match(script, /deceptive urgency language detected/);
  assert.match(script, /\.test\(html\)/);
});
