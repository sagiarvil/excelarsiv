#!/usr/bin/env node
import { spawnSync } from 'node:child_process';
import { readFileSync } from 'node:fs';

const PROTECTED = Object.freeze({
  'src/pages/index.astro': '3ab70d018c2627dfe7bc72512e3e7dffedca5a55',
  'src/pages/sablonlar.astro': 'a76ff05dec7d502562984a6bbf568da2a52516ff',
  'src/components/SiteHeader.astro': '87a854eccb4cbee9fc931f7f68b017983b3c71d5',
  'src/components/SiteFooter.astro': '8d87389b9f8c83e14c5d75cce37602a93eebc971',
  'src/layouts/CommerceLayout.astro': 'babd50f8021082e5e577f255c723f11c08e82e2a',
  'src/layouts/WorkbookLayout.astro': '4a77c4e32333543c1361bc1b1ad6b3e546d54b47',
  'src/styles/global.css': '1d8b5673d8f77ba3e8a29446ec884d0fc91b2569',
  'src/styles/home-native-info-hard-color-v33.css': 'd739fd58e4da62ca4f31f1f1327b6206ee9c21da',
  'public/images/excel-logo.png': 'fb85f04d742b13f7fe3a057fedba740e013e6b7a',
  'public/images/brand/excelarsiv-header-logo.png': 'fedfef196954861df583c2a0ff2aed8dc8fe496b',
});

function gitBlob(path) {
  const r = spawnSync('git', ['hash-object', path], {
    encoding: 'utf8',
    env: { ...process.env, HOME: '/tmp', GIT_CONFIG_NOSYSTEM: '1' },
  });
  if (r.status !== 0) return null;
  return String(r.stdout || '').trim();
}

const failures = [];
for (const [path, expected] of Object.entries(PROTECTED)) {
  const actual = gitBlob(path);
  if (actual !== expected) failures.push({ path, expected, actual: actual || 'MISSING' });
}

const special = readFileSync('src/pages/ozel-excel-sistemleri.astro', 'utf8');
const forbiddenRefs = [
  'SiteHeader',
  'SiteFooter',
  'CommerceLayout',
  'WorkbookLayout',
  '/images/brand/excelarsiv-header-logo.png',
];
for (const token of forbiddenRefs) {
  if (special.includes(token)) failures.push({ path: 'src/pages/ozel-excel-sistemleri.astro', expected: `must not reference ${token}`, actual: 'REFERENCE_FOUND' });
}

const hardColorCss = readFileSync('src/styles/home-native-info-hard-color-v33.css', 'utf8');
for (const token of [
  'body.ea-home-color-v3 .native-info--home',
  '--hm-green:#0a914a',
  '--hm-blue:#176fe5',
  '--hm-amber:#ee9d00',
  '--hm-coral:#f05a47',
  '.native-info__core',
  '.native-info__outcomes article:nth-child(4)',
  '@media(max-width:620px)',
  '@media(prefers-reduced-motion:reduce)',
]) {
  if (!hardColorCss.includes(token)) failures.push({ path: 'src/styles/home-native-info-hard-color-v33.css', expected: `must contain ${token}`, actual: 'TOKEN_MISSING' });
}

const enterpriseInjector = readFileSync('scripts/ci/inject-enterprise-light-color-suite-v3.mjs', 'utf8');
for (const token of [
  "homeHardColorCssSource = path.resolve('src/styles/home-native-info-hard-color-v33.css')",
  "homeHardColorStyleId = 'home-native-info-hard-color-v33'",
  "route.label === 'home'",
  'homepage-only hard-color v3.3',
]) {
  if (!enterpriseInjector.includes(token)) failures.push({ path: 'scripts/ci/inject-enterprise-light-color-suite-v3.mjs', expected: `must contain ${token}`, actual: 'TOKEN_MISSING' });
}

const commerceLayout = readFileSync('src/layouts/CommerceLayout.astro', 'utf8');
if (commerceLayout.includes('home-native-info-hard-color-v33.css')) {
  failures.push({ path: 'src/layouts/CommerceLayout.astro', expected: 'homepage hard-color CSS must not be globally imported', actual: 'GLOBAL_IMPORT_FOUND' });
}

const pkg = JSON.parse(readFileSync('package.json', 'utf8'));
const buildScript = String(pkg?.scripts?.build || '');
if (!buildScript.includes('PUBLIC_TEMPLATE_CARD_VARIANT=stable astro build')) {
  failures.push({
    path: 'package.json#scripts.build',
    expected: 'production build must force PUBLIC_TEMPLATE_CARD_VARIANT=stable',
    actual: buildScript || 'MISSING',
  });
}

if (failures.length) {
  console.error('PROTECTED SURFACE CONTRACT BLOCKED');
  console.error('Ana sayfa ve /sablonlar yüzeyleri özel sistem sayfasından ve deneysel katalog varyantından izole edilmiştir.');
  for (const f of failures) console.error(`- ${f.path}: expected ${f.expected}, actual ${f.actual}`);
  console.error('Bu yüzeylerden biri bilinçli olarak yeniden tasarlanacaksa koruma baselineı ayrı ve açık bir PR ile güncellenmelidir.');
  process.exit(1);
}

console.log(`PROTECTED SURFACE CONTRACT PASS — ${Object.keys(PROTECTED).length} protected paths + special-page isolation + stable catalog variant + homepage-only hard-color map`);
