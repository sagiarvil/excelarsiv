#!/usr/bin/env node
import { spawnSync } from 'node:child_process';
import { readFileSync } from 'node:fs';

const PROTECTED = Object.freeze({
  'src/pages/index.astro': 'a6d57c582218fe31cf30a8f87754c2f3b5247443',
  'src/pages/sablonlar.astro': '2530313191634ebac770cf61ec7b40870ebc6e3a',
  'src/components/SiteHeader.astro': 'd4181e23e96920cd6cbcf55bb569b0087b18c1dc',
  'src/components/SiteFooter.astro': 'bb319db54455fe133f0cd3db9dd5a1246b7dbcd6',
  'src/layouts/CommerceLayout.astro': '996f1eb58ed9b1fbe9973becac36b00a2df993be',
  'src/layouts/WorkbookLayout.astro': '9c1eef5695e15e98e2bae4fafdba1283d31de553',
  'src/styles/global.css': '38709d0884f5ec933936135857ef19dd47ff6b4a',
  'src/styles/home-native-info-hard-color-v33.css': 'd739fd58e4da62ca4f31f1f1327b6206ee9c21da',
  'public/images/excel-logo.png': 'c58227931a0e504348ac9d5e3524a6beae3b0149',
  'public/images/brand/excelarsiv-header-logo.png': 'c58227931a0e504348ac9d5e3524a6beae3b0149',
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
