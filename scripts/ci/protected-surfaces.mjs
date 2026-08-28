#!/usr/bin/env node
import { spawnSync } from 'node:child_process';
import { readFileSync } from 'node:fs';

const PROTECTED = Object.freeze({
  'src/pages/index.astro': '43a7f9588ccf489a8a2254f3780d19cead32c557',
  'src/pages/sablonlar.astro': 'c80e4e4344144ba6448d3d863480b58c92995fa5',
  'src/components/SiteHeader.astro': 'be35bf44111f3625b3ee774926574da8f89ec07d',
  'src/components/SiteFooter.astro': '98b6523588eef8df262dd49e4d5f380414329861',
  'src/layouts/CommerceLayout.astro': '6705e8e714274869ae039e4667606e25408a28a1',
  'src/layouts/WorkbookLayout.astro': '4a77c4e32333543c1361bc1b1ad6b3e546d54b47',
  'src/styles/global.css': '68183699f7eda295db71525dc17ab44976ebc608',
  'src/styles/home-native-info-hard-color-v33.css': 'd739fd58e4da62ca4f31f1f1327b6206ee9c21da',
  'public/images/excel-logo.png': '024ebb12404fa297ba04e4afa1834acf1769f442',
  'public/images/brand/excelarsiv-header-logo.png': 'fedfef196954861df583c2a0ff2aed8dc8fe496b',
});

function gitBlob(path) {
  const r = spawnSync('git', ['hash-object', path], { encoding: 'utf8' });
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

console.log(`PROTECTED SURFACE CONTRACT PASS — ${Object.keys(PROTECTED).length} protected paths + special-page isolation + stable catalog variant + hard-color home map`);