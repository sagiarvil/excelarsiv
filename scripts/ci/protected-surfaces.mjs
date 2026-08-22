#!/usr/bin/env node
import { spawnSync } from 'node:child_process';
import { readFileSync } from 'node:fs';

const PROTECTED = Object.freeze({
  'src/pages/index.astro': '43a7f9588ccf489a8a2254f3780d19cead32c557',
  'src/pages/sablonlar.astro': 'a2e1512ffd6d2f352138b35b43622fba9b56ffbf',
  'src/components/SiteHeader.astro': '6a0b7903a2d3032f16aec808e156a9ec96fda8a7',
  'src/components/SiteFooter.astro': 'a2f5388773fc9e619910828c1b5fb6fd39022eb2',
  'src/layouts/CommerceLayout.astro': 'aec57770ec93c0d18f78c9d5cb68cabd6ce61b1d',
  'src/layouts/WorkbookLayout.astro': '4a77c4e32333543c1361bc1b1ad6b3e546d54b47',
  'src/styles/global.css': '68183699f7eda295db71525dc17ab44976ebc608',
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
  '/images/excel-logo.png',
  '/images/brand/excelarsiv-header-logo.png',
];
for (const token of forbiddenRefs) {
  if (special.includes(token)) failures.push({ path: 'src/pages/ozel-excel-sistemleri.astro', expected: `must not reference ${token}`, actual: 'REFERENCE_FOUND' });
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

console.log(`PROTECTED SURFACE CONTRACT PASS — ${Object.keys(PROTECTED).length} protected paths + special-page isolation + stable catalog variant`);
