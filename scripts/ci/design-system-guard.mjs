#!/usr/bin/env node
import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs';
import { extname, join, resolve } from 'node:path';

const root = process.cwd();
const failures = [];
const warnings = [];

function read(path) {
  const full = resolve(root, path);
  if (!existsSync(full)) {
    failures.push(`${path}: missing`);
    return '';
  }
  return readFileSync(full, 'utf8');
}

function requireText(content, path, needles) {
  for (const needle of needles) {
    if (!content.includes(needle)) failures.push(`${path}: required contract text missing -> ${needle}`);
  }
}

function walk(dir, out = []) {
  const full = resolve(root, dir);
  if (!existsSync(full)) return out;
  for (const entry of readdirSync(full)) {
    const path = join(full, entry);
    if (statSync(path).isDirectory()) walk(path, out);
    else out.push(path);
  }
  return out;
}

const design = read('DESIGN.md');
const agents = read('AGENTS.md');
const globalCss = read('src/styles/global.css');
const commerceLayout = read('src/layouts/CommerceLayout.astro');
const workbookLayout = read('src/layouts/WorkbookLayout.astro');
const pkgText = read('package.json');

requireText(design, 'DESIGN.md', [
  '# Excel Arşiv Design System Contract',
  '## 2. 21st.dev reference policy',
  'https://21st.dev/community/components/s/comparison',
  '## 3. Runtime design tokens',
  '## 5. Component contract',
  '## 7. Responsive contract — non-negotiable',
  'Zero page-level horizontal overflow',
  '320px',
  '1440px',
  '## 8. Accessibility contract',
  '## 13. Protected surfaces and change discipline',
  '## 14. QA checklist for every visual change',
  '## 16. Agent execution prompt',
  '## 17. Definition of done',
]);

requireText(agents, 'AGENTS.md', [
  'Read `DESIGN.md` completely.',
  '`DESIGN.md` is the canonical visual contract.',
  'npm run guard:design',
  'npm run build',
  'npm test',
]);

requireText(globalCss, 'src/styles/global.css', [
  '--xl-green:',
  '--xl-green-dark:',
  '--xl-green-tint:',
  '--xl-canvas:',
  '--xl-paper:',
  '--xl-ink:',
  '--font-sans:',
  '--font-mono:',
  '--t-display:',
  '--r-base:',
  '--space-4:',
  '--container-main:',
  '--container-wide:',
  '--page-pad-mobile:',
  '--section-gap-mobile:',
  '@media (prefers-reduced-motion: reduce)',
]);

requireText(commerceLayout, 'src/layouts/CommerceLayout.astro', [
  'width=device-width, initial-scale=1, viewport-fit=cover',
  'href="#icerik"',
]);
requireText(workbookLayout, 'src/layouts/WorkbookLayout.astro', [
  'width=device-width, initial-scale=1',
  'href="#icerik"',
]);

let pkg = null;
try {
  pkg = JSON.parse(pkgText);
} catch (error) {
  failures.push(`package.json: invalid JSON -> ${error.message}`);
}
if (pkg) {
  const guard = String(pkg?.scripts?.['guard:design'] || '');
  const build = String(pkg?.scripts?.build || '');
  if (guard !== 'node scripts/ci/design-system-guard.mjs') {
    failures.push('package.json#scripts.guard:design: must execute scripts/ci/design-system-guard.mjs');
  }
  if (!build.includes('npm run guard:design')) {
    failures.push('package.json#scripts.build: design guard must run in production build');
  }
}

// 21st.dev is reference-only. Runtime source must not depend on it.
for (const file of [...walk('src'), ...walk('public')]) {
  const ext = extname(file).toLowerCase();
  if (!['.astro', '.css', '.js', '.mjs', '.ts', '.tsx', '.html'].includes(ext)) continue;
  const content = readFileSync(file, 'utf8');
  if (content.includes('21st.dev')) {
    failures.push(`${file.replace(root + '/', '')}: 21st.dev must remain a design reference, not a runtime dependency`);
  }
}

// New overflow masks in page/component/layout source are not acceptable. The historical
// body overflow mask lives in protected global.css and is handled as an explicit legacy exception.
for (const dir of ['src/components', 'src/pages', 'src/layouts']) {
  for (const file of walk(dir)) {
    const ext = extname(file).toLowerCase();
    if (!['.astro', '.css'].includes(ext)) continue;
    const content = readFileSync(file, 'utf8');
    if (/overflow-x\s*:\s*hidden/i.test(content)) {
      failures.push(`${file.replace(root + '/', '')}: overflow-x:hidden masks layout defects; fix the local layout instead`);
    }
  }
}

if (/body\s*\{[\s\S]{0,1200}?overflow-x\s*:\s*hidden/i.test(globalCss)) {
  warnings.push('src/styles/global.css: legacy body overflow-x:hidden is still present; edited surfaces must not rely on it');
}

if (failures.length) {
  console.error('DESIGN SYSTEM CONTRACT BLOCKED');
  for (const failure of failures) console.error(`- ${failure}`);
  for (const warning of warnings) console.warn(`WARN: ${warning}`);
  process.exit(1);
}

console.log('DESIGN SYSTEM CONTRACT PASS');
for (const warning of warnings) console.warn(`WARN: ${warning}`);
