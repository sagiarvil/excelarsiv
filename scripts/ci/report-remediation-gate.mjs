#!/usr/bin/env node
/**
 * HTML&HTML 2026 raporundaki doğrulanabilir teknik bulgular için kalıcı regression gate.
 * Amaç: raporu "bir kez düzeltmek" değil, aynı sınıf hataların yeniden üretime çıkmasını engellemek.
 */
import fs from 'node:fs';
import path from 'node:path';

const root = process.cwd();
const failures = [];
const pass = (condition, message) => { if (!condition) failures.push(message); };
const read = (p) => fs.readFileSync(path.join(root, p), 'utf8');
const exists = (p) => fs.existsSync(path.join(root, p));

const firebase = JSON.parse(read('firebase.json'));
const hosting = Array.isArray(firebase.hosting) ? firebase.hosting[0] : firebase.hosting;
pass(Boolean(hosting), 'firebase hosting config missing');

const globalHeaders = hosting?.headers?.find((x) => x.source === '/**')?.headers ?? [];
const hsts = globalHeaders.find((x) => x.key === 'Strict-Transport-Security')?.value ?? '';
for (const token of ['max-age=63072000', 'includeSubDomains']) {
  pass(hsts.includes(token), `HSTS missing token: ${token}`);
}
pass(!/(?:^|;\s*)preload(?:;|$)/i.test(hsts), 'HSTS preload must remain disabled until the dedicated approval gate allows it');

const redirects = new Map((hosting?.redirects ?? []).map((x) => [x.source, x]));
const requiredRedirects = new Map([
  ['/about', '/hakkinda'],
  ['/hakkimizda', '/hakkinda'],
  ['/privacy', '/gizlilik-politikasi'],
  ['/privacy-policy', '/gizlilik-politikasi'],
]);
for (const [source, destination] of requiredRedirects) {
  const r = redirects.get(source);
  pass(r?.type === 301 && r?.destination === destination, `missing 301 alias: ${source} -> ${destination}`);
}

const rewrites = new Map((hosting?.rewrites ?? []).map((x) => [x.source, x]));
pass(rewrites.get('/mcp')?.destination === '/mcp.json', '/mcp must resolve to /mcp.json');

const agentHeaderCount = (hosting?.headers ?? []).filter((x) => x.source === '/.well-known/agent-card.json').length;
pass(agentHeaderCount === 1, `agent-card header definition count must be 1, got ${agentHeaderCount}`);

for (const file of [
  'public/llms.txt',
  'public/llms-full.txt',
  'public/mcp.json',
  'public/.well-known/mcp.json',
  'public/.well-known/agent-card.json',
  'src/pages/hakkinda.astro',
  'src/pages/gizlilik-politikasi.astro',
]) {
  pass(exists(file), `required remediation surface missing: ${file}`);
}

for (const file of ['public/mcp.json', 'public/.well-known/mcp.json', 'public/.well-known/agent-card.json']) {
  if (!exists(file)) continue;
  try { JSON.parse(read(file)); } catch { failures.push(`invalid JSON: ${file}`); }
}

if (exists('public/llms.txt')) pass(Buffer.byteLength(read('public/llms.txt')) > 500, 'llms.txt is unexpectedly empty/thin');
if (exists('public/llms-full.txt')) pass(Buffer.byteLength(read('public/llms-full.txt')) > 5000, 'llms-full.txt is unexpectedly empty/thin');

for (const file of ['src/layouts/CommerceLayout.astro', 'src/layouts/WorkbookLayout.astro']) {
  const source = read(file);
  pass(source.includes('rel="canonical"'), `canonical missing in ${file}`);
  pass(source.includes('hreflang="tr"'), `tr hreflang missing in ${file}`);
  pass(source.includes('hreflang="x-default"'), `x-default hreflang missing in ${file}`);
  pass(source.includes('application/ld+json'), `JSON-LD missing in ${file}`);
}

const identityFiles = [
  'src/data/yazar.ts',
  'src/seo/registry.ts',
  'src/seo/schema-builder.ts',
].map(read).join('\n');

const knownInvalidEntityTokens = [
  'Q11190',
  'Q11589432',
  'Doğan Aydın',
  '2230353841',
  'bilgi@excelarsiv.com',
];
for (const token of knownInvalidEntityTokens) {
  pass(!identityFiles.includes(token), `stale/unverified entity identity remains: ${token}`);
}

if (failures.length) {
  console.error(`REPORT REMEDIATION GATE: FAIL (${failures.length})`);
  for (const failure of failures) console.error(` - ${failure}`);
  process.exit(1);
}
console.log('REPORT REMEDIATION GATE: PASS');
