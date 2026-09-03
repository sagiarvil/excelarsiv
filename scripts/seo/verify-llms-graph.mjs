import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs';
import { join, resolve } from 'node:path';

const ROOT = resolve(process.cwd());
const argDirIndex = process.argv.indexOf('--dir');
const DIR = resolve(ROOT, argDirIndex >= 0 ? process.argv[argDirIndex + 1] : 'dist');
const HOST = 'https://excelarsiv.com/';
const failures = [];
const fail = (code, message) => failures.push(`[${code}] ${message}`);

function text(rel) {
  const file = join(DIR, rel);
  if (!existsSync(file)) {
    fail('LLMS-G0', `dosya yok: ${rel}`);
    return '';
  }
  return readFileSync(file, 'utf8');
}

function walk(dir) {
  if (!existsSync(dir)) return [];
  const out = [];
  for (const name of readdirSync(dir)) {
    const p = join(dir, name);
    if (statSync(p).isDirectory()) out.push(...walk(p));
    else if (name.endsWith('.md')) out.push(p);
  }
  return out;
}

const manifest = text('llms.txt');
if (manifest.trim().length < 300) fail('LLMS-G1', 'llms.txt boş/çok kısa');
if (!manifest.includes(`${HOST}llms-full.txt`)) fail('LLMS-G2', 'llms-full yönlendirmesi yok');

const queue = [...new Set([...manifest.matchAll(/https:\/\/excelarsiv\.com\/(llms\/[A-Za-z0-9_.\/-]+\.md)/g)].map((m) => m[1]))];
const visited = new Set();
while (queue.length) {
  const rel = queue.shift();
  if (visited.has(rel)) continue;
  visited.add(rel);
  const body = text(rel);
  if (!body) continue;
  if (body.trim().length < 120) fail('LLMS-G3', `node çok kısa: ${rel}`);
  for (const m of body.matchAll(/https:\/\/excelarsiv\.com\/(llms\/[A-Za-z0-9_.\/-]+\.md)/g)) {
    if (!visited.has(m[1])) queue.push(m[1]);
  }
}

if (visited.size < 9) fail('LLMS-G4', `çok-katmanlı graph yetersiz: ${visited.size} node`);
for (const required of [
  'llms/entities/excelarsiv.md',
  'llms/categories/finansal-karar-araclari.md',
  'llms/tools/pos-karlilik.md',
]) {
  if (!visited.has(required)) fail('LLMS-G5', `topic ownership node manifest graphına bağlı değil: ${required}`);
}

const llmsDir = join(DIR, 'llms');
for (const file of walk(llmsDir)) {
  const rel = file.slice(DIR.length + 1).replaceAll('\\', '/');
  const body = readFileSync(file, 'utf8');
  if (/kredi onay garantisi verir|garantili kredi onayı/i.test(body)) fail('LLMS-G6', `yasak finansal garanti dili: ${rel}`);
}

const robots = text('robots.txt');
for (const crawler of ['OAI-SearchBot', 'ChatGPT-User', 'Claude-SearchBot', 'Claude-User', 'PerplexityBot']) {
  const block = new RegExp(`User-agent:\\s*${crawler}([\\s\\S]*?)(?=User-agent:|Sitemap:|$)`, 'i').exec(robots)?.[1] ?? '';
  if (!/Allow:\s*\//i.test(block) || /Disallow:\s*\/\s*(?:\r?\n|$)/i.test(block)) fail('ROBOTS-G1', `${crawler} retrieval policy açık değil`);
}
for (const crawler of ['GPTBot', 'ClaudeBot', 'Google-Extended', 'CCBot']) {
  const block = new RegExp(`User-agent:\\s*${crawler}([\\s\\S]*?)(?=User-agent:|Sitemap:|$)`, 'i').exec(robots)?.[1] ?? '';
  if (!/Disallow:\s*\/\s*(?:\r?\n|$)/i.test(block)) fail('ROBOTS-G2', `${crawler} training policy ayrı/kapalı değil`);
}

const sablonlar = text('llms/sablonlar.md');
if (!/owner'ı Dr\. Fin|owner'ı Dr\. Fin/i.test(sablonlar)) fail('OWNERSHIP-G1', 'DRFIN sınırı sablonlar nodeunda tanımlı değil');
if (!/Excel Arşiv.*primer ticari intent/i.test(sablonlar)) fail('OWNERSHIP-G2', 'ExcelArsiv primer intent sınırı tanımlı değil');

if (failures.length) {
  console.error('KIRMIZI — ExcelArsiv LLMS/Search Revenue graph\n' + failures.map((x) => `  ${x}`).join('\n'));
  process.exit(1);
}
console.log(`YEŞİL — ExcelArsiv LLMS graph: ${visited.size} bağlı node; retrieval/training policy ayrık`);
