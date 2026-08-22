#!/usr/bin/env node
import { createHash } from 'node:crypto';
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import { join } from 'node:path';

function git(args) {
  const r = spawnSync('git', args, { encoding: 'utf8' });
  if (r.status !== 0) throw new Error(`git ${args.join(' ')} failed: ${String(r.stderr || '').trim()}`);
  return String(r.stdout || '').trim();
}

function sha256(path) {
  return createHash('sha256').update(readFileSync(path)).digest('hex');
}

const head = git(['rev-parse', 'HEAD']);
const routes = {
  '/': join('dist', 'index.html'),
  '/sablonlar': join('dist', 'sablonlar', 'index.html'),
  '/ozel-excel-sistemleri': join('dist', 'ozel-excel-sistemleri', 'index.html'),
};

for (const [route, path] of Object.entries(routes)) {
  if (!existsSync(path)) throw new Error(`BUILD PROVENANCE: ${route} artifact missing at ${path}`);
}

const manifest = {
  schema: 1,
  gitSha: head,
  createdAt: new Date().toISOString(),
  routes: Object.fromEntries(Object.entries(routes).map(([route, path]) => [route, { path, sha256: sha256(path) }])),
};

writeFileSync(join('dist', '.build-provenance.json'), `${JSON.stringify(manifest, null, 2)}\n`, 'utf8');
console.log(`BUILD PROVENANCE WRITTEN — ${head.slice(0, 12)}; ${Object.keys(routes).join(', ')}`);
