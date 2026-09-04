#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

const root = process.cwd();
const contractPath = path.join(root, 'data/seo/authority/revenue-authority.json');
const registryPath = path.join(root, 'data/seo/registry/excelarsiv_seo_registry.json');
const outDir = path.join(root, 'public', '.well-known');
const read = p => JSON.parse(fs.readFileSync(p, 'utf8'));
const fail = m => { console.error(`AUTHORITY FAIL: ${m}`); process.exitCode = 1; };
const warn = m => console.warn(`AUTHORITY WARN: ${m}`);
const norm = r => r === '/' ? '/' : String(r).replace(/\/$/, '');

if (!fs.existsSync(contractPath) || !fs.existsSync(registryPath)) {
  console.error('AUTHORITY FAIL: contract/registry missing');
  process.exit(1);
}

const contract = read(contractPath);
const registry = read(registryPath);
const routes = new Set((registry.records || []).filter(r => r.status === 'live').map(r => norm(r.route)));
const routeParityHard = registry?.meta?.partial !== true && registry?.meta?.coldStart !== true;
const assertRoute = (route, label) => {
  if (routes.has(norm(route))) return;
  if (routeParityHard) fail(`${label} not in live registry: ${route}`);
  else warn(`${label} absent from partial/cold-start registry: ${route}`);
};
const owners = new Map();
const graph = { site: contract.site, version: contract.version, generatedAt: new Date().toISOString(), registryConfidence: registry?.meta?.confidence ?? null, clusters: [] };

for (const c of contract.clusters || []) {
  const owner = norm(c.commercialOwner);
  assertRoute(owner, `${c.id} commercial owner`);
  assertRoute(c.hub, `${c.id} hub`);
  if ((c.supportRoutes || []).length < contract.policy.minSupportRoutes) fail(`${c.id}: insufficient support routes`);
  for (const r of c.supportRoutes || []) assertRoute(r, `${c.id} support route`);
  if ((c.contextualAnchors || []).length < contract.policy.minContextualAnchors) fail(`${c.id}: contextual anchor coverage too low`);
  if (contract.policy.requireDirectAnswer && !c.answer?.question) fail(`${c.id}: direct answer contract missing`);
  if (contract.policy.requireEvidence && !(c.evidence || []).length) fail(`${c.id}: evidence contract missing`);
  if (!c.conversion) fail(`${c.id}: revenue conversion event missing`);

  const ageDays = (Date.now() - Date.parse(c.lastMaterialReview)) / 86400000;
  if (!Number.isFinite(ageDays) || ageDays < 0 || ageDays > contract.policy.maxReviewAgeDays) fail(`${c.id}: stale/invalid material review date`);

  for (const q0 of c.ownedQueries || []) {
    const q = q0.trim().toLocaleLowerCase('tr-TR');
    if (owners.has(q) && owners.get(q) !== owner) fail(`query collision: "${q0}" -> ${owners.get(q)} AND ${owner}`);
    owners.set(q, owner);
  }

  graph.clusters.push({ id: c.id, owner, hub: norm(c.hub), supports: c.supportRoutes.map(norm), queries: c.ownedQueries, anchors: c.contextualAnchors, directAnswerQuestion: c.answer.question, evidence: c.evidence, conversion: c.conversion, lastMaterialReview: c.lastMaterialReview });
}

fs.mkdirSync(outDir, { recursive: true });
fs.writeFileSync(path.join(outDir, 'authority-graph.json'), JSON.stringify(graph, null, 2) + '\n');
fs.mkdirSync(path.join(root, 'data', 'seo', 'authority', 'generated'), { recursive: true });
fs.writeFileSync(path.join(root, 'data', 'seo', 'authority', 'generated', 'contextual-link-plan.json'), JSON.stringify({ generatedAt: graph.generatedAt, rules: contract.clusters.map(c => ({ owner: norm(c.commercialOwner), sourceRoutes: [norm(c.hub), ...c.supportRoutes.map(norm)], anchors: c.contextualAnchors, insertionPolicy: 'first-meaningful-occurrence-only; never alter headings; max-one-link-per-anchor-per-page' })) }, null, 2) + '\n');

if (process.exitCode) process.exit(process.exitCode);
console.log(`AUTHORITY PASS: ${contract.clusters.length} revenue clusters, ${owners.size} owned queries; machine graph + contextual-link plan generated.`);
