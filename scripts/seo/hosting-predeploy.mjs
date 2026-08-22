// Firebase Hosting deploy gate.
// Any hosting deploy path (CI, local Firebase CLI, future automation) must carry a
// finalized sitemap index AND a dist provenance manifest tied to the exact git SHA.
// Local production deploys are additionally locked to a clean, current main checkout.
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { DIST_DIR } from './lib.mjs';
import { parseSitemapIndex } from './finalize-sitemap-index.mjs';

function run(script) {
  const node = process.env.PREDEPLOY_NODE || process.execPath;
  const result = spawnSync(node, [script], { cwd: process.cwd(), env: process.env, stdio: 'inherit' });
  if (result.error) throw result.error;
  if (result.status !== 0) process.exit(result.status ?? 1);
}

function gitSpawn(command, args, { inherit = false } = {}) {
  return spawnSync(command, args, {
    cwd: process.cwd(),
    env: process.env,
    encoding: inherit ? undefined : 'utf8',
    stdio: inherit ? 'inherit' : ['ignore', 'pipe', 'pipe'],
  });
}

function git(args, { inherit = false } = {}) {
  let result = gitSpawn('git', args, { inherit });
  const firstError = inherit ? '' : String(result.stderr || '').trim();
  const macRosettaGitFailure =
    process.platform === 'darwin' &&
    (result.error || result.status !== 0) &&
    /(?:libxcrun|missing compatible architecture|need 'x86_64'|need 'arm64')/i.test(firstError);

  if (macRosettaGitFailure) {
    result = gitSpawn('/usr/bin/arch', ['-arm64', '/usr/bin/git', ...args], { inherit });
  }

  if (result.error || result.status !== 0) {
    const detail = inherit ? '' : String(result.stderr || firstError || '').trim();
    throw new Error(`git ${args.join(' ')} başarısız${detail ? `: ${detail}` : ''}`);
  }
  return inherit ? '' : String(result.stdout || '').trim();
}

function sha256(path) {
  return createHash('sha256').update(readFileSync(path)).digest('hex');
}

export function assertLocalProductionReleaseLock() {
  if (process.env.GITHUB_ACTIONS === 'true') return;

  try {
    const branch = git(['branch', '--show-current']);
    if (branch !== 'main') throw new Error(`aktif branch "${branch || '(detached)'}"; production deploy yalnız main üzerinden yapılabilir`);

    git(['fetch', '--quiet', 'origin', 'main']);
    const head = git(['rev-parse', 'HEAD']);
    const originMain = git(['rev-parse', 'origin/main']);
    if (head !== originMain) throw new Error(`HEAD (${head.slice(0, 12)}) origin/main (${originMain.slice(0, 12)}) ile aynı değil`);

    const dirty = git(['status', '--porcelain', '--untracked-files=no']);
    if (dirty) throw new Error('çalışma ağacında commitlenmemiş tracked değişiklik var');

    console.log(`PRODUCTION RELEASE LOCK PASS — main @ ${head.slice(0, 12)} = origin/main, tracked tree clean`);
  } catch (error) {
    console.error('PRODUCTION RELEASE LOCK BLOCKED');
    console.error(error instanceof Error ? error.message : String(error));
    console.error('Çözüm: git checkout main && git pull --ff-only origin main && npm run build');
    process.exit(1);
  }
}

export function assertDistProvenance({ dist = DIST_DIR } = {}) {
  const manifestPath = join(dist, '.build-provenance.json');
  if (!existsSync(manifestPath)) {
    console.error('DIST PROVENANCE BLOCKED — dist/.build-provenance.json yok. Bu dist hangi committen üretildiği kanıtlanamıyor.');
    process.exit(1);
  }

  try {
    const manifest = JSON.parse(readFileSync(manifestPath, 'utf8'));
    const head = git(['rev-parse', 'HEAD']);
    if (manifest.gitSha !== head) {
      throw new Error(`dist gitSha ${String(manifest.gitSha).slice(0, 12)} != HEAD ${head.slice(0, 12)}`);
    }

    const requiredRoutes = ['/', '/sablonlar', '/ozel-excel-sistemleri'];
    for (const route of requiredRoutes) {
      const entry = manifest.routes?.[route];
      if (!entry?.path || !entry?.sha256) throw new Error(`manifest route eksik: ${route}`);
      if (!existsSync(entry.path)) throw new Error(`dist artefakt eksik: ${entry.path}`);
      const actual = sha256(entry.path);
      if (actual !== entry.sha256) throw new Error(`${route} dist hash değişmiş/stale: ${actual.slice(0, 12)} != ${entry.sha256.slice(0, 12)}`);
    }

    console.log(`DIST PROVENANCE PASS — dist exactly matches HEAD ${head.slice(0, 12)} and protected route hashes`);
  } catch (error) {
    console.error('DIST PROVENANCE BLOCKED');
    console.error(error instanceof Error ? error.message : String(error));
    console.error('Çözüm: npm run build komutunu mevcut HEAD üzerinde yeniden çalıştırın; eski dist ile deploy yasaktır.');
    process.exit(1);
  }
}

export function finalizedBuildMatches({ dist = DIST_DIR } = {}) {
  const manifestPath = join(dist, 'seo-artifacts.json');
  const reportPath = join(dist, 'seo-finalize-report.json');
  const indexPath = join(dist, 'sitemap.xml');
  if (![manifestPath, reportPath, indexPath].every(existsSync)) return false;
  try {
    const manifest = JSON.parse(readFileSync(manifestPath, 'utf8'));
    const report = JSON.parse(readFileSync(reportPath, 'utf8'));
    if (report.status !== 'PASS' || !Array.isArray(report.children) || !Array.isArray(manifest.children)) return false;
    const generated = new Map(manifest.children.map((child) => [child.file, child.sha256]));
    if (report.children.length !== generated.size) return false;
    for (const child of report.children) {
      if (!child.file || generated.get(child.file) !== child.generated_sha256 || !child.lastmod) return false;
    }
    const indexEntries = parseSitemapIndex(readFileSync(indexPath, 'utf8'));
    if (indexEntries.length !== generated.size || indexEntries.some((entry) => !entry.lastmod)) return false;
    return true;
  } catch {
    return false;
  }
}

async function main() {
  assertLocalProductionReleaseLock();
  assertDistProvenance();

  if (!existsSync(join(DIST_DIR, 'seo-artifacts.json'))) {
    console.error('HOSTING PREDEPLOY KALDI: dist/seo-artifacts.json yok. Önce npm run build çalıştırılmalı.');
    process.exit(1);
  }
  if (finalizedBuildMatches()) {
    console.log('HOSTING PREDEPLOY: mevcut dist aynı child hashleriyle zaten finalized; yeniden baseline zamanı üretilmedi.');
  } else {
    console.log('HOSTING PREDEPLOY: finalized build kanıtı yok/stale; production baseline finalizer çalıştırılıyor.');
    run('scripts/seo/finalize-sitemap-index.mjs');
  }
  run('scripts/seo/validate-artifacts.mjs');
  console.log('HOSTING PREDEPLOY PASS — release lock + dist provenance + finalized sitemap + artifact validation');
}

if (import.meta.url === `file://${process.argv[1]}`) main();
