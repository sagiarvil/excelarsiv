// Firebase Hosting deploy gate.
// Any hosting deploy path (CI, local Firebase CLI, future automation) must carry a
// finalized sitemap index. Local production deploys are additionally locked to a
// clean, current main checkout so an old/feature branch cannot overwrite live.
import { spawnSync } from 'node:child_process';
import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { DIST_DIR } from './lib.mjs';
import { parseSitemapIndex } from './finalize-sitemap-index.mjs';

function run(script) {
  // firebase CLI'nin paketli node'u ESM'i require edemediği için predeploy
  // sarmalayıcısı gerçek sistem node'unu PREDEPLOY_NODE olarak geçer.
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

  // Firebase CLI macOS'ta Rosetta/x86_64 Node ile predeploy başlatabildiği için
  // Apple Silicon sistemlerde /usr/bin/git -> xcrun yanlış mimaride yüklenebilir.
  // Kilidi gevşetmek yerine yalnız bu bilinen mimari hatasında native arm64 git
  // çağrısı denenir. Intel macOS veya diğer platformlarda normal git davranışı korunur.
  const macRosettaGitFailure =
    process.platform === 'darwin' &&
    (result.error || result.status !== 0) &&
    /(?:libxcrun|missing compatible architecture|need 'x86_64'|need 'arm64')/i.test(firstError);

  if (macRosettaGitFailure) {
    const nativeArgs = ['-arm64', '/usr/bin/git', ...args];
    result = gitSpawn('/usr/bin/arch', nativeArgs, { inherit });
  }

  if (result.error || result.status !== 0) {
    const detail = inherit ? '' : String(result.stderr || firstError || '').trim();
    throw new Error(`git ${args.join(' ')} başarısız${detail ? `: ${detail}` : ''}`);
  }
  return inherit ? '' : String(result.stdout || '').trim();
}

export function assertLocalProductionReleaseLock() {
  // GitHub Actions staging preview'ları feature/PR ref üzerinde çalışır; onlar CI
  // provenance kapılarıyla korunur. Bu kilit özellikle elle yapılan production
  // `firebase deploy --only hosting:excelarsiv` komutunu fail-closed korur.
  if (process.env.GITHUB_ACTIONS === 'true') return;

  try {
    const branch = git(['branch', '--show-current']);
    if (branch !== 'main') {
      throw new Error(`aktif branch "${branch || '(detached)'}"; production deploy yalnız main üzerinden yapılabilir`);
    }

    // origin/main bilgisini stale bırakmamak için deploy öncesi uzak ref yenilenir.
    git(['fetch', '--quiet', 'origin', 'main']);

    const head = git(['rev-parse', 'HEAD']);
    const originMain = git(['rev-parse', 'origin/main']);
    if (head !== originMain) {
      throw new Error(`HEAD (${head.slice(0, 12)}) origin/main (${originMain.slice(0, 12)}) ile aynı değil`);
    }

    // Tracked dosyalardaki staged/unstaged değişiklikler eski veya doğrulanmamış
    // dist üretimine yol açabilir. Untracked dosyalar build artefaktları nedeniyle
    // burada bilinçli olarak kapsam dışıdır.
    const dirty = git(['status', '--porcelain', '--untracked-files=no']);
    if (dirty) {
      throw new Error('çalışma ağacında commitlenmemiş tracked değişiklik var');
    }

    console.log(`PRODUCTION RELEASE LOCK PASS — main @ ${head.slice(0, 12)} = origin/main, tracked tree clean`);
  } catch (error) {
    console.error('PRODUCTION RELEASE LOCK BLOCKED');
    console.error(error instanceof Error ? error.message : String(error));
    console.error('Çözüm: git checkout main && git pull --ff-only origin main && npm run build');
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
  console.log('HOSTING PREDEPLOY PASS — release lock + finalized sitemap + artifact validation');
}

if (import.meta.url === `file://${process.argv[1]}`) main();
