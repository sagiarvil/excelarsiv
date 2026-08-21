import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

test('production hosting predeploy is fail-closed to clean current main for local deploys', () => {
  const source = readFileSync('scripts/seo/hosting-predeploy.mjs', 'utf8');

  assert.match(source, /assertLocalProductionReleaseLock\(\)/);
  assert.match(source, /process\.env\.GITHUB_ACTIONS === ['"]true['"]/);
  assert.match(source, /branch !== ['"]main['"]/);
  assert.match(source, /git\(\[['"]fetch['"], ['"]--quiet['"], ['"]origin['"], ['"]main['"]\]\)/);
  assert.match(source, /git\(\[['"]rev-parse['"], ['"]HEAD['"]\]\)/);
  assert.match(source, /git\(\[['"]rev-parse['"], ['"]origin\/main['"]\]\)/);
  assert.match(source, /head !== originMain/);
  assert.match(source, /git\(\[['"]status['"], ['"]--porcelain['"], ['"]--untracked-files=no['"]\]\)/);
  assert.match(source, /PRODUCTION RELEASE LOCK BLOCKED/);

  const mainCall = source.indexOf('assertLocalProductionReleaseLock();');
  const artifactCheck = source.indexOf("existsSync(join(DIST_DIR, 'seo-artifacts.json'))");
  assert.ok(mainCall >= 0 && artifactCheck >= 0 && mainCall < artifactCheck, 'release lock must run before deploy artifact validation');
});

test('macOS Rosetta git failure retries native arm64 without weakening release checks', () => {
  const source = readFileSync('scripts/seo/hosting-predeploy.mjs', 'utf8');
  assert.match(source, /process\.platform === ['"]darwin['"]/);
  assert.match(source, /libxcrun|missing compatible architecture/);
  assert.match(source, /nativeArgs = \[['"]-arm64['"], ['"]\/usr\/bin\/git['"], \.\.\.args\]/);
  assert.match(source, /gitSpawn\(['"]\/usr\/bin\/arch['"], nativeArgs/);
  assert.match(source, /branch !== ['"]main['"]/);
  assert.match(source, /head !== originMain/);
});
