import test from 'node:test';
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = resolve(fileURLToPath(new URL('../../../', import.meta.url)));

test('INV-2 ledger ve firebase 301 senkronu geçer', () => {
  const result = spawnSync(process.execPath, ['--experimental-strip-types', resolve(ROOT, 'scripts/seo/redirect-validate.ts'), '--site', 'excelarsiv'], {
    cwd: ROOT,
    encoding: 'utf8',
  });
  assert.equal(result.status, 0, result.stderr || result.stdout);
});
