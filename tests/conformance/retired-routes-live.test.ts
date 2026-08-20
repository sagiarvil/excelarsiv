import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

const liveContract = readFileSync('scripts/seo/live-contract.mjs', 'utf8');

test('live contract rejects 200 responses for retired routes', () => {
  assert.match(liveContract, /\/excel-araclari/);
  assert.match(liveContract, /\/paketler/);
  assert.match(liveContract, /retired\.status !== 200/);
});
