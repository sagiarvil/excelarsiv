import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

const liveContract = readFileSync('scripts/seo/live-contract.mjs', 'utf8');

test('live contract validates retired routes as exact 301 redirects without following them', () => {
  assert.match(liveContract, /\/excel-araclari/);
  assert.match(liveContract, /\/paketler/);
  assert.match(liveContract, /redirect:\s*'manual'/);
  assert.match(liveContract, /retired\.status === 301/);
  assert.match(liveContract, /expectedLocation/);
  assert.match(liveContract, /actualLocation === expectedLocation/);
});
