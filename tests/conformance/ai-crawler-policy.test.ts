import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

const robots = readFileSync('public/robots.txt', 'utf8');
const liveContract = readFileSync('scripts/seo/live-contract.mjs', 'utf8');

test('OAI Search crawler is explicitly allowed', () => {
  assert.match(robots, /User-agent:\s*OAI-SearchBot[\s\S]*?Allow:\s*\//i);
});

test('live contract verifies OAI Search crawler policy', () => {
  assert.match(liveContract, /OAI-SearchBot/);
  assert.match(liveContract, /Allow:\\s\*\\\//);
});
