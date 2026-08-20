import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

const productPage = readFileSync('src/pages/sablon/[slug].astro', 'utf8');

test('product schema does not invent reviews or aggregate ratings', () => {
  assert.doesNotMatch(productPage, /['\"]@type['\"]\s*:\s*['\"]Review['\"]/);
  assert.doesNotMatch(productPage, /AggregateRating/);
  assert.doesNotMatch(productPage, /aggregateRating/);
});
