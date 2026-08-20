import test from 'node:test';
import assert from 'node:assert/strict';
import { productSeo } from '../../src/data/productSeo.ts';
import { productAnswerSeo } from '../../src/data/productAnswerSeo.ts';

test('every product SEO entry has a non-placeholder answer graph entry', () => {
  assert.deepEqual(Object.keys(productAnswerSeo).sort(), Object.keys(productSeo).sort());
  for (const [slug, entry] of Object.entries(productAnswerSeo)) {
    assert.ok(entry.answerQuestion.length >= 20, `${slug}: short answerQuestion`);
    assert.ok(entry.answerSummary.length >= 80, `${slug}: short answerSummary`);
    assert.equal(entry.answerSummary, productSeo[slug].description, `${slug}: answer summary drifted from visible SEO SSOT`);
    assert.doesNotMatch(entry.answerSummary, /placeholder|lorem ipsum|örnek metin/i, `${slug}: placeholder answer copy`);
  }
});

test('primary queries are unique', () => {
  const queries = Object.values(productAnswerSeo).map((entry) => entry.primaryQuery.trim().toLocaleLowerCase('tr-TR'));
  assert.equal(new Set(queries).size, queries.length);
});
