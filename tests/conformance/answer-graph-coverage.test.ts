import { describe, it } from 'node:test';
import assert from 'node:assert';
import { productSeo } from '../../src/data/productSeo.ts';

describe('Answer Graph Coverage Contract', () => {
  it('every productSeo entry must have answerQuestion and answerSummary', () => {
    for (const [key, entry] of Object.entries(productSeo)) {
      assert.ok(entry.answerQuestion, `Missing answerQuestion in ${key}`);
      assert.ok(entry.answerSummary, `Missing answerSummary in ${key}`);
      assert.ok(entry.answerSummary.length >= 160 && entry.answerSummary.length <= 320, `answerSummary length invalid in ${key}`);
    }
  });
});
