import { describe, it } from 'node:test';
import assert from 'node:assert';
import { productSeo } from '../../src/data/productSeo.ts';

describe('Product Query Uniqueness Contract', () => {
  it('primaryQuery must be unique across all products', () => {
    const queries = new Set();
    for (const [key, entry] of Object.entries(productSeo)) {
      assert.ok(!queries.has(entry.primaryQuery), `Duplicate primaryQuery in ${key}: ${entry.primaryQuery}`);
      queries.add(entry.primaryQuery);
    }
  });
});
