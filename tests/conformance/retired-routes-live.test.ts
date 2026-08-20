import { describe, it } from 'node:test';
import assert from 'node:assert';

describe('Retired Routes Live Contract', () => {
  const RETIRED_ROUTES = ['/excel-araclari', '/paketler'];
  const BASE_URL = process.env.SEO_BASE_URL || 'https://excelarsiv.com';

  for (const route of RETIRED_ROUTES) {
    it(`Route ${route} should not return 200 OK`, async () => {
      const response = await fetch(`${BASE_URL}${route}`, { redirect: 'manual' });
      assert.ok(response.status === 301 || response.status === 410, `Route ${route} returned ${response.status}, which is forbidden (must be 301 or 410).`);
    });
  }
});
