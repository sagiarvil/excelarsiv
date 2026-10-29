import { test, expect } from '@playwright/test';

test.describe('Design Engine & Visual Integrity Gate', () => {
  test('HomePage has no horizontal scroll, no console errors, and renders pristine UI', async ({ page }) => {
    const consoleErrors: string[] = [];
    page.on('console', (msg) => {
      if (msg.type() === 'error') {
        consoleErrors.push(msg.text());
      }
    });

    const response = await page.goto('/', { waitUntil: 'domcontentloaded' });
    expect(response?.status()).toBeLessThan(400);

    // Assert no severe horizontal overflow
    const hasHorizontalScroll = await page.evaluate(() => {
      return document.documentElement.scrollWidth > document.documentElement.clientWidth;
    });
    expect(hasHorizontalScroll).toBe(false);

    // Save proof screenshot
    await page.screenshot({ path: '.screenshots/design-engine-verify.png', fullPage: false });

    // Assert zero critical runtime console exceptions
    const criticalErrors = consoleErrors.filter(
      (err) => !err.includes('favicon') && !err.includes('404')
    );
    expect(criticalErrors.length).toBe(0);
  });
});
