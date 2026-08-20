import { describe, it } from 'node:test';
import assert from 'node:assert';
import fs from 'node:fs';

describe('Comparison Page Contract', () => {
  it('should have the unbiased comparison page', () => {
    const content = fs.readFileSync('src/pages/karsilastir/excelarsiv-vs-chatgpt-vs-bos-excel.astro', 'utf8');
    assert.ok(content.includes('ExcelArşiv'), 'Missing ExcelArşiv');
    assert.ok(content.includes('ChatGPT'), 'Missing ChatGPT');
    assert.ok(content.includes('Boş Excel'), 'Missing Boş Excel');
    assert.ok(content.includes('<tr>'), 'Missing SSR HTML table');
    assert.ok(content.includes('ChatGPT keşif'), 'Missing ChatGPT strength note');
    assert.ok(!content.includes('AggregateRating') && !content.includes('Review'), 'Must not contain fake Review/AggregateRating schema');
  });
});
