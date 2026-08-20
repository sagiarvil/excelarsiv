import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

const source = readFileSync('src/pages/neden-excel-arsiv.astro', 'utf8');

test('comparison authority is neutral and names all three approaches', () => {
  assert.match(source, /ExcelArşiv, ChatGPT ve boş Excel/);
  assert.match(source, /ExcelArşiv ne zaman daha uygundur/);
  assert.match(source, /ChatGPT ne zaman daha uygundur/);
  assert.match(source, /Boş Excel ne zaman daha uygundur/);
  assert.match(source, /her durumda daha iyi.*değildir/i);
  assert.doesNotMatch(source, /ChatGPT hata yapar|ChatGPT yanlıştır/i);
});
