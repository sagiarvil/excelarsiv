import test from 'node:test';
import assert from 'node:assert/strict';
import { readdirSync } from 'node:fs';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { urunHikayeManifestleri } from '../../src/data/urunGorselManifestleri.ts';
import { validateVisualUniqueness } from '../../src/lib/urun-hikayesi/gorsel-parmak-izi.ts';
import { sahnePromptuOlustur } from '../../src/lib/urun-hikayesi/prompt-olusturucu.ts';

const ROOT = resolve(fileURLToPath(new URL('../../', import.meta.url)));

test('pilot görsel manifestleri benzersiz fingerprint taşır', () => {
  assert.doesNotThrow(() => validateVisualUniqueness(urunHikayeManifestleri));
});

test('pilot set 8 ürün içerir, slug tekrarı yok', () => {
  assert.equal(urunHikayeManifestleri.length, 8);
  const slugSet = new Set(urunHikayeManifestleri.map((m) => m.slug));
  assert.equal(slugSet.size, 8);
});

test('her manifest slug canlı katalogda mevcut', () => {
  const dosyalar = readdirSync(resolve(ROOT, 'src/content/templates'));
  for (const m of urunHikayeManifestleri) {
    assert.ok(dosyalar.includes(`${m.slug}.mdx`), `${m.slug}.mdx eksik`);
  }
});

test('her manifest story ve resultSignal taşır', () => {
  for (const m of urunHikayeManifestleri) {
    assert.ok(m.story.length >= 20, `${m.slug}: story çok kısa`);
    assert.ok(m.resultSignal.length >= 3, `${m.slug}: resultSignal boş`);
    assert.ok(m.ui.type.length > 0, `${m.slug}: ui tipi eksik`);
  }
});

test('bilerek çakışan iki manifest doğrulayıcıyı düşürür', () => {
  const kopya = structuredClone(urunHikayeManifestleri[0]);
  kopya.slug = 'kopya-slug';
  assert.throws(() => validateVisualUniqueness([urunHikayeManifestleri[0], kopya]));
});

test('sahne promptu ürün briefini içerir ve rastgele metni yasaklar', () => {
  const prompt = sahnePromptuOlustur(urunHikayeManifestleri[0]);
  assert.ok(prompt.includes(urunHikayeManifestleri[0].title));
  assert.ok(prompt.includes('do NOT render the title'));
  assert.ok(prompt.includes('do NOT add random unreadable text'));
});
