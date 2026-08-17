import type { UrunGorselManifest } from './turler.ts';

function ortakFingerprintSayisi(a: UrunGorselManifest, b: UrunGorselManifest): number {
  let eslesme = 0;
  if (a.fingerprint.layout === b.fingerprint.layout) eslesme += 1;
  if (a.fingerprint.heroObject === b.fingerprint.heroObject) eslesme += 1;
  if (a.fingerprint.uiModule === b.fingerprint.uiModule) eslesme += 1;
  if (a.fingerprint.perspective === b.fingerprint.perspective) eslesme += 1;
  if (a.fingerprint.accent === b.fingerprint.accent) eslesme += 1;
  return eslesme;
}

export function validateVisualUniqueness(ogeler: UrunGorselManifest[]): boolean {
  const catismalar: string[] = [];
  for (let i = 0; i < ogeler.length; i++) {
    for (let j = i + 1; j < ogeler.length; j++) {
      const ortak = ortakFingerprintSayisi(ogeler[i], ogeler[j]);
      if (ortak >= 3) {
        catismalar.push(`${ogeler[i].slug} <-> ${ogeler[j].slug} | ortak fingerprint=${ortak}`);
      }
    }
  }
  if (catismalar.length > 0) {
    throw new Error(`Visual fingerprint çakışması bulundu:\n${catismalar.join('\n')}`);
  }
  return true;
}
