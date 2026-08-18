import test from 'node:test';
import assert from 'node:assert/strict';
import { secUrunler } from '../../src/scripts/urun-bulucu.ts';

test('ürün bulucu en fazla 3 sonuç döner ve ana ürün problemle hizalanır', () => {
  const products = [
    { slug: '13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi', name: 'Nakit', summary: 'plan', category: 'nakit-akisi', categoryName: 'Nakit Akışı', priceTL: 999, url: '/sablon/a', demoUrl: '/demo/a' },
    { slug: 'akilli-kasa-defteri-ve-nakit-kontrol-sistemi', name: 'Kasa', summary: 'kasa', category: 'nakit-akisi', categoryName: 'Nakit Akışı', priceTL: 499, url: '/sablon/b', demoUrl: '/demo/b' },
    { slug: 'cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi', name: 'Cari', summary: 'cari', category: 'muhasebe-ve-vergi', categoryName: 'Muhasebe', priceTL: 799, url: '/sablon/c', demoUrl: '/demo/c' },
  ];
  const selected = secUrunler(products, { alan: 'nakit-akisi', tip: 'ticaret', hacim: 'orta', donem: 'hafta', problem: 'nakit' });
  assert.ok(selected.length >= 1 && selected.length <= 3);
  assert.equal(selected[0]?.slug, '13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi');
});
