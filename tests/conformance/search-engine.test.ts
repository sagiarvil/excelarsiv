import test from 'node:test';
import assert from 'node:assert/strict';
import {
  searchEngine,
  normalizeTurkish,
  levenshteinDistance,
  extractSearchIntent,
  findDidYouMeanSuggestion,
  UniversalSearchEngine,
  type SearchItem,
} from '../../src/lib/search-engine.ts';

const sampleItems: SearchItem[] = [
  {
    name: '13 Haftalık Nakit Akışı ve Ödeme Planlama Sistemi',
    summary: 'Haftalık nakit giriş-çıkışlarını ve ödeme takvimini planlayan Excel karar modeli.',
    category: 'Nakit Akışı',
    url: '/sablon/13-haftalik-nakit-akisi-ve-odeme-planlama-sistemi',
    keywords: 'nakit akisi, odeme takvimi, likidite, nakit acigi',
  },
  {
    name: 'Akıllı Kasa Defteri ve Nakit Kontrol Sistemi',
    summary: 'Günlük kasa giriş çıkışlarını ve fiili kasa mutabakatını sağlayan tablo.',
    category: 'Nakit Akışı',
    url: '/sablon/akilli-kasa-defteri-ve-nakit-kontrol-sistemi',
    keywords: 'kasa defteri, kasa acigi, gunluk kasa',
  },
  {
    name: 'Amortisman ve Sabit Kıymet Satış Zamanlama Stratejisti',
    summary: 'Sabit kıymet amortisman planı ve vergi tasarruflu satış zamanı simülasyonu.',
    category: 'Muhasebe ve Vergi',
    url: '/sablon/amortisman-ve-sabit-kiymet-satis-zamanlama-stratejisti',
    keywords: 'amortisman, sabit kiymet, yeniden degerleme, vergi',
  },
  {
    name: 'Cari Hesap Tahsilat ve Müşteri Risk Takip Sistemi',
    summary: 'Müşteri vadelerini, gecikmeleri ve tahsilat risk skorlarını ölçen Excel tablosu.',
    category: 'Muhasebe ve Vergi',
    url: '/sablon/cari-hesap-tahsilat-ve-musteri-risk-takip-sistemi',
    keywords: 'cari hesap, tahsilat, vade riski, musteri risk',
  },
  {
    name: 'Proje ve İş Bazında Gerçek Kârlılık Sistemi',
    summary: 'Her işin ve projenin net brüt marjını hesaplayan kârlılık yönetim aracı.',
    category: 'Finansal Analiz',
    url: '/sablon/proje-ve-is-bazinda-gercek-karlilik-sistemi',
    keywords: 'karlilik, proje kari, marj analizi, net kar',
  },
  {
    name: 'Çek Senet ve Vade Risk Sistemi',
    summary: 'Alınan ve verilen çeklerin vade dağılımını ve takas riskini takip eder.',
    category: 'Nakit Akışı',
    url: '/sablon/cek-senet-ve-vade-risk-sistemi',
    keywords: 'cek senet, vadeli cek, senet takibi',
  },
];

test('Search Engine: Turkish Normalization & Diacritic Folding', () => {
  assert.equal(normalizeTurkish('KÂRLILIK'), 'karlilik');
  assert.equal(normalizeTurkish('Nakit Akışı'), 'nakit akisi');
  assert.equal(normalizeTurkish('Çek & Senet'), 'cek senet');
  assert.equal(normalizeTurkish('Ödeme Planı'), 'odeme plani');
});

test('Search Engine: Levenshtein Distance Calculation', () => {
  assert.equal(levenshteinDistance('nakit', 'nakit'), 0);
  assert.equal(levenshteinDistance('amartisman', 'amortisman'), 1);
  assert.equal(levenshteinDistance('karllik', 'karlilik'), 1);
  assert.equal(levenshteinDistance('muasebe', 'muhasebe'), 1);
});

test('Search Engine: Question & Intent Extraction', () => {
  const parsed = extractSearchIntent('nakit akışını nasıl takip ederim');
  assert.ok(parsed.cleanTokens.includes('nakit'));
  assert.ok(parsed.cleanTokens.includes('akisini'));
  assert.ok(!parsed.cleanTokens.includes('nasil'));
});

test('Search Engine: Typo Handling & Fuzzy Matching', () => {
  // Misspelled: "amartisman" -> should match Amortisman item
  const res1 = searchEngine(sampleItems, 'amartisman');
  assert.ok(res1.results.length > 0);
  assert.equal(res1.results[0].name, 'Amortisman ve Sabit Kıymet Satış Zamanlama Stratejisti');

  // Misspelled: "karllık" -> should match Kârlılık item
  const res2 = searchEngine(sampleItems, 'karllık');
  assert.ok(res2.results.length > 0);
  assert.equal(res2.results[0].name, 'Proje ve İş Bazında Gerçek Kârlılık Sistemi');

  // Misspelled / joined: "ceksenet" -> should match Çek Senet
  const res3 = searchEngine(sampleItems, 'ceksenet');
  assert.ok(res3.results.length > 0 || res3.didYouMean !== null || res3.suggestedItems.length > 0);
});

test('Search Engine: Question Query Handling', () => {
  const res = searchEngine(sampleItems, 'nakit akışı nasıl planlanır');
  assert.ok(res.results.length > 0);
  assert.equal(res.results[0].name, '13 Haftalık Nakit Akışı ve Ödeme Planlama Sistemi');
});

test('Search Engine: Did You Mean Suggestions', () => {
  const res = searchEngine(sampleItems, 'nakıt akısı');
  assert.ok(res.results.length > 0);
  
  const resTypo = searchEngine(sampleItems, 'amartısman');
  assert.ok(resTypo.results.length > 0 || resTypo.didYouMean !== null);
});

test('Search Engine: UniversalSearchEngine class', () => {
  const engine = new UniversalSearchEngine(sampleItems);
  const res = engine.search('tahsilat ve cari risk');
  assert.ok(res.results.length > 0);
  assert.equal(res.results[0].name, 'Cari Hesap Tahsilat ve Müşteri Risk Takip Sistemi');
});
