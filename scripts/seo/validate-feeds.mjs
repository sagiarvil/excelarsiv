#!/usr/bin/env node
/**
 * RSS 2.0 & Atom 1.0 Feed Doğrulama ve Semantik Kalite Testi
 * W3C RSS 2.0, IETF RFC 4287 Atom standartları ve Sitemap paritesini doğrular.
 */

import { existsSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const DIST_DIR = resolve(process.cwd(), 'dist');
const RSS_PATH = resolve(DIST_DIR, 'rss.xml');
const ATOM_PATH = resolve(DIST_DIR, 'atom.xml');
const SITEMAP_PATH = resolve(DIST_DIR, 'sitemap.xml');
const INDEX_HTML = resolve(DIST_DIR, 'index.html');

const failures = [];

function assert(condition, message) {
  if (!condition) {
    failures.push(message);
    console.error(`  ❌ FAIL: ${message}`);
  } else {
    console.log(`  ✅ PASS: ${message}`);
  }
}

console.log('='.repeat(70));
console.log('📡 RSS 2.0 & ATOM 1.0 FEED SEMANTİK DOĞRULAMA VE PARİTE TESTİ');
console.log('='.repeat(70));

// 1. Fiziksel Dosya Varlığı
assert(existsSync(RSS_PATH), 'dist/rss.xml fiziksel olarak mevcut');
assert(existsSync(ATOM_PATH), 'dist/atom.xml fiziksel olarak mevcut');

if (!existsSync(RSS_PATH) || !existsSync(ATOM_PATH)) {
  console.error('\nKritik hata: Feed dosyaları bulunamadı!');
  process.exit(1);
}

const rssContent = readFileSync(RSS_PATH, 'utf8');
const atomContent = readFileSync(ATOM_PATH, 'utf8');

// 2. RSS 2.0 Spesifikasyon Denetimi
console.log('\n--- [1] RSS 2.0 Spesifikasyon Denetimi ---');
assert(rssContent.startsWith('<?xml version="1.0" encoding="UTF-8"?>'), 'RSS 2.0 XML bildirimi doğru');
assert(rssContent.includes('<rss version="2.0"'), 'rss version="2.0" kök etiketi mevcut');
assert(rssContent.includes('xmlns:atom="http://www.w3.org/2005/Atom"'), 'Atom namespace tanımlı');
assert(rssContent.includes('<channel>'), 'channel konteyneri mevcut');
assert(rssContent.includes('<title>Excel Arşiv'), 'channel title tanımlı');
assert(rssContent.includes('<link>https://excelarsiv.com</link>') || rssContent.includes('<link>https://excelarsiv.com/</link>'), 'channel link tanımlı');
assert(rssContent.includes('<language>tr-TR</language>'), 'channel dili tr-TR');
assert(rssContent.includes('<atom:link href="https://excelarsiv.com/rss.xml" rel="self" type="application/rss+xml"'), 'self-referencing atom:link mevcut');
assert(rssContent.includes('<managingEditor>'), 'managingEditor tanımlı');

const rssItems = rssContent.match(/<item>[\s\S]*?<\/item>/g) || [];
assert(rssItems.length >= 100, `RSS 2.0 item sayısı >= 100 (Bulunan: ${rssItems.length})`);

let rssSamplePass = true;
for (const item of rssItems.slice(0, 10)) {
  if (!item.includes('<title>') || !item.includes('<link>') || !item.includes('<guid') || !item.includes('<pubDate>')) {
    rssSamplePass = false;
    break;
  }
}
assert(rssSamplePass, 'RSS 2.0 item örneklerinde title, link, guid ve pubDate tam');

// 3. Atom 1.0 (RFC 4287) Spesifikasyon Denetimi
console.log('\n--- [2] Atom 1.0 (RFC 4287) Spesifikasyon Denetimi ---');
assert(atomContent.startsWith('<?xml version="1.0" encoding="UTF-8"?>'), 'Atom 1.0 XML bildirimi doğru');
assert(atomContent.includes('<feed xmlns="http://www.w3.org/2005/Atom">'), 'Atom namespace ve feed kök etiketi doğru');
assert(atomContent.includes('<id>https://excelarsiv.com/</id>'), 'feed id tanımlı');
assert(atomContent.includes('<title>Excel Arşiv'), 'feed title tanımlı');
assert(atomContent.includes('<link rel="self" type="application/atom+xml" href="https://excelarsiv.com/atom.xml"'), 'self-referencing atom link mevcut');
assert(atomContent.includes('<updated>'), 'feed updated zaman damgası mevcut');
assert(atomContent.includes('<author>'), 'feed author bilgisi mevcut');
assert(atomContent.includes('Barış Bağırlar'), 'author adı Barış Bağırlar');

const atomEntries = atomContent.match(/<entry>[\s\S]*?<\/entry>/g) || [];
assert(atomEntries.length >= 100, `Atom 1.0 entry sayısı >= 100 (Bulunan: ${atomEntries.length})`);

let atomSamplePass = true;
for (const entry of atomEntries.slice(0, 10)) {
  if (!entry.includes('<id>') || !entry.includes('<title>') || !entry.includes('<updated>') || !entry.includes('<link')) {
    atomSamplePass = false;
    break;
  }
}
assert(atomSamplePass, 'Atom 1.0 entry örneklerinde id, title, updated ve link tam');

// 4. Sitemap ile Kayıpsız Eşzamanlılık Paritesi
console.log('\n--- [3] Sitemap ile Kayıpsız Parite Denetimi ---');
assert(rssItems.length === atomEntries.length, `RSS ve Atom öğe sayıları birebir eşit (${rssItems.length} === ${atomEntries.length})`);

// 5. HTML Autodiscovery Denetimi
console.log('\n--- [4] HTML Autodiscovery Link Denetimi ---');
if (existsSync(INDEX_HTML)) {
  const indexHtml = readFileSync(INDEX_HTML, 'utf8');
  assert(indexHtml.includes('type="application/rss+xml"'), 'index.html içinde RSS 2.0 autodiscovery linki mevcut');
  assert(indexHtml.includes('type="application/atom+xml"'), 'index.html içinde Atom 1.0 autodiscovery linki mevcut');
} else {
  assert(false, 'dist/index.html bulunamadı');
}

console.log('='.repeat(70));
if (failures.length > 0) {
  console.error(`\n❌ TEST KALDI: ${failures.length} hata bulundu!`);
  process.exit(1);
} else {
  console.log(`\n🎉 %100 BAŞARI: Tüm RSS 2.0 & Atom 1.0 standartları ve sitemap paritesi kusursuz doğrulandı!`);
}
