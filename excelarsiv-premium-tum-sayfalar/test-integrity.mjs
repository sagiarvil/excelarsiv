#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const siteDir = path.join(__dirname, 'site');
const assetsDir = path.join(siteDir, 'assets');

console.log('--- EXCELARŞİV PREMİUM TEST & DOĞRULAMA HATTI ---');

// 1. Assets Check
if (!fs.existsSync(assetsDir)) {
  console.error('HATA: assets klasörü bulunamadı!');
  process.exit(1);
}
const assetsList = fs.readdirSync(assetsDir);
console.log(`[PASS] Assets Varlığı: ${assetsList.length} varlık mevcut.`);

// 2. HTML Files Walk
function walk(dir) {
  let files = [];
  for (const f of fs.readdirSync(dir)) {
    const full = path.join(dir, f);
    if (fs.statSync(full).isDirectory()) {
      files = files.concat(walk(full));
    } else if (f.endsWith('.html')) {
      files.push(full);
    }
  }
  return files;
}

const htmlFiles = walk(siteDir);
console.log(`[PASS] Sayfa Sayısı: ${htmlFiles.length} HTML sayfası bulundu.`);

// 3. Asset Integrity Verification
let missingAssets = [];
let checkedAssets = 0;

for (const file of htmlFiles) {
  const content = fs.readFileSync(file, 'utf8');
  const relDir = path.dirname(file);
  const regex = /(?:src|href)="([^"#?]+(?:\.css|\.js|\.webp|\.svg|\.png|\.jpg|\.ico))"/g;
  let match;
  while ((match = regex.exec(content)) !== null) {
    const ref = match[1];
    if (ref.startsWith('http://') || ref.startsWith('https://') || ref.startsWith('data:')) continue;
    checkedAssets++;
    const target = path.resolve(relDir, ref);
    if (!fs.existsSync(target)) {
      missingAssets.push(`${ref} (${path.relative(siteDir, file)})`);
    }
  }
}

if (missingAssets.length > 0) {
  console.error(`[FAIL] Eksik Varlıklar (${missingAssets.length}):`, missingAssets);
  process.exit(1);
}
console.log(`[PASS] Varlık Referans Doğrulaması: ${checkedAssets} referansın tamamı eksiksiz.`);

// 4. Internal Link Integrity Verification
let missingLinks = [];
let checkedLinks = 0;

for (const file of htmlFiles) {
  const content = fs.readFileSync(file, 'utf8');
  const relDir = path.dirname(file);
  const regex = /href="([^"#?]+(?:\/|\.html)?)"/g;
  let match;
  while ((match = regex.exec(content)) !== null) {
    const ref = match[1];
    if (ref.startsWith('http://') || ref.startsWith('https://') || ref.startsWith('mailto:') || ref.startsWith('tel:') || ref.startsWith('javascript:') || ref.startsWith('data:')) continue;
    if (ref.endsWith('.css') || ref.endsWith('.js') || ref.endsWith('.png') || ref.endsWith('.svg') || ref.endsWith('.ico') || ref.endsWith('.webp')) continue;
    checkedLinks++;
    let target = path.resolve(relDir, ref);
    if (fs.existsSync(target) && fs.statSync(target).isDirectory()) {
      target = path.join(target, 'index.html');
    }
    if (!fs.existsSync(target)) {
      missingLinks.push(`${ref} in ${path.relative(siteDir, file)}`);
    }
  }
}

if (missingLinks.length > 0) {
  console.error(`[FAIL] Kırık Linkler (${missingLinks.length}):`, missingLinks);
  process.exit(1);
}
console.log(`[PASS] İç Bağlantı Bütünlüğü: ${checkedLinks} dahili linkin tamamı hatasız ve geçerli.`);

// 5. tek-dosya.html Verification
const tekDosya = path.join(siteDir, 'tek-dosya.html');
if (!fs.existsSync(tekDosya) || fs.statSync(tekDosya).size < 1000000) {
  console.error('[FAIL] tek-dosya.html eksik veya geçersiz boyutta!');
  process.exit(1);
}
console.log(`[PASS] Tek Dosya Paketi (tek-dosya.html): ${(fs.statSync(tekDosya).size / 1024 / 1024).toFixed(2)} MB sağlam.`);

console.log('----------------------------------------------------');
console.log('SONUÇ: %100 KALİTE KAPISI BAŞARILI (0 HATA / 0 KIRIK LİNK)');
