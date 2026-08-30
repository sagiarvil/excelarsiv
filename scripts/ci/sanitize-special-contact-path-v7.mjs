#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

const distFile = path.resolve('dist/ozel-excel-sistemleri/index.html');
if (!fs.existsSync(distFile)) throw new Error('SPECIAL CONTACT PATH V7: dist route missing');

let html = fs.readFileSync(distFile, 'utf8');

const directWhatsApp = /href="https:\/\/wa\.me\/[^\"]+"(?:\s+target="_blank")?(?:\s+rel="noopener noreferrer")?/gu;
const matches = html.match(directWhatsApp) || [];
if (!matches.length) throw new Error('SPECIAL CONTACT PATH V7: expected direct WhatsApp CTA not found');

html = html.replace(directWhatsApp, 'href="/iletisim"');
html = html.replaceAll("WhatsApp'tan anlat", 'İhtiyacımı anlat');
html = html.replaceAll("WhatsApp'tan Sürecinizi Anlatın", 'Sürecinizi Anlatın');
html = html.replaceAll("WhatsApp'tan dosyamı anlat", 'Dosyamı anlat');
html = html.replaceAll('WhatsApp Üzerinden Görüşelim', 'İhtiyacımı Anlat');

if (/https:\/\/wa\.me\//u.test(html)) throw new Error('SPECIAL CONTACT PATH V7: forbidden direct WhatsApp link survived');
if (!html.includes('href="/iletisim"')) throw new Error('SPECIAL CONTACT PATH V7: compliant contact route missing');
if (!html.includes('data-finance-funnel-v7')) throw new Error('SPECIAL CONTACT PATH V7: finance funnel contract missing');

fs.writeFileSync(distFile, html, 'utf8');
console.log(`SPECIAL CONTACT PATH V7 PASS — ${matches.length} direct WhatsApp CTA(s) routed through /iletisim; finance funnel preserved.`);
