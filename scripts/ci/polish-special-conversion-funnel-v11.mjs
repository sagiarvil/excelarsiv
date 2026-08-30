#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

const distFile = path.resolve('dist/ozel-excel-sistemleri/index.html');
if (!fs.existsSync(distFile)) throw new Error('SPECIAL CONVERSION FUNNEL V11: dist route missing');

let html = fs.readFileSync(distFile, 'utf8');

const bodyMatch = html.match(/<body\b[^>]*>/u);
if (!bodyMatch) throw new Error('SPECIAL CONVERSION FUNNEL V11: body tag missing');
let bodyTag = bodyMatch[0];
for (const marker of ['data-special-source-v5','data-finance-funnel-v7','data-hero-proof-layout-v8','data-cfo-positioning-v9']) {
  if (!new RegExp(`\\b${marker}\\b`, 'u').test(bodyTag)) {
    throw new Error(`SPECIAL CONVERSION FUNNEL V11: prerequisite body marker missing: ${marker}`);
  }
}
if (!/\bdata-conversion-funnel-v11\b/u.test(bodyTag)) {
  const nextBodyTag = bodyTag.replace(/>$/u, ' data-conversion-funnel-v11>');
  html = html.replace(bodyTag, nextBodyTag);
  bodyTag = nextBodyTag;
}

const authorityOld = `<div class="authority-rail" aria-label="ExcelArşiv uzmanlık özeti">
          <div><strong>17 Yıl</strong><span>Ticari bankacılık, kredi, limit ve risk bakışı</span></div>
          <div><strong>Banka ↔ İşletme</strong><span>Finans, muhasebe ve saha dilini aynı modelde birleştirme</span></div>
          <div><strong>ERP → CFO Kararı</strong><span>Dağınık veriyi rapordan aksiyona taşıyan yönetim sistemi</span></div>
        </div>`;

const authorityNew = `<div class="authority-rail" aria-label="ExcelArşiv uzmanlık özeti">
          <div class="authority-item"><strong>17 Yıl:</strong><span>Ticari bankacılık, kredi, limit ve risk bakışı</span></div>
          <div class="authority-item"><strong>Banka ↔ İşletme:</strong><span>Finans, muhasebe ve saha dilini aynı modelde birleştirme</span></div>
          <div class="authority-item"><strong>ERP → CFO Kararı:</strong><span>Dağınık veriyi rapordan aksiyona taşıyan yönetim sistemi</span></div>
        </div>`;

if (!html.includes(authorityOld)) throw new Error('SPECIAL CONVERSION FUNNEL V11: authority rail source block drifted');
html = html.replace(authorityOld, authorityNew);

const waMessage = 'Merhaba Barış Bey, excelarsiv.com üzerinden ulaşıyorum. Şirketimiz için özel Excel finans sistemi hakkında görüşmek istiyorum.';
const waUrl = `https://wa.me/905419305372?text=${encodeURIComponent(waMessage)}`;
const waPattern = /https:\/\/wa\.me\/905419305372\?text=[^"'<>\s]+/gu;
const waMatches = html.match(waPattern) || [];
if (waMatches.length < 3) throw new Error(`SPECIAL CONVERSION FUNNEL V11: expected >=3 WhatsApp CTAs before rewrite, got ${waMatches.length}`);
html = html.replace(waPattern, waUrl);

const schemaPattern = /<script type="application\/ld\+json">([\s\S]*?)<\/script>/u;
const schemaMatch = html.match(schemaPattern);
if (!schemaMatch) throw new Error('SPECIAL CONVERSION FUNNEL V11: JSON-LD block missing');
let schema;
try {
  schema = JSON.parse(schemaMatch[1]);
} catch (error) {
  throw new Error(`SPECIAL CONVERSION FUNNEL V11: JSON-LD parse failed: ${error.message}`);
}
if (!schema || !Array.isArray(schema['@graph'])) throw new Error('SPECIAL CONVERSION FUNNEL V11: JSON-LD @graph missing');

const graph = schema['@graph'];
const personId = 'https://excelarsiv.com/#baris-bagirlar';
const person = graph.find((node) => node?.['@type'] === 'Person' && node?.['@id'] === personId);
if (!person) throw new Error('SPECIAL CONVERSION FUNNEL V11: Barış Bağırlar Person entity missing');
person.jobTitle = 'Ticari Bankacılık Uzmanı ve Finansal Sistem Mimarı';
person.description = '17 yıllık ticari bankacılık ve saha finans deneyimiyle şirketlerin nakit, kredi, limit-risk, maliyet, bütçe, yatırım ve yönetim raporlama ihtiyaçlarını özel Excel karar sistemlerine dönüştürür.';
person.knowsAbout = Array.from(new Set([
  ...(Array.isArray(person.knowsAbout) ? person.knowsAbout : []),
  '13 haftalık nakit akışı',
  'banka limit-risk ve teminat yönetimi',
  'rolling forecast ve senaryolu bütçe',
  'DCF, IRR ve DSCR finansal modelleme',
  'ERP verisinden CFO yönetim raporlaması',
]));

const professionalServiceId = 'https://excelarsiv.com/ozel-excel-sistemleri#professional-service';
const professionalService = {
  '@type': 'ProfessionalService',
  '@id': professionalServiceId,
  name: 'ExcelArşiv - Özel Excel Karar Sistemleri',
  url: 'https://excelarsiv.com/ozel-excel-sistemleri',
  founder: {'@id': personId},
  provider: {'@id': personId},
  description: '17 yıllık ticari bankacılık ve saha finans deneyimiyle şirketlere özel dinamik nakit akışı, banka limit-risk, maliyet, bütçe, değerleme ve ERP entegre CFO karar sistemleri tasarımı.',
  areaServed: {'@type': 'Country', name: 'Türkiye'},
  knowsAbout: [
    'Dinamik Nakit Akışı (13-Week Cash Flow)',
    'Banka Limit-Risk ve Teminat Havuzu',
    'Birim Maliyet ve Dinamik Fiyatlama',
    'Senaryolu Bütçe ve Rolling Forecast',
    'Yatırım Fizibilitesi ve DCF Değerleme',
    'ERP Entegre CFO Yönetim Kokpiti',
  ],
};
const existingProfessionalIndex = graph.findIndex((node) => node?.['@id'] === professionalServiceId || node?.['@type'] === 'ProfessionalService');
if (existingProfessionalIndex >= 0) graph[existingProfessionalIndex] = professionalService;
else graph.push(professionalService);

html = html.replace(schemaPattern, `<script type="application/ld+json">${JSON.stringify(schema)}</script>`);

const css = `
  <style id="special-conversion-funnel-v11-css">
    body[data-conversion-funnel-v11] .authority-rail>.authority-item{display:flex;align-items:flex-start;gap:10px;min-width:0}
    body[data-conversion-funnel-v11] .authority-rail>.authority-item strong{display:inline;flex:0 0 auto;white-space:nowrap;margin:0}
    body[data-conversion-funnel-v11] .authority-rail>.authority-item span{display:block;min-width:0;margin:1px 0 0;overflow-wrap:anywhere}

    body[data-conversion-funnel-v11] .hero-grid>.proof-grid{margin-bottom:14px}
    body[data-conversion-funnel-v11] .authority-comparison{padding-top:56px}

    body[data-conversion-funnel-v11] .workbook{width:100%;max-width:100%;min-width:0;overflow:hidden}
    body[data-conversion-funnel-v11] .wb-body,
    body[data-conversion-funnel-v11] .metric-grid,
    body[data-conversion-funnel-v11] .wb-chart-grid,
    body[data-conversion-funnel-v11] .wb-lists,
    body[data-conversion-funnel-v11] .panel,
    body[data-conversion-funnel-v11] .data-row,
    body[data-conversion-funnel-v11] .alert-row{min-width:0;max-width:100%}
    body[data-conversion-funnel-v11] .metric,
    body[data-conversion-funnel-v11] .panel,
    body[data-conversion-funnel-v11] .data-row,
    body[data-conversion-funnel-v11] .alert-row{overflow-wrap:anywhere}

    body[data-conversion-funnel-v11] .faq-grid{align-items:start}
    body[data-conversion-funnel-v11] .faq-copy,
    body[data-conversion-funnel-v11] .faq-list{align-self:start;min-width:0}

    @media(max-width:900px){
      body[data-conversion-funnel-v11] .authority-rail>.authority-item{gap:9px}
      body[data-conversion-funnel-v11] .authority-comparison{padding-top:50px}
    }

    @media(max-width:660px){
      body[data-conversion-funnel-v11] .workbook{width:100%;max-width:100%;border-radius:18px}
      body[data-conversion-funnel-v11] .wb-chart-grid,
      body[data-conversion-funnel-v11] .wb-lists{grid-template-columns:minmax(0,1fr)}
      body[data-conversion-funnel-v11] .hero-grid>.proof-grid{margin-bottom:10px}
      body[data-conversion-funnel-v11] .authority-comparison{padding-top:46px}
      body[data-conversion-funnel-v11] .authority-rail>.authority-item{display:grid;grid-template-columns:auto minmax(0,1fr);gap:8px}
    }

    @media(max-width:430px){
      body[data-conversion-funnel-v11] .metric-grid{grid-template-columns:minmax(0,1fr)}
    }
  </style>`;

if (!html.includes('</head>')) throw new Error('SPECIAL CONVERSION FUNNEL V11: head close missing');
html = html.replace('</head>', `${css}\n</head>`);

const finalBody = html.match(/<body\b[^>]*>/u)?.[0] || '';
if (!/\bdata-conversion-funnel-v11\b/u.test(finalBody)) throw new Error('SPECIAL CONVERSION FUNNEL V11: V11 body marker not bound');
for (const token of [
  '17 Yıl:</strong><span>',
  'Banka ↔ İşletme:</strong><span>',
  'ERP → CFO Kararı:</strong><span>',
  waUrl,
  '"@type":"ProfessionalService"',
  '"jobTitle":"Ticari Bankacılık Uzmanı ve Finansal Sistem Mimarı"',
  'body[data-conversion-funnel-v11] .workbook{width:100%;max-width:100%;min-width:0;overflow:hidden}',
  'body[data-conversion-funnel-v11] .faq-grid{align-items:start}',
]) {
  if (!html.includes(token)) throw new Error(`SPECIAL CONVERSION FUNNEL V11: required token missing: ${token}`);
}

const finalWaCount = (html.match(new RegExp(waUrl.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gu')) || []).length;
if (finalWaCount < 3) throw new Error(`SPECIAL CONVERSION FUNNEL V11: expected >=3 prefilled WhatsApp CTAs, got ${finalWaCount}`);
if (html.includes('overflow-x:hidden')) throw new Error('SPECIAL CONVERSION FUNNEL V11: forbidden global overflow hiding introduced');

fs.writeFileSync(distFile, html, 'utf8');
console.log(`SPECIAL CONVERSION FUNNEL V11 PASS — authority labels separated, ${finalWaCount} WhatsApp CTAs prefilled, mobile workbook bounded, FAQ top-aligned, ProfessionalService entity added.`);
