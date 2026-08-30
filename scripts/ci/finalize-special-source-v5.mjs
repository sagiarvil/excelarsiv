#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

const sourceFile = path.resolve('src/pages/ozel-excel-sistemleri.astro');
const distFile = path.resolve('dist/ozel-excel-sistemleri/index.html');

if (!fs.existsSync(sourceFile)) throw new Error('SPECIAL SOURCE V5: source file missing');
if (!fs.existsSync(distFile)) throw new Error('SPECIAL SOURCE V5: dist route missing');

const source = fs.readFileSync(sourceFile, 'utf8');
const rigidDeliveryColumn = 'grid-template-columns:minmax(0,1.2fr) minmax(300px,.8fr)';
const fluidDeliveryColumn = 'grid-template-columns:minmax(0,1.2fr) minmax(0,.8fr)';
const rigidDeliveryCount = source.split(rigidDeliveryColumn).length - 1;
if (rigidDeliveryCount > 1) {
  throw new Error(`SPECIAL SOURCE V5: unexpected rigid delivery column count: ${rigidDeliveryCount}`);
}

const oldAreaSection = `    <section class="section" id="alanlar">
      <div class="wrap">
        <div class="section-head"><h2>En sık hangi alanlarda sistem kuruyoruz?</h2><p>Tek bir konu için de çalışabiliriz; birbirine bağlı birkaç süreci aynı dosyada da toplayabiliriz.</p></div>
        <div class="area-grid">
          <article class="area-card"><h3>Nakit ve ödeme planı</h3><p>Kasa, banka, bekleyen tahsilat, yapılacak ödeme ve kısa vadeli nakit görünümü.</p></article>
          <article class="area-card"><h3>Cari ve tahsilat</h3><p>Müşteri bakiyesi, yaşlandırma, tahsilat takibi, ekstre ve geciken alacak kontrolü.</p></article>
          <article class="area-card"><h3>Banka ve kredi</h3><p>Limit, risk, taksit, kullanılabilir alan, ödeme takvimi ve banka bazlı karşılaştırma.</p></article>
          <article class="area-card"><h3>Kârlılık ve maliyet</h3><p>Ürün, müşteri, proje veya iş bazında gelir, maliyet, gider, marj ve sapma takibi.</p></article>
          <article class="area-card"><h3>Muhasebe kontrolü</h3><p>Eksik tanım, hatalı kayıt, mutabakat, kapanış kontrolü ve dönemsel kontrol listeleri.</p></article>
          <article class="area-card"><h3>Yönetim raporları</h3><p>Operasyon verisinden beslenen özetler, göstergeler ve düzenli yönetim raporları.</p></article>
        </div>
      </div>
    </section>`;

const newAreaSection = `    <section class="section" id="alanlar">
      <div class="wrap">
        <div class="section-head"><h2>En sık hangi alanlarda sistem kuruyoruz?</h2><p>Tek bir konu için de çalışabiliriz; birbirine bağlı birkaç süreci aynı dosyada da toplayabiliriz.</p></div>
        <div class="area-grid">
          <article class="area-card"><h3>Dinamik nakit akışı (13-Week)</h3><p>13 haftalık yuvarlanan likidite projeksiyonu, nakit açığı erken uyarısı, senaryolu tahsilat ve dinamik ödeme planı.</p></article>
          <article class="area-card"><h3>Banka limit-risk ve teminat</h3><p>Nakdi/gayrinakdi limitler, rotatif faiz yükü, kredi taksitleri, çek risk havuzu ve teminat karşılama oranları.</p></article>
          <article class="area-card"><h3>Birim maliyet ve dinamik fiyatlama</h3><p>Değişken maliyet endeksli teklif motoru, ürün ve müşteri bazlı kârlılık marjı, başabaş (break-even) simülatörü.</p></article>
          <article class="area-card"><h3>Senaryolu bütçe ve projeksiyon</h3><p>Kur ve enflasyon şoklarına duyarlı 3 senaryolu dinamik bütçe, fiili-bütçe sapma analizi ve rolling forecast.</p></article>
          <article class="area-card"><h3>Yatırım fizibilitesi ve değerleme</h3><p>İskontolanmış nakit akımları (DCF), IRR, geri ödeme süresi, EBITDA normalizasyonu ve borç servis (DSCR) modeli.</p></article>
          <article class="area-card"><h3>ERP entegre CFO yönetim kokpiti</h3><p>Logo, SAP, Mikro veya Zirve verisinden beslenen net işletme sermayesi, DSO takibi ve anlık yönetici KPI paneli.</p></article>
        </div>
      </div>
    </section>`;

const areaSectionCount = source.split(oldAreaSection).length - 1;
if (areaSectionCount !== 1) {
  throw new Error(`SPECIAL SOURCE V5: expected exactly one legacy area section, found ${areaSectionCount}`);
}

// Release artifact must never preserve the old 300px side-column floor. The replacement
// is deterministic and narrow so an unrelated CSS rule cannot be silently rewritten.
const html = source
  .replace(/^---[\s\S]*?---\s*/u, '')
  .replace(rigidDeliveryColumn, fluidDeliveryColumn)
  .replace(oldAreaSection, newAreaSection);

const required = [
  '<!doctype html>',
  'data-special-source-v5',
  'class="workbook"',
  '--font:-apple-system,BlinkMacSystemFont',
  'font-size:17px',
  'Hazır dosya işinize uymuyorsa, Excel\'i işinize göre kuralım.',
  'Özel bir sisteme ne zaman ihtiyaç olur?',
  'Hazır şablon, standart geliştirme ve size özel Excel aynı şey değil.',
  'Dinamik nakit akışı (13-Week)',
  'Banka limit-risk ve teminat',
  'Birim maliyet ve dinamik fiyatlama',
  'Senaryolu bütçe ve projeksiyon',
  'Yatırım fizibilitesi ve değerleme',
  'ERP entegre CFO yönetim kokpiti',
  '@media(max-width:800px)',
  'data-label="Size özel sistem"',
  fluidDeliveryColumn,
];
for (const token of required) {
  if (!html.includes(token)) throw new Error(`SPECIAL SOURCE V5: required token missing: ${token}`);
}

const forbidden = [
  '@font-face',
  'Manrope',
  'IBM Plex Mono',
  'data-special-light-legacy-bridge',
  'data-special-innovation=',
  'id="special-innovation-css"',
  'id="special-brand-sync-css"',
  'id="special-layout-stabilizer-css"',
  'id="native-seo-infographics"',
  'id="special-innovation-js"',
  'karar-laboratuvari',
  'min-width:780px',
  'min-width:850px',
  'diagnosis-head',
  'intent-card',
  'special-page-v4',
  rigidDeliveryColumn,
];
for (const token of forbidden) {
  if (html.includes(token)) throw new Error(`SPECIAL SOURCE V5: forbidden legacy token in release artifact: ${token}`);
}

// Taşma riski: büyük sabit width/min-width. max-width ve media breakpoint'leri
// responsive sınırlardır ve bu gate tarafından hata sayılmaz.
const fixedWidthRisk = /(?:^|[;{])\s*(?:width|min-width):\s*(?:7\d{2}|8\d{2}|9\d{2}|1\d{3,})px/gi;
const riskyMatches = [...html.matchAll(fixedWidthRisk)].map((match) => match[0].trim());
if (riskyMatches.length) {
  throw new Error(`SPECIAL SOURCE V5: large rigid width risk: ${riskyMatches.join(', ')}`);
}

const unbalancedGridRisk = /grid-template-columns:[^;}]*\b(?:300|320|340|360|400)px\b/gi;
if (unbalancedGridRisk.test(html)) throw new Error('SPECIAL SOURCE V5: rigid side column detected');

fs.mkdirSync(path.dirname(distFile), { recursive: true });
fs.writeFileSync(distFile, html, 'utf8');

const finalHtml = fs.readFileSync(distFile, 'utf8');
for (const token of required) {
  if (!finalHtml.includes(token)) throw new Error(`SPECIAL SOURCE V5: final artifact lost required token: ${token}`);
}
for (const token of forbidden) {
  if (finalHtml.includes(token)) throw new Error(`SPECIAL SOURCE V5: final artifact contains forbidden token: ${token}`);
}

console.log('SPECIAL SOURCE V5 PASS — Apple system typography, fluid two-column delivery, responsive comparison, finance-focused system areas and legacy post-build visual mutations removed from final route.');
