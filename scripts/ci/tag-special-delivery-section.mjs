#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

const file = path.resolve('dist/ozel-excel-sistemleri/index.html');
if (!fs.existsSync(file)) throw new Error('SPECIAL DELIVERY TAG: build output missing');

let html = fs.readFileSync(file, 'utf8');
if (!html.includes('data-living-workbook-v1') || !html.includes('deliver-grid')) {
  throw new Error('SPECIAL DELIVERY TAG: Living Workbook delivery anchor missing');
}

const deliverIndex = html.indexOf('deliver-grid');
const sectionStart = html.lastIndexOf('<section', deliverIndex);
if (sectionStart < 0) throw new Error('SPECIAL DELIVERY TAG: delivery parent section start missing');
const sectionEnd = html.indexOf('>', sectionStart);
if (sectionEnd < 0 || sectionEnd > deliverIndex) throw new Error('SPECIAL DELIVERY TAG: delivery parent section opening tag invalid');

let opening = html.slice(sectionStart, sectionEnd + 1);
if (!/\bclass=(['"])/i.test(opening)) {
  opening = opening.replace('<section', '<section class="delivery"');
} else {
  opening = opening.replace(/\bclass=(['"])(.*?)\1/i, (_all, quote, classes) => {
    const set = new Set(classes.split(/\s+/).filter(Boolean));
    set.add('delivery');
    return `class=${quote}${[...set].join(' ')}${quote}`;
  });
}
if (!/\bid=(['"])teslim\1/i.test(opening)) {
  opening = opening.replace('<section', '<section id="teslim"');
}

html = html.slice(0, sectionStart) + opening + html.slice(sectionEnd + 1);

const taggedIndex = html.indexOf('id="teslim"');
const deliverIndexAfter = html.indexOf('deliver-grid');
if (taggedIndex < 0 || taggedIndex > deliverIndexAfter || deliverIndexAfter - taggedIndex > 300) {
  throw new Error('SPECIAL DELIVERY TAG: id/class was not attached to the delivery parent section');
}
if (!html.slice(Math.max(0, taggedIndex - 100), deliverIndexAfter).includes('delivery')) {
  throw new Error('SPECIAL DELIVERY TAG: delivery class missing on tagged parent section');
}

fs.writeFileSync(file, html);
console.log('SPECIAL DELIVERY TAG PASS — delivery section tagged deterministically from deliver-grid anchor.');
