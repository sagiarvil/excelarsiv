/**
 * turkish_data_integrity.js — Turkish Character & JSON Data Integrity Enforcer
 * RFC 8259 / UTF-8 / Mojibake Auto-Repair / Turkish I-Case Mapping
 */

'use strict';

const fs = require('fs');
const path = require('path');

const MOJIBAKE_MAP = {
  'Ý': 'İ',
  'Ð': 'Ğ',
  'Þ': 'Ş',
  'ý': 'ı',
  'ð': 'ğ',
  'þ': 'ş',
};

const MOJIBAKE_REGEX = /[ÝÐÞýðþ]/g;

function fixMojibake(str) {
  if (typeof str !== 'string') return str;
  return str.replace(MOJIBAKE_REGEX, (ch) => MOJIBAKE_MAP[ch] || ch);
}

function toTurkishUpper(str) {
  if (typeof str !== 'string') return str;
  return str.toLocaleUpperCase('tr-TR');
}

function toTurkishLower(str) {
  if (typeof str !== 'string') return str;
  return str.toLocaleLowerCase('tr-TR');
}

function enforceHeaders(headers = {}) {
  const norm = { ...headers };
  norm['Content-Type'] = 'application/json; charset=utf-8';
  norm['Accept'] = 'application/json; charset=utf-8';
  return norm;
}

function cleanObjectMojibake(obj) {
  if (obj === null || obj === undefined) return obj;
  if (typeof obj === 'string') return fixMojibake(obj);
  if (Array.isArray(obj)) return obj.map(cleanObjectMojibake);
  if (typeof obj === 'object') {
    const res = {};
    for (const [k, v] of Object.entries(obj)) {
      res[fixMojibake(k)] = cleanObjectMojibake(v);
    }
    return res;
  }
  return obj;
}

function safeJsonStringify(data, space = 2) {
  const cleaned = cleanObjectMojibake(data);
  return JSON.stringify(cleaned, null, space);
}

module.exports = {
  fixMojibake,
  toTurkishUpper,
  toTurkishLower,
  enforceHeaders,
  cleanObjectMojibake,
  safeJsonStringify,
};

if (require.main === module) {
  const fileArg = process.argv[2];
  if (!fileArg) {
    console.log('Kullanım: node turkish_data_integrity.js <dosya.json|dosya.txt>');
    console.log('Örnek test: "Ýstanbul, Aðrý, Þanlýurfa" -> ' + fixMojibake('Ýstanbul, Aðrý, Þanlýurfa'));
    process.exit(0);
  }
  const full = path.resolve(fileArg);
  if (!fs.existsSync(full)) {
    console.error('Dosya bulunamadı:', full);
    process.exit(1);
  }
  const raw = fs.readFileSync(full, 'utf8');
  const fixed = fixMojibake(raw);
  if (raw !== fixed) {
    fs.writeFileSync(full, fixed, 'utf8');
    console.log('✅ Mojibake onarıldı ve UTF-8 olarak kaydedildi:', full);
  } else {
    console.log('ℹ️  Bozuk karakter tespit edilmedi, dosya temiz:', full);
  }
}
