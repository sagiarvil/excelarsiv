#!/usr/bin/env node
import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';

const EXPECTED_BYTES = 40921;
const EXPECTED_SHA256 = 'a91547d212116321681d10c2d6d643f1566038ef51cbf9cb881f0b5b993e0341';
const TARGET = join(process.cwd(), 'delivery', 'paid-products', 'logo-sql-cari-yaslandirma-tahsilat-karar-motoru', 'current.xlsx');
const SOURCE_COMMIT = 'f7db258433e0d5edab18e4c0a1b886e48d89c827';
const SOURCE_BASE = `https://raw.githubusercontent.com/barisbagirlar-web/excelarsiv/${SOURCE_COMMIT}/delivery/paid-products/logo-sql-cari-yaslandirma-tahsilat-karar-motoru/packed`;
const PARTS = ['current.xlsx.b64.00', 'current.xlsx.b64.01', 'current.xlsx.b64.02', 'current.xlsx.b64.03'];

function sha256(buffer) {
  return createHash('sha256').update(buffer).digest('hex');
}

function verify(buffer, label) {
  const digest = sha256(buffer);
  if (buffer.length !== EXPECTED_BYTES || digest !== EXPECTED_SHA256) {
    throw new Error(`${label}: bütünlük hatası bytes=${buffer.length}/${EXPECTED_BYTES} sha256=${digest}/${EXPECTED_SHA256}`);
  }
}

if (existsSync(TARGET)) {
  const current = readFileSync(TARGET);
  verify(current, 'mevcut Logo çalışma kitabı');
  console.log(`Logo workbook READY: ${current.length} bytes · ${EXPECTED_SHA256}`);
  process.exit(0);
}

const chunks = [];
for (const part of PARTS) {
  const url = `${SOURCE_BASE}/${part}`;
  const response = await fetch(url, { redirect: 'follow' });
  if (!response.ok) throw new Error(`${part}: kaynak alınamadı HTTP ${response.status}`);
  const text = (await response.text()).trim();
  if (!/^[A-Za-z0-9+/]+={0,2}$/.test(text)) throw new Error(`${part}: geçersiz base64 kaynak`);
  chunks.push(text);
}
const workbook = Buffer.from(chunks.join(''), 'base64');
verify(workbook, 'Logo çalışma kitabı');
mkdirSync(dirname(TARGET), { recursive: true });
writeFileSync(TARGET, workbook);
console.log(`Logo workbook MATERIALIZED: ${workbook.length} bytes · ${EXPECTED_SHA256} · source=${SOURCE_COMMIT}`);
