/**
 * read_memory.js — Proje Hafızasını ve Talimatları Okuma Aracı
 * 
 * Kullanım:
 *   node read_memory.js <proje_adi>
 */
'use strict';
const fs = require('fs');
const path = require('path');

const projectName = process.argv[2] || 'agents-skills-config';
const baseDir = path.resolve(__dirname, '..', 'messages');
const projectDir = path.join(baseDir, projectName);

console.log('📖 Proje Hafızası Getiriliyor: ' + projectName);

const okuFile = path.join(projectDir, '_OKU.txt');
if (fs.existsSync(okuFile)) {
  console.log('--- _OKU.txt ---');
  console.log(fs.readFileSync(okuFile, 'utf8'));
}

const talimatFile = path.join(projectDir, 'PROJE_TALİMATLARI.md');
if (fs.existsSync(talimatFile)) {
  console.log('--- PROJE_TALİMATLARI.md ---');
  console.log(fs.readFileSync(talimatFile, 'utf8'));
}
