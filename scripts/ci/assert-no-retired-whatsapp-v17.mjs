#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';

const root = process.cwd();
const retired = '90541' + '9305372';
const active = '905393333303';
const ignoredDirs = new Set(['.git','node_modules','dist','.astro']);
const allowedExt = new Set(['.astro','.js','.mjs','.cjs','.ts','.tsx','.jsx','.json','.md','.html','.css','.yml','.yaml']);
const offenders = [];

const walk = (dir) => {
  for (const entry of fs.readdirSync(dir, {withFileTypes:true})) {
    if (ignoredDirs.has(entry.name)) continue;
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) { walk(full); continue; }
    if (!allowedExt.has(path.extname(entry.name))) continue;
    const text = fs.readFileSync(full, 'utf8');
    if (text.includes(retired)) offenders.push(path.relative(root, full));
  }
};
walk(root);
if (offenders.length) throw new Error(`WHATSAPP V17: retired number found in source: ${offenders.join(', ')}`);

for (const file of [path.resolve('dist/index.html'), path.resolve('dist/ozel-excel-sistemleri/index.html')]) {
  if (!fs.existsSync(file)) throw new Error(`WHATSAPP V17: output missing: ${file}`);
  const html = fs.readFileSync(file, 'utf8');
  if (html.includes(retired)) throw new Error(`WHATSAPP V17: retired number leaked into ${file}`);
}

const special = fs.readFileSync(path.resolve('dist/ozel-excel-sistemleri/index.html'), 'utf8');
if (!special.includes(`https://wa.me/${active}?text=`)) throw new Error('WHATSAPP V17: active verified number missing from special page');
console.log('WHATSAPP V17 PASS — retired number absent from repository sources and final outputs; verified number active.');
