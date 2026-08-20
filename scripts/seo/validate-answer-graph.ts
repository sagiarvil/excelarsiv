import { readdirSync } from 'node:fs';
import { extname, basename, resolve } from 'node:path';
import { productAnswerSeo } from '../../src/data/productAnswerSeo.ts';

const failures: string[] = [];
const entries = Object.entries(productAnswerSeo);
const seenQueries = new Map<string, string>();

for (const [slug, entry] of entries) {
  const query = entry.primaryQuery.trim().toLocaleLowerCase('tr-TR');
  if (!query) failures.push(`${slug}: primaryQuery boş`);
  if (seenQueries.has(query)) failures.push(`${slug}: duplicate primaryQuery (${query}) ayrıca ${seenQueries.get(query)}`);
  seenQueries.set(query, slug);

  if (entry.answerQuestion.length < 20 || entry.answerQuestion.length > 140) {
    failures.push(`${slug}: answerQuestion 20-140 karakter olmalı (${entry.answerQuestion.length})`);
  }
  if (entry.answerSummary.length < 80 || entry.answerSummary.length > 240) {
    failures.push(`${slug}: answerSummary 80-240 karakter olmalı (${entry.answerSummary.length})`);
  }
  if (!entry.answerSummary.toLocaleLowerCase('tr-TR').includes('excel')) {
    failures.push(`${slug}: answerSummary ürün bağlamını belirtmiyor (Excel kelimesi yok)`);
  }
}

const templateDir = resolve('src/content/templates');
const templateSlugs = readdirSync(templateDir)
  .filter((name) => ['.md', '.mdx'].includes(extname(name)))
  .map((name) => basename(name, extname(name)))
  .sort();

for (const slug of templateSlugs) {
  if (!productAnswerSeo[slug]) failures.push(`${slug}: Answer Graph kaydı eksik`);
}
for (const slug of Object.keys(productAnswerSeo)) {
  if (!templateSlugs.includes(slug)) failures.push(`${slug}: ürün kaydı var fakat template yok`);
}

console.log(`ANSWER GRAPH: ${entries.length} ürün, ${failures.length} hata`);
if (failures.length) {
  failures.forEach((failure) => console.error(`  - ${failure}`));
  process.exit(1);
}
console.log('ANSWER GRAPH PASS');
