import { productSeo } from '../../src/data/productSeo.ts';

let failures = [];
const queries = new Set();

for (const [key, entry] of Object.entries(productSeo)) {
  if (!entry.primaryQuery) failures.push(`Missing primaryQuery in ${key}`);
  else {
    if (queries.has(entry.primaryQuery)) failures.push(`Duplicate primaryQuery in ${key}: ${entry.primaryQuery}`);
    queries.add(entry.primaryQuery);
  }
  if (!entry.title || entry.title.length < 20 || entry.title.length > 70) failures.push(`Title length invalid (20-70) for ${key}`);
  if (!entry.description || entry.description.length < 100 || entry.description.length > 160) failures.push(`Description length invalid (100-160) for ${key}`);
  if (!entry.answerSummary || entry.answerSummary.length < 160 || entry.answerSummary.length > 320) failures.push(`Answer summary length invalid (160-320) for ${key}`);
  if (entry.answerSummary && entry.answerSummary.includes('lorem')) failures.push(`Placeholder text found in ${key}`);
}
console.log(`Answer Graph Validation: ${failures.length} errors`);
if (failures.length > 0) {
  for (const f of failures) console.error(f);
  process.exit(1);
}
console.log('Answer Graph Validation PASS');
