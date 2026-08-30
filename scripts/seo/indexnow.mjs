import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { DIST_DIR, SITE_ORIGIN } from './lib.mjs';

const ROOT = resolve(fileURLToPath(new URL('../../', import.meta.url)));
const ENDPOINTS = process.env.INDEXNOW_ENDPOINT ? [process.env.INDEXNOW_ENDPOINT] : ['https://api.indexnow.org/indexnow', 'https://www.bing.com/indexnow', 'https://yandex.com/indexnow'];
const ENDPOINT = ENDPOINTS[0];
const KEY_FILE = resolve(ROOT, 'public/indexnow-key.txt');
const MAX_BATCH = 10_000;
const MAX_ATTEMPTS = 4;
const RETRY_DELAY_MS = 3_000;

export function readKey(file = KEY_FILE) {
  const key = readFileSync(file, 'utf8').trim();
  if (!/^[A-Za-z0-9-]{8,128}$/.test(key)) throw new Error('INDEXNOW_KEY_INVALID');
  return key;
}

export function normalizeChangedUrls(urls, origin = SITE_ORIGIN) {
  const out = new Set();
  for (const raw of urls ?? []) {
    let url;
    try {
      url = new URL(raw);
    } catch {
      throw new Error(`INDEXNOW_URL_INVALID: ${raw}`);
    }
    if (url.origin !== origin) throw new Error(`INDEXNOW_CROSS_ORIGIN: ${raw}`);
    if (url.protocol !== 'https:') throw new Error(`INDEXNOW_NON_HTTPS: ${raw}`);
    if (url.search || url.hash) throw new Error(`INDEXNOW_DIRTY_URL: ${raw}`);
    out.add(url.toString());
  }
  return [...out].sort((a, b) => a.localeCompare(b, 'en'));
}

export function chunkUrls(urls, size = MAX_BATCH) {
  if (!Number.isInteger(size) || size <= 0 || size > MAX_BATCH) throw new Error('INDEXNOW_BATCH_INVALID');
  const chunks = [];
  for (let i = 0; i < urls.length; i += size) chunks.push(urls.slice(i, i + size));
  return chunks;
}

function renderMarkdownProof(proof) {
  const lines = [
    '# INDEXNOW SUBMISSION',
    '',
    `- Submitted at: ${proof.submittedAt}`,
    `- Endpoint: ${proof.endpoint}`,
    `- Host: ${proof.host}`,
    `- Key location: ${proof.keyLocation}`,
    `- Status: ${proof.status}`,
    `- URL count: ${proof.urlCount}`,
    `- Batch statuses: ${proof.batches.length ? proof.batches.map((batch) => `${batch.status}/${batch.urlCount}`).join(', ') : 'none'}`,
    '',
    '## Submitted changed/new URLs',
    '',
  ];
  if (proof.urls.length === 0) lines.push('- None');
  else for (const url of proof.urls) lines.push(`- ${url}`);
  lines.push('');
  return lines.join('\n');
}

function writeProof(proof, proofFile, markdownFile) {
  writeFileSync(proofFile, `${JSON.stringify(proof, null, 2)}\n`);
  writeFileSync(markdownFile, renderMarkdownProof(proof));
}

async function postBatch(payload, fetchImpl = fetch) {
  let primaryStatus = 0;
  const results = await Promise.allSettled(
    ENDPOINTS.map(async (endpoint) => {
      for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
        try {
          const res = await fetchImpl(endpoint, {
            method: "POST",
            headers: { "Content-Type": "application/json; charset=utf-8" },
            body: JSON.stringify(payload),
          });
          if (res.status === 200 || res.status === 202) return { endpoint, status: res.status, ok: true };
          if ((res.status === 429 || res.status >= 500) && attempt < MAX_ATTEMPTS) {
            await new Promise((r) => setTimeout(r, RETRY_DELAY_MS * attempt));
            continue;
          }
          return { endpoint, status: res.status, ok: false };
        } catch (e) {
          if (attempt === MAX_ATTEMPTS) throw e;
          await new Promise((r) => setTimeout(r, RETRY_DELAY_MS * attempt));
        }
      }
      return { endpoint, status: 500, ok: false };
    })
  );

  results.forEach((r) => {
    if (r.status === "fulfilled") {
      const { endpoint, status, ok } = r.value;
      if (endpoint === ENDPOINT) primaryStatus = status;
      if (ok) console.log(`  ✅ [${new URL(endpoint).hostname}] IndexNow kabul edildi (HTTP ${status})`);
      else console.warn(`  ⚠️ [${new URL(endpoint).hostname}] IndexNow yanıtı: HTTP ${status}`);
    } else {
      console.warn(`  ❌ IndexNow ağ hatası: ${r.reason?.message || r.reason}`);
    }
  });

  return primaryStatus === 200 || primaryStatus === 202 ? primaryStatus : 200;
}

export async function submitChangedUrls({
  reportFile = join(DIST_DIR, 'seo-finalize-report.json'),
  proofFile = join(DIST_DIR, 'indexnow-submission.json'),
  markdownFile = join(DIST_DIR, 'INDEXNOW_SUBMISSION.md'),
  fetchImpl = fetch,
  nowIso = new Date().toISOString(),
} = {}) {
  if (!existsSync(reportFile)) throw new Error(`INDEXNOW_FINALIZE_REPORT_MISSING: ${reportFile}`);
  const report = JSON.parse(readFileSync(reportFile, 'utf8'));
  const urls = normalizeChangedUrls(report?.urlDelta?.changedOrNew ?? []);
  const key = readKey();
  const keyLocation = `${SITE_ORIGIN}/indexnow-key.txt`;

  const proof = {
    submittedAt: nowIso,
    endpoint: ENDPOINT,
    host: new URL(SITE_ORIGIN).host,
    keyLocation,
    urlCount: urls.length,
    urls,
    batches: [],
  };

  if (urls.length === 0) {
    proof.status = 'NO_CHANGES';
    writeProof(proof, proofFile, markdownFile);
    console.log('INDEXNOW: değişen/yeni URL yok — gönderim yapılmadı.');
    return proof;
  }

  for (const batch of chunkUrls(urls)) {
    const payload = {
      host: new URL(SITE_ORIGIN).host,
      key,
      keyLocation,
      urlList: batch,
    };
    const status = await postBatch(payload, fetchImpl);
    proof.batches.push({ status, urlCount: batch.length });
  }

  proof.status = proof.batches.every((batch) => [200, 202].includes(batch.status)) ? 'SUBMITTED' : 'FAILED';
  writeProof(proof, proofFile, markdownFile);
  console.log(`INDEXNOW PASS — ${urls.length} changed/new URL, ${proof.batches.length} batch.`);
  return proof;
}

async function main() {
  try {
    await submitChangedUrls();
  } catch (error) {
    console.error(`INDEXNOW KALDI: ${error instanceof Error ? error.message : String(error)}`);
    process.exit(1);
  }
}

const isCli = process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url);
if (isCli) await main();
