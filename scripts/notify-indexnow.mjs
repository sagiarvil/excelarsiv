import { readFileSync, existsSync } from 'node:fs';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = resolve(fileURLToPath(new URL('../', import.meta.url)));
const HOST = 'excelarsiv.com';
const BASE_URL = `https://${HOST}`;
const KEY_FILE = resolve(ROOT, 'public/indexnow-key.txt');
const INDEXNOW_KEY = existsSync(KEY_FILE) ? readFileSync(KEY_FILE, 'utf8').trim() : '7c6b5a4f3e2d1c0b9a8f7e6d5c4b3a2f';

const ENDPOINTS = [
  'https://api.indexnow.org/indexnow',
  'https://www.bing.com/indexnow',
  'https://yandex.com/indexnow'
];

export async function pushToIndexNow(urls) {
  const urlList = urls && urls.length > 0 ? urls : [
    `${BASE_URL}/`,
    `${BASE_URL}/ozel-excel-sistemleri`,
    `${BASE_URL}/sablonlar`,
    `${BASE_URL}/rehber`,
    `${BASE_URL}/demo`,
    `${BASE_URL}/sistemler/finans`,
    `${BASE_URL}/sistemler/maliyet`,
    `${BASE_URL}/sistemler/ik`,
    `${BASE_URL}/neden-excel-arsiv`,
    `${BASE_URL}/hakkinda`,
    `${BASE_URL}/llms.txt`,
    `${BASE_URL}/llms-full.txt`,
    `${BASE_URL}/ai.txt`,
    `${BASE_URL}/llms/ozel-excel-sistemleri.md`,
    `${BASE_URL}/llms/nakit-akisi-ve-finansal-modelleme.md`,
    `${BASE_URL}/llms/maliyet-ve-karlilik-analizi.md`,
    `${BASE_URL}/llms/insan-kaynaklari-ve-bordro.md`,
    `${BASE_URL}/llms/erp-veri-konsolidasyonu-ve-power-query.md`,
    `${BASE_URL}/llms/guvenlik-ve-makrosuz-formuller.md`,
    `${BASE_URL}/llms/sablonlar.md`
  ];

  const payload = {
    host: HOST,
    key: INDEXNOW_KEY,
    keyLocation: `${BASE_URL}/indexnow-key.txt`,
    urlList: urlList
  };

  console.log(`[IndexNow Broadcasting] ${urlList.length} adet URL küresel IndexNow & AI Hub ağlarına dağıtılıyor...`);

  const results = await Promise.allSettled(
    ENDPOINTS.map(async (endpoint) => {
      const res = await fetch(endpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json; charset=utf-8' },
        body: JSON.stringify(payload)
      });
      return { endpoint, status: res.status, ok: res.ok || res.status === 200 || res.status === 202 };
    })
  );

  results.forEach((r) => {
    if (r.status === 'fulfilled') {
      const { endpoint, status, ok } = r.value;
      if (ok) {
        console.log(`  ✅ [${new URL(endpoint).hostname}] ${urlList.length} URL kabul edildi (HTTP ${status})`);
      } else {
        console.warn(`  ⚠️ [${new URL(endpoint).hostname}] Yanıt: HTTP ${status}`);
      }
    } else {
      console.warn(`  ❌ Ağ Hatası: ${r.reason?.message || r.reason}`);
    }
  });
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  await pushToIndexNow();
}
