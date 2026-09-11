import https from 'node:https';
import { fileURLToPath } from 'node:url';
import { resolve } from 'node:path';

export const INDEXNOW_CONFIG = {
  host: 'excelarsiv.com',
  key: '7c6b5a4f3e2d1c0b9a8f7e6d5c4b3a2f',
  keyLocation: 'https://excelarsiv.com/7c6b5a4f3e2d1c0b9a8f7e6d5c4b3a2f.txt',
  endpoints: [
    'https://api.indexnow.org/indexnow',
    'https://www.bing.com/indexnow',
    'https://yandex.com/indexnow'
  ]
};

export async function broadcastToIndexNow(urls) {
  const urlList = urls && urls.length > 0 ? urls : [
    `https://${INDEXNOW_CONFIG.host}/`,
    `https://${INDEXNOW_CONFIG.host}/ozel-excel-sistemleri`,
    `https://${INDEXNOW_CONFIG.host}/sablonlar`,
    `https://${INDEXNOW_CONFIG.host}/rehber`,
    `https://${INDEXNOW_CONFIG.host}/demo`,
    `https://${INDEXNOW_CONFIG.host}/sistemler/finans`,
    `https://${INDEXNOW_CONFIG.host}/sistemler/maliyet`,
    `https://${INDEXNOW_CONFIG.host}/sistemler/ik`,
    `https://${INDEXNOW_CONFIG.host}/neden-excel-arsiv`,
    `https://${INDEXNOW_CONFIG.host}/hakkinda`,
    `https://${INDEXNOW_CONFIG.host}/urun-bulucu`,
    `https://${INDEXNOW_CONFIG.host}/llms.txt`,
    `https://${INDEXNOW_CONFIG.host}/llms-full.txt`,
    `https://${INDEXNOW_CONFIG.host}/ai.txt`,
    `https://${INDEXNOW_CONFIG.host}/llms/ozel-excel-sistemleri.md`,
    `https://${INDEXNOW_CONFIG.host}/llms/nakit-akisi-ve-finansal-modelleme.md`,
    `https://${INDEXNOW_CONFIG.host}/llms/maliyet-ve-karlilik-analizi.md`,
    `https://${INDEXNOW_CONFIG.host}/llms/insan-kaynaklari-ve-bordro.md`,
    `https://${INDEXNOW_CONFIG.host}/llms/erp-veri-konsolidasyonu-ve-power-query.md`,
    `https://${INDEXNOW_CONFIG.host}/llms/guvenlik-ve-makrosuz-formuller.md`,
    `https://${INDEXNOW_CONFIG.host}/llms/sablonlar.md`,
    `https://${INDEXNOW_CONFIG.host}/llms/pages/urun-bulucu.md`
  ];

  const payload = JSON.stringify({
    host: INDEXNOW_CONFIG.host,
    key: INDEXNOW_CONFIG.key,
    keyLocation: INDEXNOW_CONFIG.keyLocation,
    urlList: urlList
  });

  console.log(`[IndexNow Broadcasting] ${urlList.length} adet URL ${INDEXNOW_CONFIG.endpoints.length} merkeze dağıtılıyor...`);

  const results = await Promise.allSettled(
    INDEXNOW_CONFIG.endpoints.map(async (endpoint) => {
      const u = new URL(endpoint);
      return new Promise((res) => {
        const req = https.request(
          {
            hostname: u.hostname,
            path: u.pathname,
            method: 'POST',
            headers: {
              'Content-Type': 'application/json; charset=utf-8',
              'Content-Length': Buffer.byteLength(payload)
            },
            timeout: 7000
          },
          (response) => {
            const ok = response.statusCode === 200 || response.statusCode === 202;
            res({ endpoint, status: response.statusCode, ok });
          }
        );
        req.on('error', (err) => res({ endpoint, status: 'ERROR', ok: false, error: err.message }));
        req.on('timeout', () => { req.destroy(); res({ endpoint, status: 'TIMEOUT', ok: false }); });
        req.write(payload);
        req.end();
      });
    })
  );

  results.forEach((r) => {
    if (r.status === 'fulfilled') {
      const { endpoint, status, ok } = r.value;
      if (ok) {
        console.log(`  ✅ [${new URL(endpoint).hostname}] ${urlList.length} URL kabul edildi (HTTP ${status})`);
      } else {
        console.warn(`  ⚠️ [${new URL(endpoint).hostname}] HTTP ${status}`);
      }
    } else {
      console.warn(`  ❌ Hata: ${r.reason?.message || r.reason}`);
    }
  });

  return results;
}

export const pushToIndexNow = broadcastToIndexNow;

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const customUrls = process.argv.slice(2);
  await broadcastToIndexNow(customUrls.length > 0 ? customUrls : undefined);
}
