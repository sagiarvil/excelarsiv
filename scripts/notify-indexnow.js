/**
 * MANDATE-SEO-GEO-2026-V6
 * Multi-Hub IndexNow Anlık Dağıtım Motoru
 * Değişen URL'leri eşzamanlı olarak global indeks merkezlerine (IndexNow, Bing, Yandex) push eder.
 */

import https from 'node:https';
import { fileURLToPath } from 'node:url';

export const INDEXNOW_CONFIG = {
  host: 'excelarsiv.com',
  key: '7c6b5a4f3e2d1c0b9a8f7e6d5c4b3a2f',
  endpoints: [
    'https://api.indexnow.org/indexnow',
    'https://www.bing.com/indexnow',
    'https://yandex.com/indexnow'
  ]
};

export async function broadcastToIndexNow(urlList) {
  if (!Array.isArray(urlList) || urlList.length === 0) {
    console.warn('⚠️ [IndexNow] Gönderilecek URL listesi boş.');
    return [];
  }

  const payload = JSON.stringify({
    host: INDEXNOW_CONFIG.host,
    key: INDEXNOW_CONFIG.key,
    keyLocation: `https://${INDEXNOW_CONFIG.host}/${INDEXNOW_CONFIG.key}.txt`,
    urlList: urlList
  });

  console.log(`🚀 [IndexNow] ${urlList.length} adet URL ${INDEXNOW_CONFIG.endpoints.length} merkeze yayınlanıyor...`);

  const promises = INDEXNOW_CONFIG.endpoints.map((endpoint) => {
    return new Promise((resolve) => {
      const u = new URL(endpoint);
      const req = https.request(
        {
          hostname: u.hostname,
          path: u.pathname,
          method: 'POST',
          headers: {
            'Content-Type': 'application/json; charset=utf-8',
            'Content-Length': Buffer.byteLength(payload)
          },
          timeout: 6000
        },
        (res) => {
          const isOk = res.statusCode === 200 || res.statusCode === 202;
          resolve({ host: u.hostname, status: res.statusCode, ok: isOk });
        }
      );
      req.on('error', (err) => resolve({ host: u.hostname, status: 'ERROR', message: err.message }));
      req.on('timeout', () => { req.destroy(); resolve({ host: u.hostname, status: 'TIMEOUT' }); });
      req.write(payload);
      req.end();
    });
  });

  const results = await Promise.allSettled(promises);
  results.forEach((r) => {
    if (r.status === 'fulfilled') {
      const { host, status, ok, message } = r.value;
      if (ok) {
        console.log(`  ✅ [${host}] Başarılı (HTTP ${status})`);
      } else {
        console.warn(`  ⚠️ [${host}] Bildirim hatası (HTTP ${status} - ${message || 'Hata'})`);
      }
    }
  });

  return results;
}

const isMain = process.argv[1] && (
  process.argv[1].endsWith('notify-indexnow.js') ||
  process.argv[1] === fileURLToPath(import.meta.url)
);

if (isMain) {
  const urls = process.argv.slice(2);
  if (urls.length > 0) {
    await broadcastToIndexNow(urls);
  } else {
    console.log('Kullanım: node scripts/notify-indexnow.js <URL1> <URL2>');
    console.log('Varsayılan amiral gemisi URL listesi yayınlanıyor...');
    await broadcastToIndexNow([
      'https://excelarsiv.com/',
      'https://excelarsiv.com/ozel-excel-sistemleri',
      'https://excelarsiv.com/sablonlar',
      'https://excelarsiv.com/rehber',
      'https://excelarsiv.com/demo',
      'https://excelarsiv.com/llms.txt',
      'https://excelarsiv.com/llms/core.md',
      'https://excelarsiv.com/llms/pages/ozel-excel-sistemleri.md'
    ]);
  }
}
