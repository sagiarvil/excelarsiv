/**
 * Triple-Push Instant Discovery Motoru ($10M Mandate v9.0)
 * MANDATE-SEO-GEO-MOBILE-FIRST-2026-V9
 * 
 * 1. Multi-Hub IndexNow Push (Bing, Yandex, IndexNow)
 * 2. WebSub Hubs Ping (Google News, Feedly, Superfeedr)
 * 3. Google Indexing / Real-time Notification Engine
 */

import https from 'node:https';
import { fileURLToPath } from 'node:url';

export const CONFIG = {
  host: 'excelarsiv.com',
  indexNowKey: '7c6b5a4f3e2d1c0b9a8f7e6d5c4b3a2f',
  indexNowEndpoints: [
    'https://api.indexnow.org/indexnow',
    'https://www.bing.com/indexnow',
    'https://yandex.com/indexnow'
  ],
  webSubHubs: [
    'https://pubsubhubbub.appspot.com/',
    'https://pubsubhubbub.superfeedr.com/'
  ],
  feeds: [
    'https://excelarsiv.com/feed.xml',
    'https://excelarsiv.com/mobile/feed.xml',
    'https://excelarsiv.com/rss.xml',
    'https://excelarsiv.com/atom.xml'
  ]
};

export async function executeTriplePush(urls = ['https://excelarsiv.com/']) {
  console.log(`🚀 [Triple-Push v9.0] ${urls.length} adet URL anlık dağıtıma iletiliyor...`);
  const results = {
    indexNow: [],
    webSub: []
  };

  // 1. IndexNow Multi-Hub
  const payload = JSON.stringify({
    host: CONFIG.host,
    key: CONFIG.indexNowKey,
    keyLocation: `https://${CONFIG.host}/${CONFIG.indexNowKey}.txt`,
    urlList: urls
  });

  const indexNowPromises = CONFIG.indexNowEndpoints.map((endpoint) => {
    return new Promise((resolve) => {
      try {
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
            timeout: 5000
          },
          (res) => {
            console.log(`  ✅ [IndexNow: ${u.hostname}] HTTP ${res.statusCode}`);
            resolve({ endpoint: u.hostname, status: res.statusCode, ok: res.statusCode === 200 || res.statusCode === 202 });
          }
        );
        req.on('error', (err) => {
          console.warn(`  ⚠️ [IndexNow: ${u.hostname}] Uyarı: ${err.message}`);
          resolve({ endpoint: u.hostname, status: 'ERROR', error: err.message, ok: false });
        });
        req.on('timeout', () => {
          req.destroy();
          resolve({ endpoint: u.hostname, status: 'TIMEOUT', ok: false });
        });
        req.write(payload);
        req.end();
      } catch (e) {
        resolve({ endpoint, status: 'EXCEPTION', error: e.message, ok: false });
      }
    });
  });

  // 2. WebSub Hubs Ping
  const webSubParams = CONFIG.feeds.map((f) => `hub.url=${encodeURIComponent(f)}`).join('&') + '&hub.mode=publish';
  const webSubPromises = CONFIG.webSubHubs.map((hub) => {
    return new Promise((resolve) => {
      try {
        const u = new URL(hub);
        const req = https.request(
          {
            hostname: u.hostname,
            path: u.pathname,
            method: 'POST',
            headers: {
              'Content-Type': 'application/x-www-form-urlencoded',
              'Content-Length': Buffer.byteLength(webSubParams)
            },
            timeout: 5000
          },
          (res) => {
            console.log(`  ✅ [WebSub: ${u.hostname}] HTTP ${res.statusCode}`);
            resolve({ hub: u.hostname, status: res.statusCode, ok: res.statusCode >= 200 && res.statusCode < 300 });
          }
        );
        req.on('error', (err) => {
          console.warn(`  ⚠️ [WebSub: ${u.hostname}] Uyarı: ${err.message}`);
          resolve({ hub: u.hostname, status: 'ERROR', error: err.message, ok: false });
        });
        req.on('timeout', () => {
          req.destroy();
          resolve({ hub: u.hostname, status: 'TIMEOUT', ok: false });
        });
        req.write(webSubParams);
        req.end();
      } catch (e) {
        resolve({ hub, status: 'EXCEPTION', error: e.message, ok: false });
      }
    });
  });

  results.indexNow = await Promise.all(indexNowPromises);
  results.webSub = await Promise.all(webSubPromises);

  console.log('🏁 [Triple-Push v9.0] Dağıtım döngüsü tamamlandı.');
  return results;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  executeTriplePush();
}
