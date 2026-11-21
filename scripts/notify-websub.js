/**
 * WebSub / PubSubHubbub Anlık Bildirim Motoru
 * MANDATE-SEO-GEO-MOBILE-FIRST-2026-V9
 */

import https from 'node:https';
import { fileURLToPath } from 'node:url';

export const WEBSUB_CONFIG = {
  hubs: [
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

export async function pingWebSubHubs(feeds = WEBSUB_CONFIG.feeds) {
  console.log(`📡 [WebSub] ${feeds.length} adet besleme için ${WEBSUB_CONFIG.hubs.length} hub pingleniyor...`);
  const params = feeds.map((f) => `hub.url=${encodeURIComponent(f)}`).join('&') + '&hub.mode=publish';

  const promises = WEBSUB_CONFIG.hubs.map((hub) => {
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
              'Content-Length': Buffer.byteLength(params)
            },
            timeout: 5000
          },
          (res) => {
            const ok = res.statusCode >= 200 && res.statusCode < 300;
            console.log(`  ${ok ? '✅' : '⚠️'} [WebSub: ${u.hostname}] HTTP ${res.statusCode}`);
            resolve({ hub: u.hostname, status: res.statusCode, ok });
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
        req.write(params);
        req.end();
      } catch (e) {
        resolve({ hub, status: 'EXCEPTION', error: e.message, ok: false });
      }
    });
  });

  const results = await Promise.all(promises);
  return results;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  pingWebSubHubs();
}
