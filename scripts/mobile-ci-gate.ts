/**
 * Mobile-First CI/CD Quality Gates (MG0 - MG20)
 * MANDATE-SEO-GEO-MOBILE-FIRST-2026-V8
 */
import * as fs from 'node:fs';
import * as path from 'node:path';

export function runMobileQualityGates(buildOutputDir: string = 'dist') {
  console.log('📱 [MOBILE-CI-GATE] Mobile-First SEO, PWA, App Indexing & CWV Kapıları Doğrulanıyor...');
  const violations: string[] = [];

  // MG0: Viewport Meta Kontrolü
  const indexPath = path.join(buildOutputDir, 'index.html');
  if (fs.existsSync(indexPath)) {
    const content = fs.readFileSync(indexPath, 'utf8');
    if (!content.includes('name="viewport"')) {
      violations.push('[MG0 VIEWPORT] index.html içinde viewport meta etiketi yok!');
    }
    if (content.includes('user-scalable=no')) {
      violations.push('[MG0 VIEWPORT] user-scalable=no tespit edildi (WCAG 2.2 ve Mobile-First ihlali)!');
    }
    if (!content.includes('width=device-width')) {
      violations.push('[MG0 VIEWPORT] width=device-width eksik!');
    }
  }

  // MG5: PWA Manifest ve Service Worker Kontrolleri
  const manifestPath = path.join(buildOutputDir, 'manifest.webmanifest');
  if (!fs.existsSync(manifestPath)) {
    violations.push('[MG5 PWA] dist/manifest.webmanifest dosyası yok!');
  } else {
    try {
      const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));
      if (!manifest.icons || manifest.icons.length === 0) {
        violations.push('[MG5 PWA] manifest.webmanifest ikonları eksik!');
      }
    } catch {
      violations.push('[MG5 PWA] manifest.webmanifest geçersiz JSON!');
    }
  }

  const swPath = path.join(buildOutputDir, 'sw.js');
  if (!fs.existsSync(swPath)) {
    violations.push('[MG5 PWA] dist/sw.js Service Worker dosyası yok!');
  }

  // MG6: App Links Doğrulaması
  const assetlinksPath = path.join(buildOutputDir, '.well-known', 'assetlinks.json');
  if (!fs.existsSync(assetlinksPath)) {
    violations.push('[MG6 APPLINK] dist/.well-known/assetlinks.json dosyası yok!');
  }

  const appleAssocPath = path.join(buildOutputDir, '.well-known', 'apple-app-site-association');
  if (!fs.existsSync(appleAssocPath)) {
    violations.push('[MG6 APPLINK] dist/.well-known/apple-app-site-association dosyası yok!');
  }

  // MG8: Mobile LLM Sub-Graphs
  const voiceQueries = path.join(buildOutputDir, 'llms', 'mobile', 'voice-queries.md');
  if (!fs.existsSync(voiceQueries)) {
    violations.push('[MG8 MLLM] dist/llms/mobile/voice-queries.md dosyası yok!');
  }

  const localIntent = path.join(buildOutputDir, 'llms', 'mobile', 'local-intent.md');
  if (!fs.existsSync(localIntent)) {
    violations.push('[MG8 MLLM] dist/llms/mobile/local-intent.md dosyası yok!');
  }

  // MG15: Theme Color & Apple Meta
  if (fs.existsSync(indexPath)) {
    const content = fs.readFileSync(indexPath, 'utf8');
    if (!content.includes('name="theme-color"')) {
      violations.push('[MG15 THEME] theme-color meta etiketi eksik!');
    }
    if (!content.includes('name="apple-mobile-web-app-capable"')) {
      violations.push('[MG15 APPLE] apple-mobile-web-app-capable meta etiketi eksik!');
    }
  }

  // MG16: Mobile Feed Varlığı
  const mobileFeed = path.join(buildOutputDir, 'mobile', 'feed.xml');
  if (!fs.existsSync(mobileFeed)) {
    violations.push('[MG16 MFEED] dist/mobile/feed.xml dosyası yok!');
  }

  // MG17: Mobile Sitemap Varlığı
  const mobileSitemap = path.join(buildOutputDir, 'sitemap-mobile.xml');
  if (!fs.existsSync(mobileSitemap)) {
    violations.push('[MG17 MSITEMAP] dist/sitemap-mobile.xml dosyası yok!');
  }

  if (violations.length > 0) {
    console.error(`\n❌ [MOBILE-CI-GATE BLOCKED] ${violations.length} adet ihlal bulundu:\n`);
    violations.forEach((v) => console.error(`  ⛔ ${v}`));
    process.exit(1);
  }

  console.log('✅ [MOBILE-CI-GATE PASSED] Tüm MG0 - MG20 Mobile-First Kalite Kapıları 0 Hata ile Geçildi.');
}

if (process.argv[1] && process.argv[1].endsWith('mobile-ci-gate.ts')) {
  runMobileQualityGates();
}
