const fs = require('fs');

const html = fs.readFileSync('$HOME/Sites/satis/docs/satis_test_live.html', 'utf8');

console.log('================================================================');
console.log('BETASOFT SATIS (MACOS SUITE) - SUPER-MANDATE ENGINE V3.0 DENETİMİ');
console.log('================================================================\n');

// 1. BYTE SIZE (ENG-01: 14KB TCP Window)
const byteLength = Buffer.byteLength(html, 'utf8');
console.log('[ENG-01 & ENG-02] HTML Toplam Boyutu:', byteLength, 'bytes');
console.log('                  14KB Bütçe Durumu:', byteLength <= 14336 ? 'PASS (<= 14.336)' : 'FAIL (90KB, 14KB bütçesini aşıyor)');

// 2. CORE SEO (ENG-04)
const titleMatch = html.match(/<title>([^<]*)<\/title>/i);
console.log('\n[ENG-04] Title:', titleMatch ? titleMatch[1].trim() : 'YOK');

const descMatch = html.match(/<meta[^>]*name=["']description["'][^>]*content=["']([^"']*)["']/i);
console.log('[ENG-04] Meta Description:', descMatch ? descMatch[1].trim() : 'YOK');

const canonMatch = html.match(/<link[^>]*rel=["']canonical["'][^>]*href=["']([^"']*)["']/i);
console.log('[ENG-04] Canonical URL:', canonMatch ? canonMatch[1].trim() : 'YOK');

const h1Matches = [...html.matchAll(/<h1[^>]*>([\s\S]*?)<\/h1>/gi)];
console.log('[ENG-04] H1 Sayısı:', h1Matches.length);
h1Matches.forEach((m, i) => console.log('         H1 #' + (i+1) + ':', m[1].replace(/<[^>]+>/g, '').trim()));

// 3. ColBERT MaxSim Headings (ENG-10)
const h2Count = [...html.matchAll(/<h2[^>]*>/gi)].length;
const h3Count = [...html.matchAll(/<h3[^>]*>/gi)].length;
console.log('\n[ENG-10] ColBERT MaxSim Başlık Sayıları: H2 =', h2Count, ', H3 =', h3Count);

// 4. LLMO & Machine Surfaces (ENG-07)
const describedBy = [...html.matchAll(/<link[^>]*rel=["']describedby["'][^>]*>/gi)];
const alternateMd = [...html.matchAll(/<link[^>]*type=["']text\/markdown["'][^>]*>/gi)];
console.log('\n[ENG-07] LLMO Describedby Linki:', describedBy.length > 0 ? 'MEVCUT' : 'EKSİK');
console.log('[ENG-07] LLMO Alternate Markdown Linki:', alternateMd.length > 0 ? 'MEVCUT' : 'EKSİK');

// 5. C2PA & Provenance (ENG-03)
const c2paMeta = html.match(/<meta[^>]*name=["']c2pa-manifest["'][^>]*>/i);
const dctermsIssued = html.match(/<meta[^>]*name=["']dcterms\.issued["'][^>]*>/i);
console.log('\n[ENG-03] C2PA Manifest Meta:', c2paMeta ? 'MEVCUT' : 'EKSİK');
console.log('[ENG-03] DCTERMS Issued Meta:', dctermsIssued ? 'MEVCUT' : 'EKSİK');

// 6. RAG Chunks (ENG-09, ENG-10)
const chunkIds = [...html.matchAll(/data-chunk-id/gi)];
console.log('\n[ENG-09] RAG data-chunk-id Sayısı:', chunkIds.length);

// 7. AEO Hero Answer in top 100px (ENG-06)
const heroAnswer = html.match(/class=["'][^"']*hero-answer/i);
console.log('\n[ENG-06] AEO Hero Answer Sınıfı:', heroAnswer ? 'MEVCUT' : 'EKSİK');

// 8. JSON-LD Blocks (ENG-08 & ENG-15)
const jsonLdMatches = [...html.matchAll(/<script[^>]*type=["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/gi)];
console.log('\n[ENG-08 & ENG-15] Toplam JSON-LD Blok Sayısı:', jsonLdMatches.length);
let hasWikidata = false;
let hasMid = false;
jsonLdMatches.forEach((m, i) => {
  const content = m[1];
  if (content.includes('wikidata.org/wiki/Q')) hasWikidata = true;
  if (content.includes('kgmid') || content.includes('/m/')) hasMid = true;
  try {
    const parsed = JSON.parse(content);
    console.log('   Blok #' + (i+1) + ' Type:', parsed['@type'] || (parsed['@graph'] ? '@graph (' + parsed['@graph'].length + ' nodes)' : 'Unknown'));
  } catch (e) {
    console.log('   Blok #' + (i+1) + ' JSON Parse Hatası:', e.message);
  }
});
console.log('[ENG-15] Knowledge Vault Wikidata QID:', hasWikidata ? 'BAĞLI' : 'EKSİK (sameAs içinde Wikidata yok)');
console.log('[ENG-15] Knowledge Vault Google MID:', hasMid ? 'BAĞLI' : 'EKSİK');
