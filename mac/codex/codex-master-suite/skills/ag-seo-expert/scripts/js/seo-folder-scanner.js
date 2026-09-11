const fs = require('fs');
const path = require('path');

const targetDir = process.argv[2];

if (!targetDir) {
    console.log("Kullanım: node seo-folder-scanner.js <klasör_yolu>");
    process.exit(1);
}

function scanDirectory(dir) {
    let results = [];
    const files = fs.readdirSync(dir);
    
    for (const file of files) {
        const fullPath = path.join(dir, file);
        const stat = fs.statSync(fullPath);
        
        if (stat.isDirectory()) {
            results = results.concat(scanDirectory(fullPath));
        } else if (fullPath.endsWith('.html') || fullPath.endsWith('.htm')) {
            results.push(analyzeFile(fullPath, stat.size));
        }
    }
    return results;
}

function analyzeFile(filePath, byteSize) {
    const content = fs.readFileSync(filePath, 'utf8');
    
    // Vektör 1: Core Technical & Structural SEO
    const h1Count = (content.match(/<h1[\s>]/gi) || []).length;
    const hasCanonical = /<link[^>]*rel=["']canonical["'][^>]*>/i.test(content);
    const hasNoIndex = /<meta[^>]+name=["']robots["'][^>]+content=["'][^"']*noindex/i.test(content);
    
    // Vektör 4: LLMO (Large Language Model Optimization)
    const hasLlmsDiscovery = /rel=["']describedby["'][^>]+href=["']\/llms\.txt["']/i.test(content);
    const hasMarkdownAlternate = /rel=["']alternate["'][^>]+type=["']text\/markdown["']/i.test(content);
    
    // Vektör 6: RAG (Retrieval-Augmented Generation) & Dense Retrieval
    const hasDataChunkId = /data-chunk-id/i.test(content);
    const h2h3Count = (content.match(/<(h2|h3)[\s>]/gi) || []).length;
    
    // Vektör 7: E-E-A-T & Knowledge Vault
    const hasWikidata = /wikidata\.org\/wiki\/Q/i.test(content);
    
    return {
        file: filePath,
        size: byteSize,
        h1: h1Count,
        canonical: hasCanonical,
        noindex: hasNoIndex,
        llmsDiscovery: hasLlmsDiscovery,
        markdownAlt: hasMarkdownAlternate,
        dataChunkId: hasDataChunkId,
        h2h3: h2h3Count,
        wikidata: hasWikidata
    };
}

console.log(`\n=== SÜPER SEO MOTORU V3.0 (KAPSAMLI KLASÖR TARAMASI) ===`);
console.log(`[+] Hedef Dizin: ${path.resolve(targetDir)}\n`);

try {
    const report = scanDirectory(targetDir);
    
    if (report.length === 0) {
        console.log("[!] Belirtilen klasörde HTML dosyası bulunamadı.");
        process.exit(0);
    }
    
    report.forEach(res => {
        console.log(`📄 Dosya: ${res.file}`);
        
        // Vektör 2: 14KB Sınırı
        if (res.size > 14336) {
            console.log(`   [X] HATA (TOKEN-BLOAT-001): Boyut ${res.size} byte (14KB aşıldı!)`);
        } else {
            console.log(`   [V] PASS: Boyut ${res.size} byte (Uygun)`);
        }
        
        // Vektör 1: H1 ve Canonical
        console.log(res.h1 === 1 ? `   [V] PASS: Tek H1 etiketi (TECH-H1-001)` : `   [X] HATA: H1 Sayısı ${res.h1} (Sadece 1 adet olmalı)`);
        console.log(res.canonical ? `   [V] PASS: Canonical var (TECH-CANON-001)` : `   [X] HATA: Canonical YOK`);
        if (res.noindex) console.log(`   [X] KRİTİK HATA (TECH-NOINDEX-001): Accidental noindex tespit edildi!`);
        
        // Vektör 4: LLMO
        console.log(res.llmsDiscovery ? `   [V] PASS: llms.txt discovery (LLMS-DISCOVERY-001)` : `   [!] UYARI: describedby=llms.txt YOK`);
        
        // Vektör 6: RAG
        console.log(res.dataChunkId ? `   [V] PASS: Semantic boundaries data-chunk-id (RAG-CHUNK-001)` : `   [!] UYARI: data-chunk-id bulunamadı (RAG kopmaları riski)`);
        console.log(res.h2h3 >= 3 ? `   [V] PASS: H2/H3 Matrisi (${res.h2h3} adet) (COLBERT-MAXSIM-001)` : `   [!] UYARI: H2/H3 yapısı zayıf (${res.h2h3} adet)`);
        
        // Vektör 7: Knowledge Vault
        console.log(res.wikidata ? `   [V] PASS: Wikidata JSON-LD/sameAs referansı (ENTITY-VAULT-001)` : `   [!] UYARI: Wikidata (QID) entity referansı eksik`);
        
        console.log("--------------------------------------------------");
    });
    
} catch (err) {
    console.error(`[X] HATA: ${err.message}`);
}
