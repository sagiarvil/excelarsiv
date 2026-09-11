const https = require('https');
const http = require('http');

const url = process.argv[2];

if (!url) {
    console.log("Kullanım: node seo-checker.js <url>");
    process.exit(1);
}

const client = url.startsWith('https') ? https : http;

client.get(url, (res) => {
    let data = '';
    
    res.on('data', chunk => {
        data += chunk;
    });
    
    res.on('end', () => {
        console.log(`\n=== SÜPER SEO MOTORU (JS V3.0) ===`);
        console.log(`[+] Hedef URL: ${url}`);
        
        // 14KB AST Kontrolü (14.336 byte)
        const byteSize = Buffer.byteLength(data, 'utf8');
        console.log(`[+] Sayfa Boyutu: ${byteSize} byte`);
        if (byteSize > 14336) {
            console.log(`[!] UYARI (TOKEN-BLOAT-001): 14KB AST bütçesi aşıldı! AI motorları veriyi kesebilir.`);
        } else {
            console.log(`[+] BAŞARILI: 14KB AST bütçesi sınırları içerisinde.`);
        }
        
        // H1 Etiketi Kontrolü
        const h1Matches = data.match(/<h1[\s>]/gi) || [];
        console.log(`[+] H1 Etiketi Sayısı: ${h1Matches.length}`);
        if (h1Matches.length !== 1) {
            console.log(`[!] UYARI (TECH-H1-001): Sayfada tam olarak 1 adet H1 olmalıdır.`);
        }
        
        // Canonical Kontrolü
        const hasCanonical = /<link[^>]*rel=["']canonical["'][^>]*>/i.test(data);
        console.log(`[+] Canonical Belirteci: ${hasCanonical ? 'VAR' : 'YOK'}`);
        if (!hasCanonical) {
            console.log(`[!] UYARI (TECH-CANON-001): Canonical etiketi bulunamadı!`);
        }
        
        console.log(`==================================\n`);
    });
}).on('error', (err) => {
    console.error(`[X] HATA: ${err.message}`);
});
