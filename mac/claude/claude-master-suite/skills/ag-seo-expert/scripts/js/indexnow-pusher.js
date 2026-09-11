const https = require('https');

const host = process.argv[2];
const key = process.argv[3];
const urlListPath = process.argv[4];

if (!host || !key || !urlListPath) {
    console.log("Kullanım: node indexnow-pusher.js <host_adresi> <indexnow_key> <url_listesi_txt_yolu>");
    process.exit(1);
}

const fs = require('fs');
const urls = fs.readFileSync(urlListPath, 'utf8').split('\n').map(l => l.trim()).filter(l => l.length > 0);

const payload = JSON.stringify({
    host: host,
    key: key,
    keyLocation: `https://${host}/${key}.txt`,
    urlList: urls
});

const options = {
    hostname: 'api.indexnow.org',
    path: '/indexnow',
    method: 'POST',
    headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(payload)
    }
};

const req = https.request(options, (res) => {
    console.log(`\n=== SÜPER SEO MOTORU (INDEXNOW PUSHER V3.0) ===`);
    console.log(`[+] API Yanıt Kodu: ${res.statusCode}`);
    if (res.statusCode === 200 || res.statusCode === 202) {
        console.log(`[+] BAŞARILI: ${urls.length} URL arama motorlarına iletildi.`);
    } else {
        console.log(`[X] HATA: Gönderim başarısız.`);
    }
});

req.on('error', (e) => {
    console.error(`[X] BAĞLANTI HATASI: ${e.message}`);
});

req.write(payload);
req.end();
