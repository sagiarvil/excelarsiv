// ensure_utf8_nobom.js
// Dosya yazıldıktan veya düzenlendikten sonra başında UTF-8 BOM (EF BB BF) varsa otomatik olarak siler.
// Evrensel Kural #6 ve #9 ile tam uyumlu BOM-suz saf UTF-8 garantisi sağlar.

'use strict';
const fs = require('fs');

let raw = '';
if (!process.stdin.isTTY) {
  try {
    raw = fs.readFileSync(0, 'utf8');
  } catch (e) {}
}

if (raw) {
  try {
    const payload = JSON.parse(raw);
    const args = (payload.toolCall && payload.toolCall.args) || {};
    const file = args.TargetFile || args.FilePath || args.AbsolutePath;
    if (file && fs.existsSync(file)) {
      const ext = file.split('.').pop().toLowerCase();
      const textExts = ['html', 'htm', 'js', 'css', 'php', 'json', 'md', 'sql', 'txt', 'xml'];
      if (textExts.includes(ext)) {
        const buf = fs.readFileSync(file);
        if (buf.length >= 3 && buf[0] === 0xEF && buf[1] === 0xBB && buf[2] === 0xBF) {
          fs.writeFileSync(file, buf.slice(3));
          console.log('[BOM TEMİZLENDİ] UTF-8 BOM kaldırıldı: ' + file);
        }
      }
    }
  } catch (e) {}
}
process.exit(0);
